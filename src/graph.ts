import { PublicClientApplication } from '@azure/msal-browser'
import type { AccountInfo, AuthenticationResult } from '@azure/msal-browser'

export const graphScopes = ['User.Read', 'Calendars.ReadWrite', 'Files.Read']

const clientId = import.meta.env.VITE_ENTRA_CLIENT_ID?.trim() ?? ''
const tenantId = import.meta.env.VITE_ENTRA_TENANT_ID?.trim() ?? ''
const redirectUri = import.meta.env.VITE_ENTRA_REDIRECT_URI?.trim() || window.location.origin

export const missingConfiguration = [
  !clientId && 'VITE_ENTRA_CLIENT_ID',
  !tenantId && 'VITE_ENTRA_TENANT_ID',
].filter((value): value is string => Boolean(value))

export const msal = missingConfiguration.length
  ? null
  : new PublicClientApplication({
      auth: {
        clientId,
        authority: `https://login.microsoftonline.com/${tenantId}`,
        redirectUri,
      },
      cache: {
        cacheLocation: 'localStorage',
      },
    })

export interface GraphEvent {
  id: string
  subject?: string
  start: { dateTime: string; timeZone: string }
  end: { dateTime: string; timeZone: string }
  isOnlineMeeting?: boolean
  onlineMeetingProvider?: string | null
  onlineMeeting?: { joinUrl?: string }
  webLink?: string
  attendees?: Array<{ emailAddress?: { address?: string; name?: string } }>
}

export interface DriveRecording {
  id: string
  name: string
  size?: number
  createdDateTime?: string
  lastModifiedDateTime?: string
  webUrl?: string
  file?: { mimeType?: string }
  folder?: { childCount?: number }
  '@microsoft.graph.downloadUrl'?: string
}

export interface GraphUser {
  displayName?: string
  mail?: string
  userPrincipalName?: string
}

export class MicrosoftGraphError extends Error {
  constructor(message: string, readonly status: number, readonly code?: string) {
    super(message)
    this.name = 'MicrosoftGraphError'
  }
}

async function getAccessToken(
  client: PublicClientApplication,
  account: AccountInfo,
  scopes: string[] = graphScopes,
  allowInteractive = true,
): Promise<string> {
  try {
    const result = await client.acquireTokenSilent({ scopes, account })
    return result.accessToken
  } catch {
    if (!allowInteractive) {
      throw new Error('Microsoft necesita autorización adicional. Actualiza las grabaciones y acepta el permiso solicitado.')
    }
    const result: AuthenticationResult = await client.acquireTokenPopup({
      scopes,
      account,
    })
    return result.accessToken
  }
}

export async function requestRecordingShareAccess(
  client: PublicClientApplication,
  account: AccountInfo,
): Promise<void> {
  await getAccessToken(client, account, ['Files.ReadWrite'])
}

export async function graphRequest<T>(
  client: PublicClientApplication,
  account: AccountInfo,
  path: string,
  init: RequestInit = {},
  scopes: string[] = graphScopes,
  allowInteractive = true,
): Promise<T> {
  const token = await getAccessToken(client, account, scopes, allowInteractive)
  const response = await fetch(path.startsWith('https://') ? path : `https://graph.microsoft.com/v1.0${path}`, {
    ...init,
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: 'application/json',
      ...(init.body ? { 'Content-Type': 'application/json' } : {}),
      ...(init.headers ?? {}),
      Prefer: 'outlook.timezone="UTC"',
    },
  })

  if (!response.ok) {
    const details = await response.json().catch(() => null) as
      | { error?: { message?: string; code?: string } }
      | null
    const message = details?.error?.message || `Microsoft Graph respondió ${response.status}.`
    throw new MicrosoftGraphError(message, response.status, details?.error?.code)
  }
  if (response.status === 204) return undefined as T
  return response.json() as Promise<T>
}

export async function fetchRecordings(
  client: PublicClientApplication,
  account: AccountInfo,
): Promise<DriveRecording[]> {
  const rootParams = new URLSearchParams({
    '$select': 'id,name,folder',
    '$top': '200',
  })
  let rootUrl: string | null = `/me/drive/root/children?${rootParams.toString()}`
  let recordingsFolderId: string | null = null
  while (rootUrl && !recordingsFolderId) {
    const page: { value: DriveRecording[]; '@odata.nextLink'?: string } =
      await graphRequest(client, account, rootUrl)
    const folder = page.value.find((item) =>
      Boolean(item.folder)
      && ['recordings', 'grabaciones'].includes(item.name.trim().toLocaleLowerCase()),
    )
    recordingsFolderId = folder?.id ?? null
    rootUrl = page['@odata.nextLink'] ?? null
  }
  if (!recordingsFolderId) {
    throw new Error('No se encontró la carpeta Recordings o Grabaciones en la raíz de OneDrive.')
  }

  const params = new URLSearchParams({
    '$select': 'id,name,size,createdDateTime,lastModifiedDateTime,webUrl,file,folder',
    '$orderby': 'lastModifiedDateTime desc',
    '$top': '100',
  })
  const items: DriveRecording[] = []
  let url: string | null = `/me/drive/items/${encodeURIComponent(recordingsFolderId)}/children?${params.toString()}`
  while (url) {
    const page: { value: DriveRecording[]; '@odata.nextLink'?: string } =
      await graphRequest(client, account, url)
    items.push(...page.value.filter((item) => Boolean(item.file)))
    url = page['@odata.nextLink'] ?? null
  }
  return items
}

export async function createOrganizationRecordingLink(
  client: PublicClientApplication,
  account: AccountInfo,
  driveItemId: string,
  allowInteractive = true,
): Promise<string> {
  const result = await graphRequest<{
    link?: { webUrl?: string; scope?: string; type?: string }
  }>(client, account, `/me/drive/items/${encodeURIComponent(driveItemId)}/createLink`, {
    method: 'POST',
    body: JSON.stringify({ type: 'view', scope: 'organization' }),
  }, ['Files.ReadWrite'], allowInteractive)
  const webUrl = result.link?.webUrl
  if (!webUrl || result.link?.scope !== 'organization') {
    throw new Error('Microsoft no devolvió un enlace institucional válido para la grabación.')
  }
  const url = new URL(webUrl)
  if (url.protocol !== 'https:'
    || url.username
    || url.password
    || !(url.hostname.endsWith('.sharepoint.com')
      || url.hostname === '1drv.ms'
      || url.hostname === 'onedrive.live.com')) {
    throw new Error('Microsoft devolvió un enlace de grabación no válido.')
  }
  return url.toString()
}

export async function fetchSignedInUser(
  client: PublicClientApplication,
  account: AccountInfo,
): Promise<GraphUser> {
  return graphRequest<GraphUser>(client, account, '/me?$select=displayName,mail,userPrincipalName')
}

export async function createTeamsEvent(
  client: PublicClientApplication,
  account: AccountInfo,
  data: {
    subject: string
    start: string
    end: string
    timeZone: string
    attendees: string[]
    description?: string
    location?: string
  },
): Promise<GraphEvent> {
  return graphRequest<GraphEvent>(client, account, '/me/events', {
    method: 'POST',
    body: JSON.stringify({
      subject: data.subject,
      start: { dateTime: data.start, timeZone: data.timeZone },
      end: { dateTime: data.end, timeZone: data.timeZone },
      isOnlineMeeting: true,
      onlineMeetingProvider: 'teamsForBusiness',
      body: data.description
        ? { contentType: 'text', content: data.description }
        : undefined,
      location: data.location ? { displayName: data.location } : undefined,
      attendees: data.attendees.map((address) => ({
        emailAddress: { address },
        type: 'required',
      })),
    }),
  })
}

export async function deleteTeamsEvent(
  client: PublicClientApplication,
  account: AccountInfo,
  eventId: string,
): Promise<void> {
  await graphRequest<void>(client, account, `/me/events/${encodeURIComponent(eventId)}`, {
    method: 'DELETE',
  })
}

export async function waitForTeamsJoinLink(
  client: PublicClientApplication,
  account: AccountInfo,
  event: GraphEvent,
): Promise<GraphEvent> {
  if (!event.id) throw new Error('Microsoft no devolvió el identificador de la reunión creada.')
  for (let attempt = 0; attempt < 6; attempt++) {
    const current = event.onlineMeeting?.joinUrl
      ? event
      : await graphRequest<GraphEvent>(
          client,
          account,
          `/me/events/${encodeURIComponent(event.id)}?$select=id,subject,start,end,isOnlineMeeting,onlineMeetingProvider,onlineMeeting,webLink`,
        )
    if (current.isOnlineMeeting === true
      && current.onlineMeetingProvider === 'teamsForBusiness'
      && typeof current.onlineMeeting?.joinUrl === 'string') {
      return current
    }
    if (attempt < 5) await new Promise((resolve) => window.setTimeout(resolve, 1000))
  }
  throw new Error('Microsoft creó el evento, pero todavía no generó el enlace de Teams. Reintenta la publicación en unos segundos.')
}
