import { useEffect, useMemo, useState } from 'react'
import type { FormEvent, ReactNode } from 'react'
import {
  AlertCircle, ArrowRight, CalendarDays, Check, ChevronLeft,
  ChevronRight, Clock3, FileVideo2, FolderOpen, LoaderCircle, LogOut, Mail,
  Pencil, Play, Plus, RotateCw, ShieldCheck, Trash2, Video, X,
} from 'lucide-react'
import type { AccountInfo, AuthenticationResult } from '@azure/msal-browser'
import {
  createOrganizationRecordingLink, createTeamsEvent, fetchRecordings, fetchSignedInUser, graphScopes,
  deleteTeamsEvent, MicrosoftGraphError, requestRecordingShareAccess,
  waitForTeamsJoinLink,
  msal,
} from './graph'
import type { DriveRecording, GraphEvent, GraphUser } from './graph'

type View = 'schedule' | 'recordings' | 'summary' | 'notes'
type MeetingPeriod = 'upcoming' | 'past' | 'personal'
type DraftStatus = 'review' | 'ready' | 'created'
type SessionDraft = {
  id: string
  title: string
  date: string
  start: string
  end: string
  timeZone: string
  attendees: string[]
  description: string
  location: string
  recurrence: boolean
  status: DraftStatus
  graphEvent?: GraphEvent
}
type CourseContext = {
  contextId: string
  courseName?: string
  courseCode?: string
  term?: string
  academicPeriod?: string
  nrc?: string
  instructorName?: string
}
type TrustedRole = 'instructor' | 'learner'
type LearnerSession = {
  id: string
  start: string
  end: string
  timeZone: string
  topic?: string
  joinUrl?: string
}
type CourseSession = {
  id: string
  graphEventId: string
  start: string
  end: string
  timeZone: string
  topic: string
  joinUrl: string
}
type CourseRecording = {
  id: string
  sessionId?: string
  driveItemId?: string
  title: string
  createdAt?: string
  webUrl: string
}
type FormMode = 'weekly' | 'single'

const weekdays = [
  { short: 'Lun', label: 'Lunes', value: 1 },
  { short: 'Mar', label: 'Martes', value: 2 },
  { short: 'Mié', label: 'Miércoles', value: 3 },
  { short: 'Jue', label: 'Jueves', value: 4 },
  { short: 'Vie', label: 'Viernes', value: 5 },
  { short: 'Sáb', label: 'Sábado', value: 6 },
]
const timeZones = [
  'Central Standard Time (Mexico)', 'Pacific Standard Time',
  'Mountain Standard Time', 'Central Standard Time', 'Eastern Standard Time', 'UTC',
]
const displayTimeZones: Record<string, string> = {
  'Central Standard Time (Mexico)': 'America/Mexico_City',
  'Pacific Standard Time': 'America/Los_Angeles',
  'Mountain Standard Time': 'America/Denver',
  'Central Standard Time': 'America/Chicago',
  'Eastern Standard Time': 'America/New_York',
  UTC: 'Etc/UTC',
}
const hours = Array.from({ length: 15 }, (_, index) => index + 7)
const today = new Date()
const localISODate = (date: Date) => {
  const adjusted = new Date(date.getTime() - date.getTimezoneOffset() * 60_000)
  return adjusted.toISOString().slice(0, 10)
}
const addDays = (date: Date, count: number) => {
  const result = new Date(date)
  result.setDate(result.getDate() + count)
  return result
}
const dayStart = (date: Date) => {
  const result = new Date(date)
  result.setHours(0, 0, 0, 0)
  result.setDate(result.getDate() - ((result.getDay() + 6) % 7))
  return result
}
const timeLabel = (value: string) => value.slice(0, 5)
const formatDay = (date: Date, options: Intl.DateTimeFormatOptions = {}) =>
  new Intl.DateTimeFormat('es-MX', { day: 'numeric', month: 'short', ...options }).format(date)
const formatLongDay = (date: Date) =>
  new Intl.DateTimeFormat('es-MX', { weekday: 'long', day: 'numeric', month: 'long' }).format(date)
const graphDate = (value: string) =>
  new Date(/[zZ]|[+-]\d\d:\d\d$/.test(value) ? value : `${value}Z`)
const graphLocalDate = (value: string) => {
  const date = graphDate(value)
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`
}
const graphLocalTime = (value: string) =>
  new Intl.DateTimeFormat('en-GB', { hour: '2-digit', minute: '2-digit' }).format(graphDate(value))
const isTeamsOnlineEvent = (event: GraphEvent) => {
  if (event.isOnlineMeeting !== true || event.onlineMeetingProvider !== 'teamsForBusiness') return false
  try {
    return new URL(event.onlineMeeting?.joinUrl ?? '').protocol === 'https:'
  } catch {
    return false
  }
}
const initials = (name?: string) => (name || 'Cuenta').split(/\s+/).slice(0, 2).map((part) => part[0]).join('').toUpperCase()
const minutesOf = (value: string) => {
  const [hour, minute] = value.split(':').map(Number)
  return hour * 60 + minute
}
function parseCourseSession(value: unknown): CourseSession | null {
  if (!isObject(value) || typeof value.id !== 'string' || typeof value.graphEventId !== 'string'
    || typeof value.start !== 'string' || !Number.isFinite(Date.parse(value.start))
    || typeof value.end !== 'string' || !Number.isFinite(Date.parse(value.end))
    || Date.parse(value.end) <= Date.parse(value.start)
    || typeof value.timeZone !== 'string' || typeof value.topic !== 'string') return null
  const joinUrl = secureWebUrl(value.joinUrl)
  if (!joinUrl) return null
  return {
    id: value.id,
    graphEventId: value.graphEventId,
    start: value.start,
    end: value.end,
    timeZone: value.timeZone,
    topic: value.topic,
    joinUrl,
  }
}

async function fetchCourseSessions(): Promise<CourseSession[]> {
  const response = await fetch('/api/course/sessions', {
    credentials: 'same-origin',
    headers: { Accept: 'application/json' },
  })
  if (!response.ok) throw new Error(`No se pudieron consultar las reuniones de este curso (HTTP ${response.status}).`)
  const data: unknown = await response.json()
  if (!isObject(data) || !Array.isArray(data.sessions)) {
    throw new Error('La respuesta de reuniones del curso no tiene el formato esperado.')
  }
  const sessions = data.sessions.map(parseCourseSession)
  if (!sessions.every((session): session is CourseSession => session !== null)) {
    throw new Error('El curso contiene reuniones publicadas con datos incompletos o no válidos.')
  }
  return sessions.sort((a, b) => Date.parse(a.start) - Date.parse(b.start))
}

async function deleteCourseSession(sessionId: string): Promise<void> {
  const response = await fetch(`/api/course/sessions/${encodeURIComponent(sessionId)}`, {
    method: 'DELETE',
    credentials: 'same-origin',
    headers: { Accept: 'application/json' },
  })
  if (response.status === 404) return
  if (!response.ok) {
    let message = `No se pudo eliminar la sesión del curso (HTTP ${response.status}).`
    try {
      const data: unknown = await response.json()
      if (isObject(data) && typeof data.error === 'string') message = data.error
    } catch {
      // Keep the HTTP error when the server does not return JSON.
    }
    throw new Error(message)
  }
}

const pendingCourseSessionDeletions = new Set<string>()

function courseSessionDeletionKey(contextId: string, sessionId: string): string {
  return `teams.pending-course-session-delete:${encodeURIComponent(contextId)}:${sessionId}`
}

function microsoftAccountBindingKey(contextId: string): string {
  return `teams.microsoft-account:${encodeURIComponent(contextId)}`
}

async function fetchCourseRecordings(): Promise<CourseRecording[]> {
  const response = await fetch('/api/course/recordings', {
    credentials: 'same-origin',
    headers: { Accept: 'application/json' },
  })
  if (!response.ok) throw new Error(`No se pudieron consultar las grabaciones de este curso (HTTP ${response.status}).`)
  const data: unknown = await response.json()
  if (!isObject(data) || !Array.isArray(data.recordings)) {
    throw new Error('La respuesta de grabaciones del curso no tiene el formato esperado.')
  }
  const recordings = data.recordings.flatMap((raw): CourseRecording[] => {
    if (!isObject(raw) || typeof raw.id !== 'string' || typeof raw.title !== 'string') return []
    const webUrl = secureWebUrl(raw.webUrl)
    if (!webUrl) return []
    return [{
      id: raw.id,
      sessionId: typeof raw.sessionId === 'string' ? raw.sessionId : undefined,
      driveItemId: typeof raw.driveItemId === 'string' ? raw.driveItemId : undefined,
      title: raw.title,
      createdAt: typeof raw.createdAt === 'string' && Number.isFinite(Date.parse(raw.createdAt)) ? raw.createdAt : undefined,
      webUrl,
    }]
  })
  if (recordings.length !== data.recordings.length) {
    throw new Error('El curso contiene grabaciones incompletas o no válidas.')
  }
  return recordings
}

async function publishCourseRecording(
  session: CourseSession,
  recording: DriveRecording,
  webUrl: string,
): Promise<CourseRecording> {
  if (!recording.createdDateTime) throw new Error('Microsoft no devolvió la fecha de la grabación.')
  const response = await fetch('/api/course/recordings', {
    method: 'POST',
    credentials: 'same-origin',
    headers: { Accept: 'application/json', 'Content-Type': 'application/json' },
    body: JSON.stringify({
      sessionId: session.id,
      driveItemId: recording.id,
      title: recording.name,
      createdAt: recording.createdDateTime,
      webUrl,
    }),
  })
  if (!response.ok) throw new Error(`No se pudo publicar la grabación para el curso (HTTP ${response.status}).`)
  const data: unknown = await response.json()
  if (!isObject(data) || !isObject(data.recording)) {
    throw new Error('La respuesta al publicar la grabación no tiene el formato esperado.')
  }
  const published = data.recording
  const safeUrl = secureWebUrl(published.webUrl)
  if (typeof published.id !== 'string' || !safeUrl || typeof published.title !== 'string') {
    throw new Error('El servidor devolvió una grabación incompleta o no válida.')
  }
  return {
    id: published.id,
    sessionId: typeof published.sessionId === 'string' ? published.sessionId : session.id,
    driveItemId: typeof published.driveItemId === 'string' ? published.driveItemId : recording.id,
    title: published.title,
    createdAt: typeof published.createdAt === 'string' ? published.createdAt : recording.createdDateTime,
    webUrl: safeUrl,
  }
}

function recordingSessionMatch(
  recording: DriveRecording,
  sessions: CourseSession[],
  courseCode: string,
): CourseSession | null {
  const videoFile = recording.file?.mimeType?.startsWith('video/')
    || /\.(mp4|mov|m4v|webm)$/i.test(recording.name)
  const createdAt = recording.createdDateTime ? Date.parse(recording.createdDateTime) : Number.NaN
  if (!videoFile || !Number.isFinite(createdAt)) return null
  const filenameTimestamp = recording.name.match(/(?:^|[^\d])(\d{4})(\d{2})(\d{2})_(\d{2})(\d{2})(\d{2})(?:[^\d]|$)/)
  if (filenameTimestamp) {
    const recordingWallTime = Date.UTC(
      Number(filenameTimestamp[1]),
      Number(filenameTimestamp[2]) - 1,
      Number(filenameTimestamp[3]),
      Number(filenameTimestamp[4]),
      Number(filenameTimestamp[5]),
      Number(filenameTimestamp[6]),
    )
    const candidates = sessions.flatMap((session) => {
      const timeZone = displayTimeZones[session.timeZone] ?? session.timeZone
      const toWallTime = (value: string) => {
        const parts = new Intl.DateTimeFormat('en-GB', {
          timeZone,
          year: 'numeric',
          month: '2-digit',
          day: '2-digit',
          hour: '2-digit',
          minute: '2-digit',
          second: '2-digit',
          hourCycle: 'h23',
        }).formatToParts(new Date(value))
        const component = (type: Intl.DateTimeFormatPartTypes) =>
          Number(parts.find((part) => part.type === type)?.value)
        return Date.UTC(
          component('year'),
          component('month') - 1,
          component('day'),
          component('hour'),
          component('minute'),
          component('second'),
        )
      }
      const start = toWallTime(session.start)
      const end = toWallTime(session.end)
      if (recordingWallTime < start - 15 * 60_000 || recordingWallTime > end + 60 * 60_000) return []
      const distance = recordingWallTime < start
        ? start - recordingWallTime
        : recordingWallTime > end ? recordingWallTime - end : 0
      return [{ session, distance }]
    }).sort((a, b) => a.distance - b.distance)
    if (candidates.length === 1
      || (candidates.length > 1 && candidates[1].distance - candidates[0].distance > 15 * 60_000)) {
      return candidates[0]?.session ?? null
    }
    return null
  }

  const normalizedName = recording.name.normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '')
  const normalizedCourseCode = courseCode.normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '')
  if (normalizedCourseCode.length < 3 || !normalizedName.includes(normalizedCourseCode)) return null
  const candidates = sessions.flatMap((session) => {
    const endedAt = Date.parse(session.end)
    const normalizedTopic = session.topic.normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .toLowerCase()
      .replace(/[^a-z0-9]/g, '')
    const minutesSinceEnd = (createdAt - endedAt) / 60_000
    if (normalizedTopic.length < 5
      || !normalizedName.includes(normalizedTopic)
      || minutesSinceEnd < -10
      || minutesSinceEnd > 48 * 60) return []
    return [{ session, distance: Math.abs(minutesSinceEnd) }]
  })
  candidates.sort((a, b) => a.distance - b.distance)
  return candidates[0]?.session ?? null
}

async function syncCourseRecordings(
  client: NonNullable<typeof msal>,
  account: AccountInfo,
  courseCode: string,
  allowInteractive: boolean,
): Promise<CourseRecording[]> {
  const [sessions, savedRecordings, driveItems] = await Promise.all([
    fetchCourseSessions(),
    fetchCourseRecordings(),
    fetchRecordings(client, account),
  ])
  const knownDriveIds = new Set(savedRecordings.flatMap((item) => item.driveItemId ? [item.driveItemId] : []))
  const claimedDriveIds = new Set<string>()
  const result = [...savedRecordings]
  for (const item of driveItems) {
    if (!item.id || knownDriveIds.has(item.id) || claimedDriveIds.has(item.id)) continue
    const session = recordingSessionMatch(item, sessions, courseCode)
    if (!session) continue
    const shareUrl = await createOrganizationRecordingLink(client, account, item.id, allowInteractive)
    const published = await publishCourseRecording(session, item, shareUrl)
    result.push(published)
    claimedDriveIds.add(item.id)
  }
  return result.sort((a, b) => (b.createdAt ?? '').localeCompare(a.createdAt ?? ''))
}

function courseSessionEvent(session: CourseSession): GraphEvent {
  return {
    id: session.graphEventId,
    subject: session.topic,
    start: { dateTime: session.start, timeZone: session.timeZone },
    end: { dateTime: session.end, timeZone: session.timeZone },
    isOnlineMeeting: true,
    onlineMeetingProvider: 'teamsForBusiness',
    onlineMeeting: { joinUrl: session.joinUrl },
    webLink: session.joinUrl,
  }
}

function App() {
  const [view, setView] = useState<View>('schedule')
  const [advancedToolsOpen, setAdvancedToolsOpen] = useState(false)
  const [teacherPeriod, setTeacherPeriod] = useState<MeetingPeriod>('upcoming')
  const [teacherOverflowOpen, setTeacherOverflowOpen] = useState(false)
  const [account, setAccount] = useState<AccountInfo | null>(null)
  const [user, setUser] = useState<GraphUser | null>(null)
  const [course, setCourse] = useState<CourseContext | null>(null)
  const [role, setRole] = useState<TrustedRole | null>(null)
  const [ltiState, setLtiState] = useState<'loading' | 'unavailable' | 'active' | 'forbidden'>('loading')
  const [courseSessions, setCourseSessions] = useState<CourseSession[]>([])
  const [courseSessionsError, setCourseSessionsError] = useState('')
  const [courseSessionsLoading, setCourseSessionsLoading] = useState(false)
  const [deletingSessionId, setDeletingSessionId] = useState('')
  const [recordings, setRecordings] = useState<CourseRecording[]>([])
  const [courseRecordingLoading, setCourseRecordingLoading] = useState(false)
  const [authReady, setAuthReady] = useState(!msal)
  const [loading, setLoading] = useState(false)
  const [publishingBatch, setPublishingBatch] = useState(false)
  const [signingIn, setSigningIn] = useState(false)
  const [error, setError] = useState('')
  const [reviewError, setReviewError] = useState('')
  const [recordingError, setRecordingError] = useState('')
  const [notice, setNotice] = useState('')
  const [drafts, setDrafts] = useState<SessionDraft[]>([])
  const [week, setWeek] = useState(() => dayStart(today))
  const [mobileDay, setMobileDay] = useState(Math.min((today.getDay() + 6) % 7, 5))
  const [modal, setModal] = useState<FormMode | null>(null)
  const [reviewing, setReviewing] = useState<SessionDraft | null>(null)
  const [saving, setSaving] = useState(false)
  const [editingId, setEditingId] = useState<string | null>(null)
  const [formTitle, setFormTitle] = useState('')
  const [startTime, setStartTime] = useState('09:00')
  const [endTime, setEndTime] = useState('10:00')
  const [formDate, setFormDate] = useState(localISODate(addDays(today, 1)))
  const [termEnd, setTermEnd] = useState(localISODate(addDays(today, 60)))
  const [selectedDays, setSelectedDays] = useState<number[]>([1])
  const [weeklyDaysInitialized, setWeeklyDaysInitialized] = useState(false)
  const [timeZone, setTimeZone] = useState(timeZones[0])
  const [attendees, setAttendees] = useState('')
  const [description, setDescription] = useState('')
  const [location, setLocation] = useState('')
  const [notesOpen, setNotesOpen] = useState(false)
  const [notes, setNotes] = useState('')

  useEffect(() => {
    let alive = true
    fetch('/api/lti/session', { credentials: 'same-origin', headers: { Accept: 'application/json' } })
      .then(async (response) => {
        if (!response.ok) throw new Error('LTI session unavailable')
        return response.json() as Promise<{ authenticated?: boolean; role?: string; course?: CourseContext }>
      })
      .then((result) => {
        if (!alive) return
        if (result.authenticated === true
          && (result.role === 'instructor' || result.role === 'learner')
          && result.course?.contextId?.trim()
          && result.course?.courseName
          && result.course.courseCode
          && result.course.nrc) {
          setCourse(result.course)
          setRole(result.role)
          setLtiState('active')
        } else if (result.authenticated === true) {
          setLtiState('forbidden')
        } else setLtiState('unavailable')
      })
      .catch(() => { if (alive) setLtiState('unavailable') })
    return () => { alive = false }
  }, [])

  useEffect(() => {
    let alive = true
    async function initialize() {
      if (ltiState === 'loading') return
      if (role !== 'instructor' || ltiState !== 'active') {
        setAuthReady(true)
        return
      }
      if (!msal) return
      if (course?.contextId) {
        setAccount(null)
        setUser(null)
        setRecordings([])
      }
      try {
        await msal.initialize()
        const result: AuthenticationResult | null = await msal.handleRedirectPromise()
        const accountKey = microsoftAccountBindingKey(course?.contextId ?? '')
        let selected = result?.account ?? null
        if (selected) {
          if (!course?.contextId) throw new Error('No se pudo identificar el curso para conectar la cuenta Microsoft.')
          window.localStorage.setItem(microsoftAccountBindingKey(course.contextId), selected.homeAccountId)
          msal.setActiveAccount(selected)
          if (alive) setAccount(selected)
        } else if (course?.contextId) {
          const savedAccountId = window.localStorage.getItem(accountKey)
          selected = savedAccountId
            ? msal.getAllAccounts().find((candidate) => candidate.homeAccountId === savedAccountId) ?? null
            : null
          if (selected) {
            msal.setActiveAccount(selected)
            if (alive) setAccount(selected)
          } else if (savedAccountId) {
            window.localStorage.removeItem(accountKey)
          }
        }
        if (new URLSearchParams(window.location.search).get('microsoftLogin') === '1') {
          window.history.replaceState(null, document.title, `${window.location.pathname}${window.location.hash}`)
          await msal.loginRedirect({ scopes: graphScopes, prompt: 'select_account' })
        }
      } catch (cause) {
        if (alive) setError(cause instanceof Error ? cause.message : 'No se pudo iniciar sesión.')
      } finally {
        if (alive) setAuthReady(true)
      }
    }
    void initialize()
    return () => { alive = false }
  }, [ltiState, role, course?.contextId])

  useEffect(() => {
    if (!msal || role !== 'instructor' || ltiState !== 'active' || !course?.contextId) return
    const client = msal
    const syncAccount = () => {
      const accountKey = microsoftAccountBindingKey(course.contextId)
      const savedAccountId = window.localStorage.getItem(accountKey)
      const selected = savedAccountId
        ? client.getAllAccounts().find((candidate) => candidate.homeAccountId === savedAccountId) ?? null
        : null
      if (!selected) {
        setAccount(null)
        setUser(null)
        return
      }
      client.setActiveAccount(selected)
      setAccount(selected)
    }
    window.addEventListener('focus', syncAccount)
    document.addEventListener('visibilitychange', syncAccount)
    return () => {
      window.removeEventListener('focus', syncAccount)
      document.removeEventListener('visibilitychange', syncAccount)
    }
  }, [course?.contextId, ltiState, role])

  useEffect(() => {
    if (role !== 'instructor' || ltiState !== 'active') return
    let alive = true
    setCourseSessionsLoading(true)
    setCourseSessionsError('')
    void fetchCourseSessions()
      .then((sessions) => {
        if (alive) setCourseSessions(sessions)
      })
      .catch((cause: unknown) => {
        if (alive) setCourseSessionsError(cause instanceof Error
          ? cause.message : 'No se pudieron consultar las reuniones de este curso.')
      })
      .finally(() => {
        if (alive) setCourseSessionsLoading(false)
      })
    return () => { alive = false }
  }, [role, ltiState])

  useEffect(() => {
    let alive = true
    async function load() {
      if (!msal || !account || role !== 'instructor' || ltiState !== 'active') return
      setLoading(true)
      setError('')
      const profile = await Promise.allSettled([fetchSignedInUser(msal, account)])
      if (!alive) return
      if (profile[0].status === 'fulfilled') setUser(profile[0].value)
      setLoading(false)
    }
    void load()
    return () => { alive = false }
  }, [account, role, ltiState])

  useEffect(() => {
    if (!msal || !account || role !== 'instructor' || ltiState !== 'active') return
    const client = msal
    const connectedAccount = account
    let alive = true
    let running = false
    async function synchronize(allowInteractive = false) {
      if (running) return
      running = true
      setCourseRecordingLoading(true)
      try {
        const result = await syncCourseRecordings(client, connectedAccount, course?.courseCode ?? '', allowInteractive)
        if (alive) {
          setRecordings(result)
          setRecordingError('')
        }
      } catch (cause) {
        if (alive) setRecordingError(cause instanceof Error
          ? cause.message : 'No se pudieron sincronizar las grabaciones de este curso.')
      } finally {
        running = false
        if (alive) setCourseRecordingLoading(false)
      }
    }
    const refresh = () => {
      if (document.visibilityState === 'visible') void synchronize()
    }
    void synchronize()
    const interval = window.setInterval(refresh, 60_000)
    window.addEventListener('focus', refresh)
    document.addEventListener('visibilitychange', refresh)
    return () => {
      alive = false
      window.clearInterval(interval)
      window.removeEventListener('focus', refresh)
      document.removeEventListener('visibilitychange', refresh)
    }
  }, [account, course?.courseCode, course?.nrc, ltiState, role])

  const weekDays = useMemo(() => weekdays.map((day, index) => {
    const date = addDays(week, index)
    return { ...day, date, key: localISODate(date) }
  }), [week])
  const visibleDays = weekDays.filter((day) => day.value <= 6)
  const weekDrafts = drafts.filter((item) => {
    const date = new Date(`${item.date}T12:00:00`)
    return date >= week && date < addDays(week, 7)
  })
  const courseEvents = courseSessions.map(courseSessionEvent)
  const weekEvents = courseEvents.filter((item) => {
    const date = graphDate(item.start.dateTime)
    return date >= week && date < addDays(week, 7)
  })
  const pendingCount = drafts.filter((draft) => draft.status !== 'created').length
  const unpublishedDrafts = drafts.filter((draft) => draft.status !== 'created')
  const displayName = user?.displayName || account?.name || 'Docente'
  const courseTitle = course?.courseName || 'Espacio docente'
  const termTitle = course?.academicPeriod || course?.term || 'Período no disponible'
  const courseVerified = ltiState === 'active'
    && role === 'instructor'
    && Boolean(course?.courseName && course.courseCode && course.nrc)
  const eligibleMeetings = courseEvents.filter(isTeamsOnlineEvent)
  const currentWeekLabel = `${formatDay(week)} – ${formatDay(addDays(week, 5), { year: week.getFullYear() !== addDays(week, 5).getFullYear() ? 'numeric' : undefined })}`

  function signIn() {
    if (!msal || role !== 'instructor' || !courseVerified) return
    setSigningIn(true)
    setError('')
    const loginUrl = new URL('/', window.location.origin)
    loginUrl.searchParams.set('microsoftLogin', '1')
    const loginWindow = window.open(loginUrl, '_blank')
    setSigningIn(false)
    if (!loginWindow) {
      setError('El navegador bloqueó la nueva pestaña. Permite ventanas emergentes para este sitio e inténtalo de nuevo.')
      return
    }
    loginWindow.opener = null
    setNotice('Se abrió una nueva pestaña para iniciar sesión con Microsoft.')
    window.setTimeout(() => setNotice(''), 7000)
  }

  async function signOut() {
    if (!msal || !account) return
    await msal.logoutPopup({ account })
    if (course?.contextId) {
      window.localStorage.removeItem(microsoftAccountBindingKey(course.contextId))
    }
    setAccount(null)
    setUser(null)
    setRecordings([])
  }

  async function refreshGraphData() {
    if (role !== 'instructor' || !courseVerified) return
    setLoading(true)
    setError('')
    setRecordingError('')
    setCourseSessionsLoading(true)
    const recordingsRequest = msal && account
      ? syncCourseRecordings(msal, account, course?.courseCode ?? '', false)
      : Promise.resolve<CourseRecording[] | null>(null)
    const [courseMeetings, courseRecordings] = await Promise.allSettled([
      fetchCourseSessions(),
      recordingsRequest,
    ])
    if (courseMeetings.status === 'fulfilled') {
      setCourseSessions(courseMeetings.value)
      setCourseSessionsError('')
    } else setCourseSessionsError(courseMeetings.reason instanceof Error
      ? courseMeetings.reason.message : 'No se pudieron consultar las reuniones de este curso.')
    setCourseSessionsLoading(false)
    if (courseRecordings.status === 'fulfilled' && courseRecordings.value) setRecordings(courseRecordings.value)
    else if (courseRecordings.status === 'rejected') setRecordingError(courseRecordings.reason instanceof Error ? courseRecordings.reason.message : 'No se pudieron consultar las grabaciones.')
    setLoading(false)
  }

  async function refreshCourseRecordingFeed() {
    if (!msal || !account || role !== 'instructor' || !courseVerified) return
    setCourseRecordingLoading(true)
    setRecordingError('')
    try {
      await requestRecordingShareAccess(msal, account)
      const result = await syncCourseRecordings(msal, account, course?.courseCode ?? '', false)
      setRecordings(result)
    } catch (cause) {
      setRecordingError(cause instanceof Error
        ? cause.message : 'No se pudieron sincronizar las grabaciones de este curso.')
    } finally {
      setCourseRecordingLoading(false)
    }
  }

  async function removeCourseSession(meeting: GraphEvent) {
    if (role !== 'instructor' || !courseVerified || !course || deletingSessionId) return
    if (!msal || !account) {
      setCourseSessionsError('Conecta la cuenta Microsoft que creó esta reunión antes de eliminarla.')
      return
    }
    const session = courseSessions.find((item) => item.graphEventId === meeting.id)
    if (!session) {
      setCourseSessionsError('No se encontró la sesión que se quiere eliminar. Actualiza la lista e inténtalo de nuevo.')
      return
    }
    const recoveryKey = courseSessionDeletionKey(course.contextId, session.id)
    let microsoftAlreadyCanceled = pendingCourseSessionDeletions.has(recoveryKey)
    try {
      microsoftAlreadyCanceled ||= window.localStorage.getItem(recoveryKey) === '1'
    } catch {
      setCourseSessionsError('El navegador no permitió comprobar el estado de eliminación. Habilita el almacenamiento local e inténtalo de nuevo.')
      return
    }
    const confirmed = window.confirm(
      microsoftAlreadyCanceled
        ? `Microsoft ya canceló "${session.topic}". ¿Quieres quitar ahora su publicación de este curso?\n\n`
          + 'También se quitará la grabación vinculada del curso, si existe; el archivo de OneDrive y su vínculo se conservarán.'
        : `¿Cancelar y eliminar "${session.topic}" del calendario de Microsoft Teams y de este curso?\n\n`
          + 'Microsoft puede enviar una notificación de cancelación a las personas invitadas. ' +
      'También se quitará del curso la grabación vinculada, si existe, pero no se eliminará el archivo de OneDrive ni se revocará su vínculo.'
    )
    if (!confirmed) return

    setDeletingSessionId(session.id)
    setCourseSessionsError('')
    try {
      let recoverySaved = microsoftAlreadyCanceled
      if (!microsoftAlreadyCanceled) {
        try {
          await deleteTeamsEvent(msal, account, session.graphEventId)
        } catch (cause) {
          if (cause instanceof MicrosoftGraphError && cause.status === 404) {
            throw new Error(
              'Microsoft no encontró esta reunión en el calendario de la cuenta conectada. No se quitó la sesión del curso. ' +
              'Conecta la cuenta docente que creó la reunión y vuelve a intentarlo.'
            )
          }
          throw new Error(
            `No se pudo cancelar la reunión en Microsoft Teams; la sesión sigue publicada en el curso. ${
              cause instanceof Error ? cause.message : ''
            }`.trim()
          )
        }
        microsoftAlreadyCanceled = true
        pendingCourseSessionDeletions.add(recoveryKey)
        try {
          window.localStorage.setItem(recoveryKey, '1')
          recoverySaved = true
        } catch {
          recoverySaved = false
        }
      }

      try {
        await deleteCourseSession(session.id)
      } catch (cause) {
        throw new Error(
          `La reunión ya se canceló en Microsoft Teams, pero no se pudo quitar su publicación del curso. ${
            cause instanceof Error ? cause.message : ''
          } ${recoverySaved
            ? 'Actualiza la página y vuelve a pulsar Eliminar para completar la limpieza.'
            : 'No actualices la página; vuelve a pulsar Eliminar para reintentar sin cancelar otra vez la reunión.'
          }`.trim()
        )
      }
      pendingCourseSessionDeletions.delete(recoveryKey)
      try {
        window.localStorage.removeItem(recoveryKey)
      } catch {
        console.warn('No se pudo limpiar la marca local de eliminación de la sesión.')
      }
      setCourseSessions((current) => current.filter((item) => item.id !== session.id))
      setRecordings((current) => current.filter((item) => item.sessionId !== session.id))
      setNotice('La reunión se canceló en Microsoft Teams y se quitó del curso. La grabación de OneDrive no se eliminó ni se revocó su vínculo.')
    } catch (cause) {
      setCourseSessionsError(cause instanceof Error
        ? cause.message : 'No se pudo eliminar la sesión del curso.')
    } finally {
      setDeletingSessionId('')
    }
  }

  function openComposer(mode: FormMode, draft?: SessionDraft) {
    if (role !== 'instructor' || !courseVerified) return
    const defaultSessionTitle = [course?.courseCode, course?.courseName]
      .filter((value): value is string => Boolean(value?.trim()))
      .filter((value, index, values) => values.findIndex((candidate) => candidate.toLocaleLowerCase() === value.toLocaleLowerCase()) === index)
      .join(' - ')
    setEditingId(draft?.id ?? null)
    setModal(mode)
    setFormTitle(draft?.title ?? defaultSessionTitle.slice(0, 255))
    setStartTime(draft?.start ?? '09:00')
    setEndTime(draft?.end ?? '10:00')
    setFormDate(draft?.date ?? localISODate(addDays(today, 1)))
    setTermEnd(localISODate(addDays(today, 60)))
    setSelectedDays(draft ? [new Date(`${draft.date}T12:00:00`).getDay() || 7] : [1])
    setWeeklyDaysInitialized(mode === 'weekly' || Boolean(draft))
    setTimeZone(draft?.timeZone ?? timeZones[0])
    setAttendees(draft?.attendees.join(', ') ?? '')
    setDescription(draft?.description ?? '')
    setLocation(draft?.location ?? '')
  }

  function changeComposerMode(mode: FormMode) {
    if (!modal || editingId) return
    if (mode === 'weekly') {
      if (!weeklyDaysInitialized) {
        setSelectedDays([new Date(`${formDate}T12:00:00`).getDay() || 7])
        setWeeklyDaysInitialized(true)
      }
      if (termEnd < formDate) setTermEnd(formDate)
    }
    setModal(mode)
  }

  function createDrafts(event: FormEvent<HTMLFormElement>) {
    event.preventDefault()
    if (role !== 'instructor' || !courseVerified) return
    if (startTime >= endTime) {
      setError('La hora de término debe ser posterior a la hora de inicio.')
      return
    }
    const people = attendees.split(/[,\n;]/).map((email) => email.trim()).filter(Boolean)
    const occurrences: string[] = []
    let limitNotice = ''
    if (modal === 'weekly') {
      if (!selectedDays.length) {
        setError('Selecciona al menos un día de clase.')
        return
      }
      const from = new Date(`${formDate}T12:00:00`)
      const until = new Date(`${termEnd}T12:00:00`)
      if (until < from) {
        setError('El fin del período debe ser posterior al inicio.')
        return
      }
      for (let cursor = from; cursor <= until && occurrences.length < 120; cursor = addDays(cursor, 1)) {
        const weekday = cursor.getDay() || 7
        if (selectedDays.includes(weekday)) occurrences.push(localISODate(cursor))
      }
      if (!occurrences.length) {
        setError('El rango elegido no contiene ninguno de los días seleccionados.')
        return
      }
      limitNotice = occurrences.length === 120
        ? 'Se generaron las primeras 120 fechas del período para mantener una revisión manejable. '
        : ''
    } else occurrences.push(formDate)

    const common = { title: formTitle.trim(), start: startTime, end: endTime, timeZone, attendees: modal === 'weekly' && !editingId ? [] : people, description: description.trim(), location: location.trim(), recurrence: modal === 'weekly' }
    let reviewItem: SessionDraft | null = null
    if (editingId) {
      const existing = drafts.find((draft) => draft.id === editingId)
      if (existing && existing.status !== 'created') {
        const updatedItem = { ...existing, ...common, date: occurrences[0] ?? existing.date, status: 'review' as const }
        reviewItem = updatedItem
        setDrafts((current) => current.map((draft) => (
          draft.id === editingId && draft.status !== 'created' ? updatedItem : draft
        )))
      }
    } else {
      const created = occurrences.map((date) => ({
        id: crypto.randomUUID(), ...common, date,
        status: 'ready' as const,
      }))
      if (modal === 'single') reviewItem = created[0] ?? null
      setDrafts((current) => [...current, ...created])
      if (modal === 'weekly') {
        setError('')
        setModal(null)
        setView('schedule')
        setWeek(dayStart(new Date(`${occurrences[0]}T12:00:00`)))
        setMobileDay(Math.min((new Date(`${occurrences[0]}T12:00:00`).getDay() + 6) % 7, 5))
        if (msal && account) {
          void publishSessions(created)
        } else {
          setNotice('El horario quedó preparado, pero todavía no se publicó. Conecta Microsoft y luego pulsa “Reintentar publicación”.')
          window.setTimeout(() => setNotice(''), 12000)
        }
        return
      }
    }
    setError('')
    setModal(null)
    setView('schedule')
    if (modal === 'single' && reviewItem) {
      setReviewError('')
      setReviewing(reviewItem)
    } else {
      setNotice(`${limitNotice}${occurrences.length} sesiones guardadas como borrador local. Revísalas y confirma cada reunión de Teams.`)
    }
    window.setTimeout(() => setNotice(''), 7500)
    if (occurrences.length) {
      const firstDate = new Date(`${occurrences[0]}T12:00:00`)
      setWeek(dayStart(firstDate))
      setMobileDay(Math.min((firstDate.getDay() + 6) % 7, 5))
    }
  }

  async function createReviewedMeeting(session: SessionDraft) {
    if (role !== 'instructor' || !courseVerified) {
      setReviewError('Solo un docente con un curso verificado puede crear reuniones.')
      return
    }
    if (!msal || !account) {
      setReviewError('Inicia sesión en Microsoft antes de crear el evento de Teams.')
      return
    }
    setSaving(true)
    setReviewError('')
    try {
      const created = session.graphEvent ?? await createTeamsEvent(msal, account, {
        subject: session.title,
        start: `${session.date}T${session.start}:00`,
        end: `${session.date}T${session.end}:00`,
        timeZone: session.timeZone,
        attendees: session.attendees,
        description: session.description,
        location: session.location,
      })
      setDrafts((current) => current.map((item) => item.id === session.id
        ? { ...item, graphEvent: created }
        : item))
      const readyToPublish = await waitForTeamsJoinLink(msal, account, created)
      const courseSession = await publishCourseMeeting(readyToPublish)
      setCourseSessions((current) => [
        ...current.filter((item) => item.graphEventId !== courseSession.graphEventId),
        courseSession,
      ].sort((a, b) => Date.parse(a.start) - Date.parse(b.start)))
      setCourseSessionsError('')
      setDrafts((current) => current.map((item) => item.id === session.id
        ? { ...item, status: 'created', graphEvent: readyToPublish }
        : item))
      setReviewing(null)
      setNotice('Reunión creada en Teams y publicada para los estudiantes de este curso.')
      window.setTimeout(() => setNotice(''), 7000)
    } catch (cause) {
      setReviewError(cause instanceof Error ? cause.message : 'No se pudo crear el evento en Teams.')
    } finally { setSaving(false) }
  }

  function sessionClass(status: DraftStatus) {
    return status === 'created' ? 'created' : status === 'ready' ? 'ready' : 'review'
  }

  async function publishCourseMeeting(event: GraphEvent): Promise<CourseSession> {
    const joinUrl = event.onlineMeeting?.joinUrl
    if (!event.id || !event.subject || !event.start.dateTime || !event.end.dateTime
      || !event.start.timeZone || !joinUrl || !isTeamsOnlineEvent(event)) {
      throw new Error('Microsoft creó el evento, pero no devolvió un enlace de reunión de Teams válido para publicarlo al curso.')
    }
    const response = await fetch('/api/course/sessions', {
      method: 'POST',
      credentials: 'same-origin',
      headers: { Accept: 'application/json', 'Content-Type': 'application/json' },
      body: JSON.stringify({
        graphEventId: event.id,
        title: event.subject,
        start: event.start.dateTime,
        end: event.end.dateTime,
        timeZone: event.start.timeZone,
        joinUrl,
      }),
    })
    if (!response.ok) {
      const result = await response.json().catch(() => null) as { error?: string } | null
      throw new Error(result?.error || `No se pudo publicar la reunión al curso (HTTP ${response.status}).`)
    }
    const result: unknown = await response.json()
    const published = isObject(result) ? parseCourseSession(result.session) : null
    if (!published) {
      throw new Error('El servidor publicó la reunión, pero devolvió datos incompletos del curso.')
    }
    return published
  }

  async function publishSessions(toPublish: SessionDraft[]) {
    if (role !== 'instructor' || !courseVerified || !msal || !account || publishingBatch || !toPublish.length) return
    setPublishingBatch(true)
    setError('')
    let published = 0
    const failures: string[] = []
    try {
      for (const session of toPublish) {
        try {
          const created = session.graphEvent ?? await createTeamsEvent(msal, account, {
              subject: session.title,
              start: `${session.date}T${session.start}:00`,
              end: `${session.date}T${session.end}:00`,
              timeZone: session.timeZone,
              attendees: session.attendees,
              description: session.description,
              location: session.location,
            })
          setDrafts((current) => current.map((item) => item.id === session.id
            ? { ...item, graphEvent: created }
            : item))
          const readyToPublish = await waitForTeamsJoinLink(msal, account, created)
          const courseSession = await publishCourseMeeting(readyToPublish)
          setCourseSessions((current) => [
            ...current.filter((item) => item.graphEventId !== courseSession.graphEventId),
            courseSession,
          ].sort((a, b) => Date.parse(a.start) - Date.parse(b.start)))
          setCourseSessionsError('')
          setDrafts((current) => current.map((item) => item.id === session.id
            ? { ...item, status: 'created', graphEvent: readyToPublish }
            : item))
          published++
        } catch (cause) {
          failures.push(`${session.title} (${session.date}): ${cause instanceof Error ? cause.message : 'error desconocido'}`)
        }
      }
      setNotice(failures.length
        ? `Se publicaron ${published} de ${toPublish.length} sesiones. Fallaron ${failures.length}; puedes reintentar desde el botón de publicación. ${failures.slice(0, 3).join(' · ')}`
        : `Se crearon y publicaron ${published} sesiones en Teams para este curso. Los estudiantes ya pueden ver sus enlaces.`)
      window.setTimeout(() => setNotice(''), 12000)
    } finally {
      setPublishingBatch(false)
    }
  }

  async function publishAllDrafts() {
    await publishSessions(unpublishedDrafts.slice(0, 120))
  }

  const mobileDate = weekDays[Math.min(mobileDay, 5)]
  const mobileItems: Array<{ kind: 'draft'; item: SessionDraft } | { kind: 'graph'; item: GraphEvent }> = [
    ...weekDrafts.filter((item) => item.date === mobileDate.key).map((item) => ({ kind: 'draft' as const, item })),
    ...weekEvents.filter((item) => graphLocalDate(item.start.dateTime) === mobileDate.key).map((item) => ({ kind: 'graph' as const, item })),
  ].sort((a, b) => (a.kind === 'draft' ? a.item.start : graphLocalTime(a.item.start.dateTime)).localeCompare(b.kind === 'draft' ? b.item.start : graphLocalTime(b.item.start.dateTime)))

  if (role === 'learner' && ltiState === 'active' && course) {
    return <LearnerWorkspace view={view} setView={setView} />
  }
  if (ltiState === 'loading') {
    return <LtiLoading />
  }
  if (ltiState !== 'active' || role !== 'instructor' || !course) {
    return <AccessDenied />
  }

  return (
    <div className={advancedToolsOpen ? 'app-shell advanced-shell' : 'app-shell compact-shell'}>
      <aside className="sidebar">
        <a className="brand" href="#" onClick={(event) => { event.preventDefault(); setView('schedule') }} aria-label="Microsoft Teams, inicio">
          <span className="brand-mark"><span /><span /><span /><span /></span><span className="brand-name">Microsoft Teams<span>.</span></span>
        </a>
        <div className="workspace-label">ESPACIO DOCENTE</div>
        <div className="course-switcher">
          <span className="course-avatar">{course?.courseCode ? course.courseCode.slice(0, 2).toUpperCase() : '—'}</span>
          <span className="course-switch-copy"><strong>{courseTitle}</strong><small>{course?.courseCode || 'Sin curso validado'}</small></span>
          <ChevronRight size={15} />
        </div>
        <nav className="primary-nav" aria-label="Navegación del curso">
          <button className={view === 'schedule' ? 'nav-item active' : 'nav-item'} onClick={() => setView('schedule')}>
            <CalendarDays size={18} /><span>Cronograma</span>{pendingCount > 0 && <span className="nav-count">{pendingCount}</span>}
          </button>
          <button className={view === 'recordings' ? 'nav-item active' : 'nav-item'} onClick={() => setView('recordings')}>
            <Play size={18} /><span>Grabaciones</span>
          </button>
          <button className={view === 'summary' ? 'nav-item active' : 'nav-item'} onClick={() => setView('summary')}>
            <FileVideo2 size={18} /><span>Resumen de reuniones</span>
          </button>
        </nav>
        <div className="sidebar-bottom">
          <div className="sidebar-note"><span className="note-spark"><ShieldCheck size={17} /></span><p>Planifica primero. Publica solo cuando estés listo.</p><span className="note-caption">BORRADORES VISIBLES SOLO AQUÍ</span></div>
          {account ? (
            <div className="account-block">
              <span className="avatar">{initials(user?.displayName || account.name)}</span>
              <span className="account-copy"><strong>{user?.displayName || account.name || 'Cuenta conectada'}</strong><small>Microsoft 365 · Graph</small></span>
              <button className="icon-button" onClick={() => void signOut()} aria-label="Cerrar sesión"><LogOut size={16} /></button>
            </div>
          ) : (
            <button className="sidebar-signin" onClick={() => void signIn()} disabled={!msal || signingIn}>
              <span className="avatar muted-avatar"><Video size={16} /></span><span className="account-copy"><strong>Conectar Teams</strong><small>{msal ? 'Cuenta institucional' : 'Configura Microsoft'}</small></span><ArrowRight size={15} />
            </button>
          )}
        </div>
      </aside>

      <main className="main-content">
        <header className="topbar">
          <div className="breadcrumb"><span>Brightspace</span><span className="slash">/</span><strong>{course?.courseCode || 'Curso sin contexto'}</strong></div>
          <div className="topbar-right"><span className="term-label">{termTitle}</span>
            {advancedToolsOpen && <button className="compact-view-button" onClick={() => setAdvancedToolsOpen(false)}>Vista compacta</button>}
            {account
              ? <button className="user-pill" onClick={() => void signOut()} title="Cerrar sesión"><span className="avatar tiny-avatar">{initials(displayName)}</span><span className="pill-name">{displayName.split(' ')[0]}</span><ChevronRight size={14} /></button>
              : <span className="connection-state"><span className="state-dot" />{ltiState === 'active' ? 'Curso verificado' : 'Vista de borrador'}</span>}
          </div>
        </header>

        {!authReady ? <div className="center-state"><LoaderCircle className="spin" size={23} />Preparando Microsoft 365…</div> : (
          <div className={advancedToolsOpen ? 'page-content advanced-page-content' : 'page-content compact-page-content'}>
            {!advancedToolsOpen ? (
              <section className="zoom-workspace" aria-label="Reuniones del curso">
                <header className="zoom-brand-row">
                  <a className="zoom-wordmark" href="#" onClick={(event) => { event.preventDefault(); setView('schedule') }}>Microsoft Teams</a>
                  <nav className="zoom-shortcuts" aria-label="Navegación principal">
                    <button onClick={() => { setView('schedule'); setTeacherPeriod('upcoming') }}>Inicio</button>
                    <button onClick={() => setView('schedule')}>Citas</button>
                  </nav>
                </header>
                <div className="zoom-timezone">Su idioma y su zona horaria actual son ({new Intl.DateTimeFormat('es-CO', { timeZoneName: 'short' }).format(new Date()).split(' ').pop()}), Español</div>
                <div className="zoom-action-row">
                  <button className="zoom-schedule-button" onClick={() => openComposer('single')}><Plus size={14} />Programar una reunión nueva</button>
                  <button className="zoom-all-link" onClick={() => setView('summary')}>Reuniones de esta clase</button>
                  <div className="zoom-overflow-wrap">
                    <button className="zoom-overflow" aria-label="Más acciones" aria-expanded={teacherOverflowOpen} onClick={() => setTeacherOverflowOpen((open) => !open)}>⋮</button>
                    {teacherOverflowOpen && <div className="zoom-overflow-menu">
                      <button onClick={() => { setTeacherOverflowOpen(false); openComposer('weekly') }}>Programar sesiones semanales</button>
                      {unpublishedDrafts.length > 0 && <button onClick={() => { setTeacherOverflowOpen(false); void publishAllDrafts() }} disabled={!courseVerified || !account || publishingBatch}>Crear y publicar {unpublishedDrafts.length} sesiones</button>}
                      <button onClick={() => { setAdvancedToolsOpen(true); setTeacherOverflowOpen(false) }}>Herramientas de planificación</button>
                      <button onClick={() => { setTeacherOverflowOpen(false); void refreshGraphData() }} disabled={!account || loading}>Actualizar reuniones de este curso</button>
                      {!account && <button onClick={() => { setTeacherOverflowOpen(false); void signIn() }} disabled={!msal || signingIn}>Conectar Microsoft</button>}
                    </div>}
                  </div>
                </div>
                {publishingBatch && <div className="success-banner"><LoaderCircle size={16} className="spin" />Creando las reuniones en Microsoft Teams y publicándolas para este curso. No cierres esta página…</div>}
                {notice && <div className="success-banner"><Check size={16} />{notice}<button onClick={() => setNotice('')} aria-label="Cerrar aviso"><X size={15} /></button></div>}
                {error && <ErrorBanner message={error} />}
                <nav className="zoom-tabs" aria-label="Secciones de reuniones">
                  <button className={view === 'schedule' ? 'zoom-tab active' : 'zoom-tab'} onClick={() => setView('schedule')}>Reuniones</button>
                  <button className={view === 'recordings' ? 'zoom-tab active' : 'zoom-tab'} onClick={() => setView('recordings')}>Grabaciones en la nube</button>
                  <button className={view === 'notes' ? 'zoom-tab active' : 'zoom-tab'} onClick={() => setView('notes')}>Mis notas</button>
                  <button className={view === 'summary' ? 'zoom-tab active' : 'zoom-tab'} onClick={() => setView('summary')}>Resumen de reuniones</button>
                </nav>
                {view === 'schedule' && <>
                  <nav className="zoom-subtabs" aria-label="Filtro de reuniones">
                    <button className={teacherPeriod === 'upcoming' ? 'zoom-subtab active' : 'zoom-subtab'} onClick={() => setTeacherPeriod('upcoming')}>Próximas reuniones</button>
                    <button className={teacherPeriod === 'past' ? 'zoom-subtab active' : 'zoom-subtab'} onClick={() => setTeacherPeriod('past')}>Reuniones pasadas</button>
                    <button className={teacherPeriod === 'personal' ? 'zoom-subtab active' : 'zoom-subtab'} onClick={() => setTeacherPeriod('personal')}>Sala de reunión personal</button>
                  </nav>
                  {teacherPeriod === 'personal' ? <div className="zoom-empty-state">La sala de reunión personal no está configurada para esta cuenta.</div>
                    : courseSessionsError ? <ErrorBanner message={courseSessionsError} />
                      : courseSessionsLoading && !courseEvents.length ? <div className="zoom-empty-state">Consultando reuniones de este curso…</div>
                        : <CompactTeacherMeetingTable meetings={eligibleMeetings.filter((meeting) => teacherPeriod === 'upcoming' ? Date.parse(meeting.end.dateTime) >= Date.now() : Date.parse(meeting.end.dateTime) < Date.now()).slice().sort((a, b) => teacherPeriod === 'upcoming' ? Date.parse(a.start.dateTime) - Date.parse(b.start.dateTime) : Date.parse(b.start.dateTime) - Date.parse(a.start.dateTime))} onDelete={removeCourseSession} deletingSessionId={deletingSessionId} />}
                </>}
                {view === 'recordings' && <section className="zoom-content-panel">
                  <div className="zoom-section-title">Grabaciones en la nube</div>
                  {account && <button className="secondary-button" onClick={() => void refreshCourseRecordingFeed()} disabled={courseRecordingLoading}><RotateCw size={15} />Actualizar grabaciones</button>}
                  {!account ? <div className="zoom-empty-state">Conecta Microsoft para detectar las grabaciones de las reuniones de este curso.<button className="outline-button" onClick={() => void signIn()} disabled={!msal || signingIn}>Conectar Microsoft</button></div>
                    : recordingError ? <ErrorBanner message={recordingError} />
                    : recordings.length ? <div className="recordings-library">{recordings.map((item) => <CourseRecordingRow key={item.id} recording={item} />)}</div>
                      : <div className="zoom-empty-state">{courseRecordingLoading ? 'Buscando grabaciones de este curso…' : 'Todavía no hay grabaciones publicadas para este curso.'}</div>}
                </section>}
                {view === 'summary' && <section className="zoom-content-panel">
                  <div className="zoom-section-title">Resumen de reuniones</div>
                  {courseSessionsError ? <ErrorBanner message={courseSessionsError} />
                    : <div className="zoom-summary-groups">{[{ label: 'Próximas reuniones', items: eligibleMeetings.filter((meeting) => Date.parse(meeting.end.dateTime) >= Date.now()) }, { label: 'Reuniones pasadas', items: eligibleMeetings.filter((meeting) => Date.parse(meeting.end.dateTime) < Date.now()) }].map((group) => <section key={group.label}><h2>{group.label}</h2><CompactTeacherMeetingTable meetings={group.items} onDelete={removeCourseSession} deletingSessionId={deletingSessionId} /></section>)}</div>}
                </section>}
                {view === 'notes' && <section className="zoom-content-panel zoom-notes-panel"><label htmlFor="teacher-notes">Mis notas</label><textarea id="teacher-notes" value={notes} onChange={(event) => setNotes(event.target.value)} placeholder="Anota aquí recordatorios para este curso…" /><small>Notas temporales: solo permanecen en esta pestaña y no se comparten con estudiantes.</small></section>}
              </section>
            ) : view === 'schedule' ? (
              <>
                <section className="course-heading">
                  <div className="course-heading-copy">
                    <div className="eyebrow">PLANIFICACIÓN DOCENTE <span className="eyebrow-divider">/</span> {course?.courseCode || 'SIN CONTEXTO LTI'}</div>
                    <h1>{courseTitle}<span className="heading-period">.</span></h1>
                    <p>{course?.instructorName || displayName} <span className="bullet">·</span> {termTitle}</p>
                  </div>
                  <div className="heading-actions">
                    <button className="secondary-button" onClick={() => setNotesOpen((open) => !open)} aria-expanded={notesOpen}><Pencil size={15} />Mis notas</button>
                    <button className="secondary-button personal-room-button" disabled title="La sala personal no está configurada en este entorno"><Video size={15} />Sala de reunión personal</button>
                    <button className="secondary-button" onClick={() => openComposer('single')}><Plus size={16} />Reunión única</button>
                    <button className="primary-button" onClick={() => openComposer('weekly')}><CalendarDays size={16} />Crear horario semanal</button>
                  </div>
                </section>

                {notesOpen && <section className="teacher-notes" aria-label="Mis notas">
                  <label htmlFor="teacher-notes">Mis notas</label>
                  <textarea id="teacher-notes" value={notes} onChange={(event) => setNotes(event.target.value)} placeholder="Anota aquí recordatorios para este curso…" />
                  <small>Notas temporales: solo permanecen en esta pestaña y no se comparten con estudiantes.</small>
                </section>}
                {ltiState !== 'active' && <div className="integration-notice"><AlertCircle size={17} /><span><strong>Brightspace aún no ha validado el contexto de curso.</strong> Esta es una vista de planificación; no representa un curso conectado. Los borradores se mantienen solo en esta sesión del navegador.</span><a href="/README.md">Requisitos LTI 1.3 <ArrowRight size={13} /></a></div>}
                {error && <ErrorBanner message={error} />}
                {notice && <div className="success-banner"><Check size={16} />{notice}<button onClick={() => setNotice('')} aria-label="Cerrar aviso"><X size={15} /></button></div>}

                <section className="existing-meetings-section" aria-labelledby="existing-meetings-title">
                  <div className="section-heading existing-meetings-heading">
                    <div><div className="eyebrow">SESIONES DEL CONTEXTO BRIGHTSPACE ACTUAL</div><h2 id="existing-meetings-title">Reuniones de esta clase</h2></div>
                    {courseVerified && <span className="course-scope-badge">{course?.courseName} · {course?.courseCode}{course?.academicPeriod ? ` · ${course.academicPeriod}` : ''} · NRC {course?.nrc}</span>}
                  </div>
                  {courseSessionsError ? <ErrorBanner message={courseSessionsError} />
                    : courseSessionsLoading && !courseSessions.length ? <div className="session-empty"><LoaderCircle className="spin" size={17} /><span><strong>Consultando las reuniones de este curso…</strong></span></div>
                      : courseSessions.length ? <CompactTeacherMeetingTable meetings={eligibleMeetings} onDelete={removeCourseSession} deletingSessionId={deletingSessionId} />
                        : <div className="session-empty"><span className="empty-icon"><CalendarDays size={18} /></span><span><strong>Este curso aún no tiene reuniones publicadas</strong><small>Solo se muestran aquí las reuniones creadas y publicadas desde el curso actual.</small></span></div>}
                </section>

                <section className="schedule-card">
                  <div className="schedule-toolbar">
                    <div><div className="eyebrow">VISTA SEMANAL</div><h2>Cronograma del curso</h2></div>
                    <div className="week-controls">
                      <button className="icon-button week-arrow" onClick={() => setWeek(addDays(week, -7))} aria-label="Semana anterior"><ChevronLeft size={17} /></button>
                      <span>{currentWeekLabel}</span>
                      <button className="icon-button week-arrow" onClick={() => setWeek(addDays(week, 7))} aria-label="Semana siguiente"><ChevronRight size={17} /></button>
                      <button className="today-button" onClick={() => { setWeek(dayStart(today)); setMobileDay(Math.min((today.getDay() + 6) % 7, 5)) }}>Hoy</button>
                    </div>
                  </div>
                  <div className="schedule-legend">
                    <span><i className="legend-dot draft-dot" />Borrador / pendiente de revisar</span>
                    <span><i className="legend-dot ready-dot" />Lista para crear en Teams</span>
                    <span><i className="legend-dot created-dot" />Reunión de Teams creada</span>
                    {courseEvents.length > 0 && <span><i className="legend-dot calendar-dot" />Publicado en este curso</span>}
                  </div>

                  <div className="weekly-grid">
                    <div className="time-column">
                      <div className="grid-day-head time-head">HORA LOCAL</div>
                      {hours.map((hour) => <div className="hour-label" key={hour}>{`${String(hour).padStart(2, '0')}:00`}</div>)}
                    </div>
                    {visibleDays.map((day) => (
                      <div className="day-column" key={day.key}>
                        <div className={`grid-day-head ${day.key === localISODate(today) ? 'is-today' : ''}`}><span>{day.short}</span><strong>{day.date.getDate()}</strong></div>
                        <div className="day-track">
                          {hours.map((hour) => <div className="hour-line" key={hour} />)}
                          {weekDrafts.filter((item) => item.date === day.key).map((item) => <DraftBlock key={item.id} item={item} onOpen={() => setReviewing(item)} statusClass={sessionClass(item.status)} />)}
                          {weekEvents.filter((item) => graphLocalDate(item.start.dateTime) === day.key).map((item) => <GraphBlock key={item.id} item={item} />)}
                        </div>
                      </div>
                    ))}
                    {weekDrafts.length === 0 && weekEvents.length === 0 && (
                      <div className="calendar-empty-cue">
                        <span><CalendarDays size={16} /></span>
                        <strong>Semana sin sesiones</strong>
                        <small>Los horarios que planifiques aparecerán aquí.</small>
                        <button className="empty-grid-action" onClick={() => openComposer('weekly')}>
                          Crear horario <ArrowRight size={13} />
                        </button>
                      </div>
                    )}
                  </div>
                  <div className="mobile-schedule">
                    <div className="mobile-day-picker">
                      {visibleDays.map((day, index) => <button key={day.key} className={mobileDay === index ? 'mobile-day active' : 'mobile-day'} onClick={() => setMobileDay(index)}><span>{day.short}</span><strong>{day.date.getDate()}</strong></button>)}
                    </div>
                    <h3>{formatLongDay(mobileDate.date)}</h3>
                    {mobileItems.length ? <div className="mobile-session-list">{mobileItems.map((entry) => entry.kind === 'draft'
                      ? <DraftListItem key={entry.item.id} item={entry.item} onOpen={() => setReviewing(entry.item)} statusClass={sessionClass(entry.item.status)} />
                      : <GraphListItem key={entry.item.id} item={entry.item} />)}</div>
                      : <div className="mobile-empty"><CalendarDays size={18} /><span>No hay sesiones programadas para este día.</span><button className="text-action" onClick={() => openComposer('weekly')}>Crear horario <ArrowRight size={13} /></button></div>}
                  </div>
                  <div className="schedule-foot"><span><Clock3 size={13} /> Horario local del docente · zona del evento según formulario</span><span>{weekDrafts.length + weekEvents.length === 1 ? '1 sesión esta semana' : `${weekDrafts.length + weekEvents.length} sesiones esta semana`}</span></div>
                </section>

                <section className="session-review-section">
                  <div className="section-heading"><div><div className="eyebrow">REVISIÓN ANTES DE PUBLICAR</div><h2>Sesiones del período <span className="section-count">{drafts.length}</span></h2></div>
                    <div className="section-actions">
                      {unpublishedDrafts.length > 0 && <button className="primary-button" onClick={() => void publishAllDrafts()} disabled={!courseVerified || !account || publishingBatch}>
                        {publishingBatch ? <LoaderCircle size={15} className="spin" /> : <Video size={15} />}
                        {publishingBatch ? 'Creando sesiones…' : `Reintentar publicación de ${unpublishedDrafts.length} sesiones`}
                      </button>}
                      <button className="text-action" onClick={() => openComposer('weekly')}><Plus size={14} /> Añadir horario</button>
                    </div></div>
                  {!drafts.length ? <div className="session-empty"><span className="empty-icon"><CalendarDays size={18} /></span><span><strong>Aún no hay sesiones planificadas</strong><small>Agrega un horario semanal y las sesiones se crearán en Microsoft Teams y se publicarán para los estudiantes del curso.</small></span></div>
                    : <div className="session-table">{drafts.slice().sort((a, b) => `${a.date}${a.start}`.localeCompare(`${b.date}${b.start}`)).slice(0, 80).map((item) =>
                      <SessionRow key={item.id} item={item} onReview={() => setReviewing(item)} onEdit={() => openComposer(item.recurrence ? 'weekly' : 'single', item)} onCreate={() => setReviewing(item)} statusClass={sessionClass(item.status)} />)}
                    {drafts.length > 80 && <p className="limit-note">Mostrando 80 sesiones. La serie completa continúa en borradores locales.</p>}</div>}
                  <p className="draft-disclaimer"><ShieldCheck size={14} />El horario semanal crea una reunión independiente en Teams para cada día seleccionado y publica los enlaces en este curso. Deja abierta esta página hasta que finalice la publicación.</p>
                </section>

                <section className="graph-calendar-section">
                  <div className="section-heading"><div><div className="eyebrow">CURSO ACTUAL · {course?.courseCode}</div><h2>Reuniones publicadas en esta clase</h2></div><button className="text-action" onClick={() => void refreshGraphData()} disabled={loading || courseSessionsLoading}><RotateCw size={13} /> Actualizar</button></div>
                  {courseSessionsError ? <ErrorBanner message={courseSessionsError} />
                    : <div className="graph-event-list">
                      {courseEvents.length ? courseEvents.map((item) => <GraphListItem key={item.id} item={item} />)
                        : <div className="mobile-empty">{courseSessionsLoading ? 'Consultando las reuniones de este curso…' : 'Todavía no hay reuniones publicadas para este curso.'}</div>}
                    </div>}
                </section>
              </>
            ) : view === 'recordings' ? (
              <section className="recordings-page">
                <div className="course-heading recordings-heading"><div className="course-heading-copy"><div className="eyebrow">MICROSOFT ONEDRIVE · {course?.courseCode}</div><h1>Grabaciones<span className="heading-period">.</span></h1><p>Grabaciones vinculadas a las reuniones publicadas en este curso.</p></div>{account && <button className="secondary-button" onClick={() => void refreshCourseRecordingFeed()} disabled={loading || courseRecordingLoading}><RotateCw size={15} />Actualizar grabaciones</button>}</div>
                <div className="integration-notice"><FolderOpen size={17} /><span><strong>Acceso institucional:</strong> los enlaces se comparten con cuentas TESA que tengan el vínculo. Para evitar cruces entre clases, solo se publica un video si su nombre coincide con el código y el tema de la reunión.</span></div>
                {!account ? <div className="graph-connection"><span className="connect-icon"><Video size={18} /></span><span><strong>Inicia sesión para detectar grabaciones</strong><small>No se muestran archivos de otros cursos.</small></span><button className="outline-button" onClick={() => void signIn()} disabled={!msal || signingIn}>Conectar Microsoft</button></div>
                  : recordingError ? <ErrorBanner message={recordingError} />
                    : courseRecordingLoading && recordings.length === 0 ? <div className="session-empty"><LoaderCircle className="spin" size={17} /><span><strong>Buscando grabaciones del curso…</strong><small>Las grabaciones se revisan automáticamente mientras esta página está abierta.</small></span></div>
                    : recordings.length ? <div className="recordings-library">{recordings.map((item) => <CourseRecordingRow key={item.id} recording={item} />)}</div>
                      : <div className="session-empty"><span className="empty-icon"><FolderOpen size={18} /></span><span><strong>{courseRecordingLoading ? 'Buscando grabaciones…' : 'Todavía no hay grabaciones para este curso'}</strong><small>El video debe estar en Recordings o Grabaciones de OneDrive y coincidir con el horario de la sesión.</small></span></div>}
              </section>
            ) : (
              <section className="recordings-page meeting-summary-page">
                <div className="course-heading recordings-heading"><div className="course-heading-copy"><div className="eyebrow">MICROSOFT TEAMS · {course?.courseCode}</div><h1>Resumen de reuniones<span className="heading-period">.</span></h1><p>Reuniones publicadas para este curso.</p></div><button className="secondary-button" onClick={() => void refreshGraphData()} disabled={!account || loading || courseSessionsLoading}><RotateCw size={15} />Actualizar</button></div>
                {courseSessionsError ? <ErrorBanner message={courseSessionsError} />
                  : courseSessionsLoading && !courseEvents.length ? <div className="session-empty"><LoaderCircle className="spin" size={17} /><span><strong>Consultando reuniones…</strong></span></div>
                    : <div className="meeting-summary-columns">
                    {[{ title: 'Próximas reuniones', items: eligibleMeetings.filter((item) => Date.parse(item.end.dateTime) >= Date.now()) },
                      { title: 'Reuniones anteriores', items: eligibleMeetings.filter((item) => Date.parse(item.end.dateTime) < Date.now()).slice().reverse() }].map((group) => (
                      <section className="meeting-summary-group" key={group.title}>
                        <h2>{group.title}<span>{group.items.length}</span></h2>
                        {group.items.length ? <div className="graph-event-list">{group.items.slice(0, 5).map((item) => <GraphListItem key={item.id} item={item} />)}</div>
                          : <div className="mobile-empty">Sin datos</div>}
                      </section>
                    ))}
                    </div>}
              </section>
            )}
            <footer className="page-footer"><span>Microsoft Teams <span className="footer-dot">✳</span></span><span>Sesiones locales · calendario y archivos conectados a Microsoft Graph cuando inicias sesión</span></footer>
          </div>
        )}
      </main>

      {modal && <Composer mode={modal} onModeChange={changeComposerMode} formTitle={formTitle} setTitle={setFormTitle} start={startTime} setStart={setStartTime} end={endTime} setEnd={setEndTime} date={formDate} setDate={(value) => { setFormDate(value); if (modal === 'weekly' && value > termEnd) setTermEnd(value) }} termEnd={termEnd} setTermEnd={setTermEnd} selectedDays={selectedDays} setSelectedDays={setSelectedDays} timeZone={timeZone} setTimeZone={setTimeZone} attendees={attendees} setAttendees={setAttendees} description={description} setDescription={setDescription} location={location} setLocation={setLocation} editing={Boolean(editingId)} error={error} onClose={() => setModal(null)} onSubmit={createDrafts} />}

      {reviewing && <ReviewDialog item={reviewing} error={reviewError} saving={saving} connected={Boolean(account && msal)} canConnect={Boolean(msal)} onConnect={() => void signIn()} onClose={() => { setReviewing(null); setReviewError('') }} onEdit={() => { const selected = reviewing; setReviewing(null); setReviewError(''); openComposer(selected.recurrence ? 'weekly' : 'single', selected) }} onMarkReady={() => { if (role !== 'instructor' || !courseVerified) return; setDrafts((current) => current.map((item) => item.id === reviewing.id ? { ...item, status: 'ready' } : item)); setReviewing({ ...reviewing, status: 'ready' }) }} onCreate={() => void createReviewedMeeting(reviewing)} />}
    </div>
  )
}

function DraftBlock({ item, onOpen, statusClass }: { item: SessionDraft; onOpen: () => void; statusClass: string }) {
  const top = ((minutesOf(item.start) - 420) / 60) * 36
  const height = Math.max(30, ((minutesOf(item.end) - minutesOf(item.start)) / 60) * 36 - 4)
  return <button className={`calendar-block draft-block ${statusClass}`} style={{ top: `${top}px`, height: `${height}px` }} onClick={onOpen} title={`${item.title} · ${timeLabel(item.start)}–${timeLabel(item.end)} · ${item.status === 'created' ? 'Reunión de Teams creada' : item.status === 'ready' ? 'Lista para crear en Teams' : 'Borrador / pendiente de revisar'}`}><strong>{item.title}</strong><span>{timeLabel(item.start)}–{timeLabel(item.end)}</span><small>{item.status === 'created' ? 'Teams creado' : item.status === 'ready' ? 'Lista para Teams' : 'Borrador'}</small></button>
}

function GraphBlock({ item }: { item: GraphEvent }) {
  const start = graphLocalTime(item.start.dateTime)
  const end = graphLocalTime(item.end.dateTime)
  const top = ((minutesOf(start) - 420) / 60) * 36
  const height = Math.max(30, ((minutesOf(end) - minutesOf(start)) / 60) * 36 - 4)
  return <a className="calendar-block graph-block" style={{ top: `${top}px`, height: `${height}px` }} href={item.webLink} target="_blank" rel="noreferrer"><strong>{item.subject || 'Evento de Teams'}</strong><span>{start}–{end}</span><small>Publicado en este curso</small></a>
}

function CompactTeacherMeetingTable({
  meetings,
  onDelete,
  deletingSessionId,
}: {
  meetings: GraphEvent[]
  onDelete?: (meeting: GraphEvent) => void
  deletingSessionId?: string
}) {
  return (
    <div className="zoom-table-wrap">
      <table className="zoom-meeting-table">
        <thead><tr><th>Día y horario</th><th>Tema</th><th>ID de la reunión</th><th>Acciones</th></tr></thead>
        <tbody>{meetings.length ? meetings.map((meeting) => (
          <tr key={meeting.id}>
            <td>{new Intl.DateTimeFormat('es-MX', { weekday: 'short', day: '2-digit', month: 'short', year: 'numeric' }).format(graphDate(meeting.start.dateTime))} · {graphLocalTime(meeting.start.dateTime)}–{graphLocalTime(meeting.end.dateTime)}</td>
            <td>{meeting.subject || 'Reunión de Teams sin título'}</td>
            <td><code>{meeting.id}</code></td>
            <td className="zoom-table-actions">
              {meeting.onlineMeeting?.joinUrl
                ? <a className="zoom-open-link" href={meeting.onlineMeeting.joinUrl} target="_blank" rel="noreferrer">Abrir</a>
                : <span>—</span>}
              {onDelete && <button
                className="delete-meeting-button"
                type="button"
                onClick={() => onDelete(meeting)}
                disabled={deletingSessionId !== undefined && deletingSessionId !== ''}
                aria-label={`Eliminar ${meeting.subject || 'reunión'} del curso`}
                title="Eliminar del curso"
              ><Trash2 size={14} /><span>Eliminar</span></button>}
            </td>
          </tr>
        )) : <tr><td colSpan={4} className="zoom-table-empty">Sin datos</td></tr>}</tbody>
      </table>
    </div>
  )
}

function DraftListItem({ item, onOpen, statusClass }: { item: SessionDraft; onOpen: () => void; statusClass: string }) {
  return <article className={`session-list-card ${statusClass}`}><span className="list-time">{timeLabel(item.start)}<small>{timeLabel(item.end)}</small></span><span className="list-card-copy"><strong>{item.title}</strong><small>{item.recurrence ? 'Horario recurrente' : 'Reunión puntual'} · {item.location || 'En línea'}</small></span><StatusBadge status={item.status} /><button className="row-action" onClick={onOpen} aria-label="Revisar sesión"><ChevronRight size={17} /></button></article>
}

function GraphListItem({ item }: { item: GraphEvent }) {
  const start = graphDate(item.start.dateTime)
  return <article className="session-list-card graph-list-card"><span className="list-time">{graphLocalTime(item.start.dateTime)}<small>{formatDay(start)}</small></span><span className="list-card-copy"><strong>{item.subject || 'Reunión de Teams'}</strong><small>Publicada para el curso actual</small></span>{item.onlineMeeting?.joinUrl ? <a className="row-action join-action" href={item.onlineMeeting.joinUrl} target="_blank" rel="noreferrer" aria-label="Unirse a reunión"><Video size={15} /></a> : <span className="source-badge">Teams</span>}</article>
}

function SessionRow({ item, onReview, onEdit, onCreate, statusClass }: { item: SessionDraft; onReview: () => void; onEdit: () => void; onCreate: () => void; statusClass: string }) {
  const actionLabel = item.status === 'created' ? 'Ver Teams' : item.status === 'ready' ? 'Confirmar creación' : 'Revisar'
  return <article className={`session-row ${statusClass}`}><div className="session-date"><strong>{formatDay(new Date(`${item.date}T12:00:00`), { day: '2-digit' })}</strong><span>{formatDay(new Date(`${item.date}T12:00:00`), { month: 'short' })}</span></div><div className="session-row-main"><strong>{item.title}</strong><small>{item.recurrence ? 'Recurrente' : 'Puntual'} · {timeLabel(item.start)}–{timeLabel(item.end)} · {item.location || 'En línea'}</small></div><StatusBadge status={item.status} /><button className="edit-button" onClick={item.status === 'created' ? onReview : onEdit} aria-label="Editar o ver sesión" disabled={item.status === 'created'}><Pencil size={14} /></button><button className={item.status === 'created' ? 'row-link created-link' : 'row-link'} onClick={item.status === 'created' ? onReview : onCreate}>{actionLabel}<ArrowRight size={13} /></button></article>
}

function StatusBadge({ status }: { status: DraftStatus }) {
  const labels = { review: 'Pendiente de revisar', ready: 'Lista para crear en Teams', created: 'Reunión de Teams creada' }
  return <span className={`status-badge ${status}`}><i />{labels[status]}</span>
}

function Composer(props: {
  mode: FormMode; onModeChange: (mode: FormMode) => void; formTitle: string; setTitle: (value: string) => void; start: string; setStart: (value: string) => void; end: string; setEnd: (value: string) => void; date: string; setDate: (value: string) => void; termEnd: string; setTermEnd: (value: string) => void; selectedDays: number[]; setSelectedDays: (value: number[]) => void; timeZone: string; setTimeZone: (value: string) => void; attendees: string; setAttendees: (value: string) => void; description: string; setDescription: (value: string) => void; location: string; setLocation: (value: string) => void; editing: boolean; error: string; onClose: () => void; onSubmit: (event: FormEvent<HTMLFormElement>) => void
}) {
  const isWeekly = props.mode === 'weekly'
  const occurrenceCount = useMemo(() => {
    if (!isWeekly || !props.selectedDays.length || !props.date || !props.termEnd) return 0
    const from = new Date(`${props.date}T12:00:00`)
    const to = new Date(`${props.termEnd}T12:00:00`)
    let count = 0
    for (let cursor = from; cursor <= to && count < 120; cursor = addDays(cursor, 1)) if (props.selectedDays.includes(cursor.getDay() || 7)) count++
    return count
  }, [isWeekly, props.date, props.termEnd, props.selectedDays])
  return <div className="modal-backdrop" role="presentation" onMouseDown={(event) => { if (event.target === event.currentTarget) props.onClose() }}><section className="meeting-modal composer-modal" role="dialog" aria-modal="true" aria-labelledby="composer-title">
    <header className="modal-header"><div><div className="eyebrow">{isWeekly ? 'CRONOGRAMA DEL CURSO' : 'ENCUENTRO PUNTUAL'}</div><h2 id="composer-title">{props.editing ? 'Editar sesión' : isWeekly ? 'Crear horario semanal' : 'Programar reunión única'}</h2></div><button className="icon-button" onClick={props.onClose} aria-label="Cerrar"><X size={19} /></button></header>
    <form onSubmit={props.onSubmit}>
      {!props.editing && <fieldset className="composer-mode-switch">
        <legend>Tipo de sesión</legend>
        <button type="button" aria-pressed={!isWeekly} className={!isWeekly ? 'composer-mode-option selected' : 'composer-mode-option'} onClick={() => props.onModeChange('single')}>Reunión única</button>
        <button type="button" aria-pressed={isWeekly} className={isWeekly ? 'composer-mode-option selected' : 'composer-mode-option'} onClick={() => props.onModeChange('weekly')}>Sesiones semanales</button>
      </fieldset>}
      <label className="field-label" htmlFor="session-title">Tema de la sesión</label><input id="session-title" required maxLength={255} value={props.formTitle} onChange={(event) => props.setTitle(event.target.value)} placeholder="Ej. Seminario de investigación" />
      {isWeekly && <><span className="field-label">Días de clase</span><div className="weekday-picker">{weekdays.map((day) => <button type="button" key={day.value} className={props.selectedDays.includes(day.value) ? 'weekday-chip selected' : 'weekday-chip'} aria-pressed={props.selectedDays.includes(day.value)} onClick={() => props.setSelectedDays(props.selectedDays.includes(day.value) ? props.selectedDays.filter((value) => value !== day.value) : [...props.selectedDays, day.value].sort())}>{day.short}</button>)}</div></>}
      <div className="form-row"><div><label className="field-label" htmlFor="session-date">{isWeekly ? 'Inicia el período' : 'Fecha'}</label><input id="session-date" type="date" required value={props.date} onChange={(event) => props.setDate(event.target.value)} /></div>{isWeekly ? <div><label className="field-label" htmlFor="term-end">Termina el período</label><input id="term-end" type="date" min={props.date} required value={props.termEnd} onChange={(event) => props.setTermEnd(event.target.value)} /></div> : <div><label className="field-label" htmlFor="session-location">Ubicación <span>(opcional)</span></label><input id="session-location" value={props.location} onChange={(event) => props.setLocation(event.target.value)} placeholder="Aula o en línea" /></div>}</div>
      <div className="form-row"><div><label className="field-label" htmlFor="session-start">Hora de inicio</label><input id="session-start" type="time" required value={props.start} onChange={(event) => props.setStart(event.target.value)} /></div><div><label className="field-label" htmlFor="session-end">Hora de término</label><input id="session-end" type="time" required value={props.end} onChange={(event) => props.setEnd(event.target.value)} /></div></div>
      <label className="field-label" htmlFor="session-zone">Zona horaria del evento</label><select id="session-zone" value={props.timeZone} onChange={(event) => props.setTimeZone(event.target.value)}>{timeZones.map((zone) => <option key={zone}>{zone}</option>)}</select>
      {isWeekly &&       <div className="period-preview"><CalendarDays size={16} /><span><strong>Vista previa de sesiones</strong><small>{occurrenceCount} sesión{occurrenceCount === 1 ? '' : 'es'} en el período. Se crearán individualmente en Teams y se publicarán para los estudiantes del curso.</small></span></div>}
      {!isWeekly && <><label className="field-label" htmlFor="session-attendees">Invitar por correo <span>(opcional; separado por comas)</span></label><div className="input-with-icon"><Mail size={15} /><input id="session-attendees" value={props.attendees} onChange={(event) => props.setAttendees(event.target.value)} placeholder="docente@institucion.edu" /></div></>}
      <label className="field-label" htmlFor="session-description">Descripción <span>(opcional)</span></label><textarea id="session-description" rows={2} value={props.description} onChange={(event) => props.setDescription(event.target.value)} placeholder="Objetivo, materiales o indicaciones" />
      {isWeekly && <><label className="field-label" htmlFor="session-location-weekly">Ubicación <span>(opcional)</span></label><input id="session-location-weekly" value={props.location} onChange={(event) => props.setLocation(event.target.value)} placeholder="Aula o en línea" /></>}
      <p className="form-note"><ShieldCheck size={15} />{isWeekly
        ? 'Al guardar, se creará una reunión de Teams por cada día seleccionado y se publicará automáticamente para este curso.'
        : 'Al continuar revisarás los detalles y confirmarás antes de crear el evento de Teams.'}</p>
      {props.error && <ErrorBanner message={props.error} />}
      <div className="modal-actions"><button type="button" className="cancel-button" onClick={props.onClose}>Cancelar</button><button type="submit" className="primary-button"><Check size={16} />{props.editing ? 'Guardar cambios' : isWeekly ? `Crear y publicar ${occurrenceCount} sesiones` : 'Revisar reunión'}</button></div>
    </form>
  </section></div>
}

function ReviewDialog({ item, error, saving, connected, canConnect, onConnect, onClose, onEdit, onMarkReady, onCreate }: { item: SessionDraft; error: string; saving: boolean; connected: boolean; canConnect: boolean; onConnect: () => void; onClose: () => void; onEdit: () => void; onMarkReady: () => void; onCreate: () => void }) {
  const date = new Date(`${item.date}T12:00:00`)
  const created = item.status === 'created'
  return <div className="modal-backdrop" role="presentation" onMouseDown={(event) => { if (event.target === event.currentTarget) onClose() }}><section className="meeting-modal review-modal" role="dialog" aria-modal="true" aria-labelledby="review-title">
    <header className="modal-header"><div><div className="eyebrow">{created ? 'EVENTO DE MICROSOFT GRAPH' : 'REVISIÓN DE SESIÓN'}</div><h2 id="review-title">{created ? 'Reunión creada' : 'Confirma los detalles'}</h2></div><button className="icon-button" onClick={onClose} aria-label="Cerrar"><X size={19} /></button></header>
    <StatusBadge status={item.status} />
    <div className="review-summary"><h3>{item.title}</h3><div><CalendarDays size={15} />{formatLongDay(date)} · {timeLabel(item.start)}–{timeLabel(item.end)}</div><div><Clock3 size={15} />{item.timeZone}</div>{item.location && <div><FolderOpen size={15} />{item.location}</div>}{item.attendees.length > 0 && <div><Mail size={15} />{item.attendees.join(', ')}</div>}{item.description && <p>{item.description}</p>}</div>
    {created && item.graphEvent?.onlineMeeting?.joinUrl && <a className="hero-join review-join" href={item.graphEvent.onlineMeeting.joinUrl} target="_blank" rel="noreferrer"><Video size={15} />Abrir reunión de Teams<ArrowRight size={14} /></a>}
    {!created && <p className="review-warning"><ShieldCheck size={16} />Aún no se ha publicado. Al confirmar, se creará un evento real en tu calendario Microsoft 365 con reunión Teams.</p>}
    {error && <ErrorBanner message={error} />}
    <div className="modal-actions review-actions">
      {!created && <button className="cancel-button" onClick={onEdit}><Pencil size={14} />Editar</button>}
      {!created && item.status === 'review' && <button className="secondary-button" onClick={onMarkReady}>Marcar lista</button>}
      {!created && item.status === 'ready' && (connected
        ? <button className="primary-button" onClick={onCreate} disabled={saving}>{saving ? <LoaderCircle size={16} className="spin" /> : <Video size={16} />}{saving ? 'Creando…' : 'Confirmar y crear en Teams'}</button>
        : <button className="secondary-button" onClick={onConnect} disabled={!canConnect || saving}><Video size={15} />{canConnect ? 'Conectar Microsoft' : 'Configura Microsoft primero'}</button>)}
      {created && item.graphEvent?.webLink && <a className="row-link" href={item.graphEvent.webLink} target="_blank" rel="noreferrer">Ver evento <ArrowRight size={13} /></a>}
    </div>
  </section></div>
}

function CourseRecordingRow({ recording }: { recording: CourseRecording }) {
  return <article className="recording-row">
    <span className="recording-thumb video-thumb"><FileVideo2 size={20} /><span className="thumb-play"><Play size={10} fill="currentColor" /></span></span>
    <div className="recording-info">
      <span className="recording-name">{recording.title}</span>
      <span>{recording.createdAt ? formatDay(new Date(recording.createdAt), { year: 'numeric' }) : 'Fecha no disponible'}</span>
    </div>
    <a className="download-link" href={recording.webUrl} target="_blank" rel="noreferrer" aria-label={`Abrir ${recording.title}`} title="Ver grabación en OneDrive">
      <ArrowRight size={16} />
    </a>
  </article>
}

type FeedState<T> =
  | { state: 'loading' }
  | { state: 'ready'; items: T[] }
  | { state: 'unavailable' }
  | { state: 'forbidden' }
  | { state: 'error'; message: string }

function isObject(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function secureWebUrl(value: unknown): string | undefined {
  if (typeof value !== 'string') return undefined
  try {
    const url = new URL(value)
    return url.protocol === 'https:' ? url.href : undefined
  } catch {
    return undefined
  }
}

function learnerScheduleLabel(session: LearnerSession): string {
  const timeZone = displayTimeZones[session.timeZone] ?? session.timeZone
  const start = new Date(session.start)
  const end = new Date(session.end)
  const startOptions: Intl.DateTimeFormatOptions = {
    weekday: 'long', day: '2-digit', month: 'short', year: 'numeric',
    hour: '2-digit', minute: '2-digit', timeZone,
  }
  const endOptions: Intl.DateTimeFormatOptions = { hour: '2-digit', minute: '2-digit', timeZone }
  let startLabel: string
  let endLabel: string
  try {
    startLabel = new Intl.DateTimeFormat('es-MX', startOptions).format(start)
    endLabel = new Intl.DateTimeFormat('es-MX', endOptions).format(end)
  } catch {
    startLabel = new Intl.DateTimeFormat('es-MX', {
      weekday: 'long', day: '2-digit', month: 'short', year: 'numeric',
      hour: '2-digit', minute: '2-digit',
    }).format(start)
    endLabel = new Intl.DateTimeFormat('es-MX', { hour: '2-digit', minute: '2-digit' }).format(end)
  }
  return `${startLabel} – ${endLabel}`
}

function LearnerWorkspace({ view, setView }: {
  view: View
  setView: (view: View) => void
}) {
  const [sessionFeed, setSessionFeed] = useState<FeedState<LearnerSession>>({ state: 'loading' })
  const [recordingFeed, setRecordingFeed] = useState<FeedState<CourseRecording>>({ state: 'loading' })
  const [meetingPeriod, setMeetingPeriod] = useState<MeetingPeriod>('upcoming')
  const feedView = view === 'summary' ? 'schedule' : view

  useEffect(() => {
    let alive = true
    let loadingRequest = false
    const endpoint = feedView === 'schedule' ? '/api/course/sessions' : '/api/course/recordings'
    async function load(silent = false) {
      if (loadingRequest) return
      loadingRequest = true
      const setStatus = (state: Exclude<FeedState<never>, { state: 'ready' }>) => {
        if (feedView === 'schedule') setSessionFeed(state)
        else setRecordingFeed(state)
      }
      if (!silent) setStatus({ state: 'loading' })
      try {
        const response = await fetch(endpoint, {
          credentials: 'same-origin',
          headers: { Accept: 'application/json' },
        })
        if (!alive) return
        if (response.status === 401) {
          setStatus({ state: 'unavailable' })
          return
        }
        if (response.status === 403) {
          setStatus({ state: 'forbidden' })
          return
        }
        if (response.status === 404 || response.status === 501 || response.status === 503) {
          setStatus({ state: 'unavailable' })
          return
        }
        if (!response.ok) {
          setStatus({ state: 'error', message: `El servicio respondió ${response.status}.` })
          return
        }
        const data: unknown = await response.json()
        if (!alive) return
        if (!isObject(data)) {
          setStatus({ state: 'error', message: 'La respuesta del servicio no tiene el formato esperado.' })
          return
        }
        if (feedView === 'schedule') {
          const rawItems = Array.isArray(data.sessions) ? data.sessions : null
          if (!rawItems) {
            setStatus({ state: 'error', message: 'La respuesta de sesiones no tiene el formato esperado.' })
            return
          }
          const items = rawItems.flatMap((raw): LearnerSession[] => {
            if (!isObject(raw) || typeof raw.id !== 'string'
              || typeof raw.start !== 'string' || !Number.isFinite(Date.parse(raw.start))
              || typeof raw.end !== 'string' || !Number.isFinite(Date.parse(raw.end))
              || Date.parse(raw.end) <= Date.parse(raw.start)
              || typeof raw.timeZone !== 'string') return []
            return [{
              id: raw.id,
              start: raw.start,
              end: raw.end,
              timeZone: raw.timeZone,
              topic: typeof raw.topic === 'string' ? raw.topic : typeof raw.title === 'string' ? raw.title : undefined,
              joinUrl: secureWebUrl(raw.joinUrl),
            }]
          }).sort((a, b) => Date.parse(a.start) - Date.parse(b.start))
          if (items.length !== rawItems.length) {
            setStatus({ state: 'error', message: 'La respuesta contiene sesiones incompletas o no válidas.' })
            return
          }
          setSessionFeed({ state: 'ready', items })
        } else {
          const rawItems = Array.isArray(data.recordings) ? data.recordings : null
          if (!rawItems) {
            setStatus({ state: 'error', message: 'La respuesta de grabaciones no tiene el formato esperado.' })
            return
          }
          const items = rawItems.flatMap((raw): CourseRecording[] => {
            const webUrl = isObject(raw) ? secureWebUrl(raw.webUrl) : undefined
            if (!isObject(raw) || typeof raw.id !== 'string' || typeof raw.title !== 'string' || !webUrl) return []
            return [{
              id: raw.id,
              title: raw.title,
              createdAt: typeof raw.createdAt === 'string' && Number.isFinite(Date.parse(raw.createdAt)) ? raw.createdAt : undefined,
              webUrl,
            }]
          }).sort((a, b) => (b.createdAt ?? '').localeCompare(a.createdAt ?? ''))
          if (items.length !== rawItems.length) {
            setStatus({ state: 'error', message: 'La respuesta contiene grabaciones incompletas o no válidas.' })
            return
          }
          setRecordingFeed({ state: 'ready', items })
        }
      } catch (cause) {
        if (alive) setStatus({ state: 'unavailable' })
      } finally {
        loadingRequest = false
      }
    }
    void load()
    const refreshFeed = () => {
      if (document.visibilityState === 'visible') void load(true)
    }
    const interval = window.setInterval(refreshFeed, feedView === 'schedule' ? 5000 : 30_000)
    window.addEventListener('focus', refreshFeed)
    document.addEventListener('visibilitychange', refreshFeed)
    return () => {
      alive = false
      window.clearInterval(interval)
      window.removeEventListener('focus', refreshFeed)
      document.removeEventListener('visibilitychange', refreshFeed)
    }
  }, [feedView])

  const feed = view === 'recordings' ? recordingFeed : sessionFeed
  const allSessions = view !== 'recordings' && feed.state === 'ready'
    ? (feed as FeedState<LearnerSession> & { state: 'ready' }).items
    : []
  const sessions = allSessions.filter((session) => meetingPeriod === 'past'
    ? Date.parse(session.end) < Date.now()
    : Date.parse(session.end) >= Date.now())
  const courseRecordings = view === 'recordings' && feed.state === 'ready'
    ? (feed as FeedState<CourseRecording> & { state: 'ready' }).items
    : []

  return (
    <div className="learner-shell">
      <header className="learner-topbar">
        <a className="zoom-wordmark" href="#" onClick={(event) => event.preventDefault()} aria-label="Microsoft Teams">
          Microsoft Teams
        </a>
        <nav className="zoom-shortcuts" aria-label="Navegación principal">
          <button onClick={() => setView('schedule')}>Inicio</button>
          <button onClick={() => setView('schedule')}>Citas</button>
        </nav>
      </header>
      <main className="learner-content">
        <div className="zoom-timezone">Su idioma y su zona horaria actual son ({new Intl.DateTimeFormat('es-CO', { timeZoneName: 'short' }).format(new Date()).split(' ').pop()}), Español</div>
        <nav className="learner-tabs" aria-label="Contenido del curso">
          <button className={view === 'schedule' ? 'learner-tab active' : 'learner-tab'} onClick={() => setView('schedule')}>Reuniones</button>
          <button className={view === 'recordings' ? 'learner-tab active' : 'learner-tab'} onClick={() => setView('recordings')}>Grabaciones en la nube</button>
          <button className={view === 'summary' ? 'learner-tab active' : 'learner-tab'} onClick={() => setView('summary')}>Resumen de reuniones</button>
        </nav>
        {(view === 'schedule' || view === 'summary') && <nav className="zoom-subtabs" aria-label="Filtro de reuniones">
          <button className={meetingPeriod === 'upcoming' ? 'zoom-subtab active' : 'zoom-subtab'} onClick={() => setMeetingPeriod('upcoming')}>Próximas reuniones</button>
          <button className={meetingPeriod === 'past' ? 'zoom-subtab active' : 'zoom-subtab'} onClick={() => setMeetingPeriod('past')}>Reuniones pasadas</button>
        </nav>}
        <section className="learner-panel">
          <LearnerFeedState feed={feed} view={view}>
            {view === 'schedule' && feed.state === 'ready' && (
              sessions.length
                ? <LearnerSessionTable sessions={sessions} />
                : <LearnerSessionTable sessions={[]} />
            )}
            {view === 'summary' && feed.state === 'ready' && (
              <LearnerSessionTable sessions={sessions} />
            )}
            {view === 'recordings' && feed.state === 'ready' && (
              courseRecordings.length
                ? <div className="learner-recording-list">{courseRecordings.map((recording) => <LearnerRecordingRow key={recording.id} recording={recording} />)}</div>
                : <LearnerEmpty icon={<FolderOpen size={19} />} title="Sin datos" detail="Todavía no hay grabaciones compartidas para este curso." />
            )}
          </LearnerFeedState>
        </section>
      </main>
    </div>
  )
}

function LearnerFeedState({ feed, view, children }: {
  feed: FeedState<LearnerSession> | FeedState<CourseRecording>
  view: View
  children: ReactNode
}) {
  if (feed.state === 'loading') return <div className="learner-state"><LoaderCircle className="spin" size={21} /><span>Consultando contenido publicado del curso…</span></div>
  if (feed.state === 'unavailable') return <LearnerEmpty icon={<AlertCircle size={19} />} title="Contenido del curso no disponible" detail={`El servicio de ${view === 'recordings' ? 'grabaciones' : 'reuniones'} del curso no está conectado o no responde. No se consultan calendarios ni archivos personales.`} kind="notice" />
  if (feed.state === 'forbidden') return <LearnerEmpty icon={<ShieldCheck size={19} />} title="Acceso no autorizado" detail="El servidor no autorizó este contenido para el curso de la sesión actual." kind="notice" />
  if (feed.state === 'error') return <LearnerEmpty icon={<AlertCircle size={19} />} title="No se pudo cargar el contenido" detail={feed.message} kind="error" />
  return <>{children}</>
}

function LearnerSessionTable({ sessions }: { sessions: LearnerSession[] }) {
  return (
    <div className="learner-meeting-table-wrap">
      <table className="learner-meeting-table">
        <thead><tr><th>Día y horario</th><th>Tema</th><th>ID de la reunión</th><th>Acciones</th></tr></thead>
        <tbody>{sessions.length ? sessions.map((session) => (
          <tr key={session.id}>
            <td data-label="Día y horario">{learnerScheduleLabel(session)}</td>
            <td data-label="Tema">{session.topic || 'Reunión de Teams'}</td>
            <td data-label="ID de reunión"><code>{session.id}</code></td>
            <td data-label="Acciones">{session.joinUrl
              ? <a className="learner-table-action" href={session.joinUrl} target="_blank" rel="noreferrer">Unirse <ArrowRight size={13} /></a>
              : <span className="join-unavailable">No disponible</span>}</td>
          </tr>
        )) : <tr><td colSpan={4} className="learner-table-empty">Sin datos</td></tr>}</tbody>
      </table>
    </div>
  )
}

function LearnerRecordingRow({ recording }: { recording: CourseRecording }) {
  return (
    <article className="learner-recording-row">
      <span className="learner-session-icon"><FileVideo2 size={18} /></span>
      <div className="learner-session-copy">
        <strong>{recording.title}</strong>
        <span>{recording.createdAt ? formatDay(new Date(recording.createdAt), { year: 'numeric' }) : 'Fecha no disponible'}</span>
      </div>
      <a className="secondary-button learner-watch-button" href={recording.webUrl} target="_blank" rel="noreferrer"><Play size={14} />Ver grabación</a>
    </article>
  )
}

function LearnerEmpty({ icon, title, detail, kind = 'empty' }: {
  icon: ReactNode
  title: string
  detail: string
  kind?: 'empty' | 'notice' | 'error'
}) {
  return <div className={`learner-empty ${kind}`}><span>{icon}</span><strong>{title}</strong><p>{detail}</p></div>
}

function LtiLoading() {
  return <div className="learner-shell"><main className="learner-content"><div className="learner-state"><LoaderCircle className="spin" size={21} /><span>Verificando la sesión institucional de Brightspace…</span></div></main></div>
}

function AccessDenied() {
  return (
    <div className="learner-shell">
      <header className="learner-topbar"><span className="brand-name">Microsoft Teams<span>.</span></span></header>
      <main className="learner-content"><LearnerEmpty kind="notice" icon={<ShieldCheck size={20} />} title="No se pudo validar el rol del curso" detail="No hay una sesión institucional con rol y contexto de curso verificados. Vuelve a abrir esta herramienta desde Brightspace o contacta al administrador." /></main>
    </div>
  )
}

function ErrorBanner({ message }: { message: string }) {
  return <div className="error-banner" role="alert"><AlertCircle size={16} /><span>{message}</span></div>
}

export default App