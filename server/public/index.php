<?php

declare(strict_types=1);

use GuzzleHttp\Client;
use Packback\Lti1p3\Factories\MessageFactory;
use Packback\Lti1p3\JwksEndpoint;
use Packback\Lti1p3\LtiOidcLogin;
use Packback\Lti1p3\LtiServiceConnector;
use Packback\Lti1p3\LtiException;
use Packback\Lti1p3\Messages\ResourceLinkRequest;
use Packback\Lti1p3\OidcException;
use TeamsBrightspace\Config;
use TeamsBrightspace\CourseRecordingRepository;
use TeamsBrightspace\CourseSessionRepository;
use TeamsBrightspace\HttpException;
use TeamsBrightspace\LtiCookie;
use TeamsBrightspace\LtiRegistrationRepository;
use TeamsBrightspace\LtiSession;

require dirname(__DIR__).'/vendor/autoload.php';

function respondJson(int $status, array $payload): never
{
    http_response_code($status);
    header('Content-Type: application/json; charset=utf-8');
    header('Cache-Control: no-store, private');
    echo json_encode($payload, JSON_UNESCAPED_SLASHES | JSON_UNESCAPED_UNICODE | JSON_THROW_ON_ERROR);
    exit;
}

function requireMethod(string $method): void
{
    if (($_SERVER['REQUEST_METHOD'] ?? '') !== $method) {
        header('Allow: '.$method);
        throw new HttpException(405, 'Método no permitido.');
    }
}

function getOrigin(string $url): string
{
    $parts = parse_url($url);
    if (!is_array($parts) || !isset($parts['scheme'], $parts['host'])) {
        throw new RuntimeException('APP_BASE_URL debe ser una URL absoluta.');
    }

    return $parts['scheme'].'://'.$parts['host'].(isset($parts['port']) ? ':'.$parts['port'] : '');
}

function normalizedRole(array $roles): ?string
{
    $learnerFound = false;
    $allowedRoles = [
        'http://purl.imsglobal.org/vocab/lis/v2/membership#instructor' => 'instructor',
        'http://purl.imsglobal.org/vocab/lis/v2/membership#learner' => 'learner',
    ];
    foreach ($roles as $role) {
        if (!is_string($role)) {
            continue;
        }
        $normalized = strtolower($role);
        if ($normalized === 'instructor' || ($allowedRoles[$normalized] ?? null) === 'instructor') {
            return 'instructor';
        }
        if ($normalized === 'learner' || ($allowedRoles[$normalized] ?? null) === 'learner') {
            $learnerFound = true;
        }
    }

    return $learnerFound ? 'learner' : null;
}

function claimValue(array $claims, string $key): string
{
    $value = $claims[$key] ?? '';
    return is_string($value) ? trim($value) : '';
}

function roleForInstitutionalEmail(string $email): ?string
{
    $email = strtolower(trim($email));
    if (filter_var($email, FILTER_VALIDATE_EMAIL) === false) {
        return null;
    }

    if (str_ends_with($email, '@estud.tesa.edu.ec')) {
        return 'learner';
    }
    if (str_ends_with($email, '@tesa.edu.ec')) {
        return 'instructor';
    }

    return null;
}

try {
    $config = new Config();
    $baseUrl = $config->appBaseUrl();
    $basePath = rtrim((string) (parse_url($baseUrl, PHP_URL_PATH) ?: ''), '/');
    $requestPath = (string) (parse_url($_SERVER['REQUEST_URI'] ?? '/', PHP_URL_PATH) ?: '/');
    if ($basePath !== '' && ($requestPath === $basePath || str_starts_with($requestPath, $basePath.'/'))) {
        $requestPath = substr($requestPath, strlen($basePath)) ?: '/';
    }
    $requestPath = '/'.trim($requestPath, '/');
    if ($requestPath !== '/lti/jwks') {
        session_name('teams_lti_session');
        ini_set('session.use_strict_mode', '1');
        ini_set('session.use_only_cookies', '1');
        ini_set('session.use_trans_sid', '0');
        session_set_cookie_params([
            'lifetime' => 0,
            'path' => $basePath === '' ? '/' : $basePath.'/',
            'secure' => parse_url($baseUrl, PHP_URL_SCHEME) === 'https',
            'httponly' => true,
            'samesite' => 'None',
        ]);
        if (!session_start()) {
            throw new RuntimeException('No se pudo iniciar la sesión del servidor.');
        }
    }
    $registration = new LtiRegistrationRepository($config);

    if ($requestPath === '/lti/jwks') {
        if ($config->ltiRegistration()['toolKeyId'] === '') {
            throw new HttpException(503, 'El proveedor LTI aún no está configurado.');
        }
        $jwks = JwksEndpoint::fromRegistration(
            $registration->findRegistrationByIssuer($config->ltiRegistration()['issuer'])
                ?? throw new HttpException(503, 'El proveedor LTI aún no está configurado.')
        );
        respondJson(200, $jwks->getPublicJwks());
    }

    if ($requestPath === '/lti/login') {
        $method = $_SERVER['REQUEST_METHOD'] ?? '';
        if (!in_array($method, ['GET', 'POST'], true)) {
            header('Allow: GET, POST');
            throw new HttpException(405, 'Método no permitido.');
        }
        if (!$config->ltiReady()) {
            throw new HttpException(503, 'El inicio de Brightspace aún no está configurado en el servidor.');
        }
        $loginRequest = $method === 'POST' ? $_POST : $_GET;
        $expectedTarget = rtrim($baseUrl, '/').'/';
        $providedTarget = $loginRequest['target_link_uri'] ?? null;
        if (!is_string($providedTarget) || rtrim($providedTarget, '/') !== rtrim($expectedTarget, '/')) {
            throw new HttpException(400, 'El enlace de inicio no coincide con la aplicación registrada.');
        }
        $providedIssuer = $loginRequest['iss'] ?? null;
        $providedClientId = $loginRequest['client_id'] ?? null;
        $settings = $config->ltiRegistration();
        if (!is_string($providedIssuer) || !hash_equals($settings['issuer'], $providedIssuer)
            || !is_string($providedClientId) || !hash_equals($settings['clientId'], $providedClientId)) {
            throw new HttpException(400, 'La plataforma de inicio no coincide con el registro institucional.');
        }

        $cookie = new LtiCookie();
        $cache = new LtiSession();
        $login = new LtiOidcLogin($registration, $cache, $cookie);
        $redirectUrl = $login->getRedirectUrl($baseUrl.'/lti/launch', $loginRequest);
        header('Location: '.$redirectUrl, true, 302);
        exit;
    }

    if ($requestPath === '/lti/launch') {
        requireMethod('POST');
        unset($_SESSION['lti_user']);
        if (!$config->ltiReady()) {
            throw new HttpException(503, 'El inicio de Brightspace aún no está configurado en el servidor.');
        }
        if (!isset($_POST['state'], $_POST['id_token'])
            || !is_string($_POST['state']) || $_POST['state'] === ''
            || !is_string($_POST['id_token']) || $_POST['id_token'] === '') {
            throw new HttpException(400, 'El lanzamiento de Brightspace está incompleto.');
        }

        $cache = new LtiSession();
        $factory = new MessageFactory(
            $registration,
            new LtiServiceConnector($cache, new Client()),
            $cache,
            new LtiCookie()
        );
        $message = $factory->create($_POST);
        if (!$message instanceof ResourceLinkRequest) {
            throw new HttpException(403, 'Este tipo de lanzamiento de Brightspace no está habilitado.');
        }

        $body = $message->getBody();
        if (($body['iss'] ?? null) !== $config->ltiRegistration()['issuer']
            || rtrim((string) ($body['https://purl.imsglobal.org/spec/lti/claim/target_link_uri'] ?? ''), '/')
                !== rtrim($baseUrl.'/', '/')) {
            throw new HttpException(403, 'El lanzamiento no pertenece a esta aplicación.');
        }

        $context = $body['https://purl.imsglobal.org/spec/lti/claim/context'] ?? null;
        $roles = $body['https://purl.imsglobal.org/spec/lti/claim/roles'] ?? null;
        $ltiRole = is_array($roles) ? normalizedRole($roles) : null;
        $email = claimValue($body, 'email');
        $emailVerified = $body['email_verified'] ?? null;
        $role = roleForInstitutionalEmail($email);
        if (!is_array($context) || !is_string($context['id'] ?? null)
            || $ltiRole === null || $role === null
            || $emailVerified === false
            || ($role === 'instructor' && $ltiRole !== 'instructor')) {
            throw new HttpException(403, 'El lanzamiento no contiene un curso o un rol autorizado.');
        }

        $custom = $body['https://purl.imsglobal.org/spec/lti/claim/custom'] ?? [];
        if (!is_array($custom)) {
            $custom = [];
        }
        $courseCode = claimValue($context, 'label') ?: claimValue($context, 'id');
        $course = [
            'courseName' => claimValue($context, 'title') ?: claimValue($context, 'label'),
            'courseCode' => $courseCode,
            'academicPeriod' => claimValue($custom, $config->coursePeriodClaim()),
            'nrc' => claimValue($custom, $config->courseNrcClaim()) ?: $courseCode,
            'contextId' => trim($context['id']),
        ];
        if ($role === 'instructor' && isset($body['name']) && is_string($body['name'])) {
            $course['instructorName'] = $body['name'];
        }

        session_regenerate_id(true);
        $_SESSION['lti_user'] = [
            'role' => $role,
            'course' => $course,
            'expires_at' => time() + 14_400,
        ];
        header('Location: '.$baseUrl.'/', true, 303);
        exit;
    }

    if ($requestPath === '/api/lti/session') {
        requireMethod('GET');
        $user = $_SESSION['lti_user'] ?? null;
        if (!is_array($user) || ($user['expires_at'] ?? 0) < time()
            || !is_array($user['course'] ?? null)
            || !in_array($user['role'] ?? null, ['instructor', 'learner'], true)) {
            respondJson(200, ['authenticated' => false]);
        }

        respondJson(200, [
            'authenticated' => true,
            'role' => $user['role'],
            'course' => $user['course'],
        ]);
    }

    $sessionDeleteMatch = [];
    $isSessionDeleteRoute = preg_match(
        '#^/api/course/sessions/([a-f0-9]{32})$#D',
        $requestPath,
        $sessionDeleteMatch
    ) === 1;
    if ($requestPath === '/api/course/sessions' || $isSessionDeleteRoute) {
        $method = $_SERVER['REQUEST_METHOD'] ?? '';
        $sessionId = $sessionDeleteMatch[1] ?? null;
        $allowedMethods = $sessionId === null ? ['GET', 'POST'] : ['DELETE'];
        if (!in_array($method, $allowedMethods, true)) {
            header('Allow: '.implode(', ', $allowedMethods));
            throw new HttpException(405, 'Método no permitido.');
        }
        $user = $_SESSION['lti_user'] ?? null;
        if (!is_array($user) || ($user['expires_at'] ?? 0) < time()
            || !is_array($user['course'] ?? null)
            || !is_string($user['course']['contextId'] ?? null)
            || !in_array($user['role'] ?? null, ['instructor', 'learner'], true)) {
            throw new HttpException(401, 'Inicia la herramienta desde el curso en Brightspace.');
        }

        $repository = new CourseSessionRepository();
        if ($method === 'GET') {
            respondJson(200, ['sessions' => $repository->listForCourse($user['course']['contextId'])]);
        }
        if ($user['role'] !== 'instructor') {
            throw new HttpException(403, 'Solo el docente puede modificar las sesiones de este curso.');
        }
        $origin = $_SERVER['HTTP_ORIGIN'] ?? '';
        if (!is_string($origin) || !hash_equals(getOrigin($baseUrl), $origin)) {
            throw new HttpException(403, 'Origen de solicitud no permitido.');
        }
        if ($method === 'DELETE') {
            if (!is_string($sessionId)) {
                throw new HttpException(404, 'No se encontró la sesión publicada en este curso.');
            }
            if (!$repository->deleteForCourse($user['course']['contextId'], $sessionId)) {
                throw new HttpException(404, 'No se encontró la sesión publicada en este curso.');
            }
            respondJson(200, ['deleted' => true]);
        }
        if (!str_contains(strtolower($_SERVER['CONTENT_TYPE'] ?? ''), 'application/json')) {
            throw new HttpException(415, 'El contenido debe enviarse como JSON.');
        }
        $rawBody = file_get_contents('php://input');
        if (!is_string($rawBody) || strlen($rawBody) > 16_384) {
            throw new HttpException(413, 'La solicitud de publicación es demasiado grande.');
        }
        try {
            $payload = json_decode($rawBody, true, 16, JSON_THROW_ON_ERROR);
        } catch (\JsonException) {
            throw new HttpException(400, 'La solicitud de publicación no contiene JSON válido.');
        }
        if (!is_array($payload)
            || !is_string($payload['graphEventId'] ?? null)
            || trim($payload['graphEventId']) === ''
            || strlen($payload['graphEventId']) > 512
            || !is_string($payload['title'] ?? null)
            || trim($payload['title']) === ''
            || strlen($payload['title']) > 1024
            || !is_string($payload['start'] ?? null)
            || !is_string($payload['end'] ?? null)
            || !is_string($payload['timeZone'] ?? null)
            || !is_string($payload['joinUrl'] ?? null)) {
            throw new HttpException(400, 'Faltan datos válidos de la reunión de Teams.');
        }
        $timeZone = trim($payload['timeZone']);
        $dateTimePattern = '/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d{1,7})?(?:Z|[+-]\d{2}:\d{2})?$/D';
        $ianaTimeZones = [
            'Central Standard Time (Mexico)' => 'America/Mexico_City',
            'Pacific Standard Time' => 'America/Los_Angeles',
            'Mountain Standard Time' => 'America/Denver',
            'Central Standard Time' => 'America/Chicago',
            'Eastern Standard Time' => 'America/New_York',
            'UTC' => 'UTC',
        ];
        if ($timeZone === '' || strlen($timeZone) > 128
            || !preg_match($dateTimePattern, $payload['start'])
            || !preg_match($dateTimePattern, $payload['end'])) {
            throw new HttpException(400, 'El enlace o el horario de la reunión no es válido.');
        }
        try {
            $zone = new \DateTimeZone($ianaTimeZones[$timeZone] ?? $timeZone);
            $startDate = new \DateTimeImmutable($payload['start'], $zone);
            $endDate = new \DateTimeImmutable($payload['end'], $zone);
        } catch (\Exception) {
            throw new HttpException(400, 'La zona horaria o el horario de la reunión no es válido.');
        }
        $start = $startDate->getTimestamp();
        $end = $endDate->getTimestamp();
        $joinUrl = filter_var($payload['joinUrl'], FILTER_VALIDATE_URL);
        $joinParts = is_string($joinUrl) ? parse_url($joinUrl) : false;
        if ($start === false || $end === false || $end <= $start
            || !is_array($joinParts)
            || ($joinParts['scheme'] ?? '') !== 'https'
            || strtolower((string) ($joinParts['host'] ?? '')) !== 'teams.microsoft.com'
            || !str_starts_with((string) ($joinParts['path'] ?? ''), '/l/meetup-join/')) {
            throw new HttpException(400, 'El enlace o el horario de la reunión no es válido.');
        }
        $published = $repository->publish($user['course']['contextId'], [
            'graphEventId' => trim($payload['graphEventId']),
            'title' => trim($payload['title']),
            'start' => gmdate(DATE_ATOM, $start),
            'end' => gmdate(DATE_ATOM, $end),
            'timeZone' => $timeZone,
            'joinUrl' => $joinUrl,
        ]);
        $published['graphEventId'] = trim($payload['graphEventId']);
        respondJson(201, ['session' => $published]);
    }

    if ($requestPath === '/api/course/recordings') {
        $method = $_SERVER['REQUEST_METHOD'] ?? '';
        if (!in_array($method, ['GET', 'POST'], true)) {
            header('Allow: GET, POST');
            throw new HttpException(405, 'Método no permitido.');
        }
        $user = $_SESSION['lti_user'] ?? null;
        if (!is_array($user) || ($user['expires_at'] ?? 0) < time()
            || !is_array($user['course'] ?? null)
            || !is_string($user['course']['contextId'] ?? null)
            || !in_array($user['role'] ?? null, ['instructor', 'learner'], true)) {
            throw new HttpException(401, 'Inicia la herramienta desde el curso en Brightspace.');
        }

        $contextId = $user['course']['contextId'];
        new CourseSessionRepository();
        $repository = new CourseRecordingRepository();
        if ($method === 'GET') {
            respondJson(200, [
                'recordings' => $repository->listForCourse($contextId, $user['role'] === 'instructor'),
            ]);
        }
        if ($user['role'] !== 'instructor') {
            throw new HttpException(403, 'Solo el docente puede publicar grabaciones en este curso.');
        }
        $origin = $_SERVER['HTTP_ORIGIN'] ?? '';
        if (!is_string($origin) || !hash_equals(getOrigin($baseUrl), $origin)) {
            throw new HttpException(403, 'Origen de solicitud no permitido.');
        }
        if (!str_contains(strtolower($_SERVER['CONTENT_TYPE'] ?? ''), 'application/json')) {
            throw new HttpException(415, 'El contenido debe enviarse como JSON.');
        }
        $rawBody = file_get_contents('php://input');
        if (!is_string($rawBody) || strlen($rawBody) > 16_384) {
            throw new HttpException(413, 'La solicitud de grabación es demasiado grande.');
        }
        try {
            $payload = json_decode($rawBody, true, 16, JSON_THROW_ON_ERROR);
        } catch (\JsonException) {
            throw new HttpException(400, 'La solicitud de grabación no contiene JSON válido.');
        }
        if (!is_array($payload)
            || !is_string($payload['sessionId'] ?? null)
            || trim($payload['sessionId']) === ''
            || strlen($payload['sessionId']) > 128
            || !is_string($payload['driveItemId'] ?? null)
            || trim($payload['driveItemId']) === ''
            || strlen($payload['driveItemId']) > 1024
            || !is_string($payload['title'] ?? null)
            || trim($payload['title']) === ''
            || strlen($payload['title']) > 1024
            || !is_string($payload['createdAt'] ?? null)
            || !is_string($payload['webUrl'] ?? null)) {
            throw new HttpException(400, 'Faltan datos válidos de la grabación.');
        }
        if (!preg_match(
            '/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d{1,7})?(?:Z|[+-]\d{2}:\d{2})$/D',
            $payload['createdAt']
        )) {
            throw new HttpException(400, 'La fecha de la grabación no es válida.');
        }
        try {
            $createdAt = (new \DateTimeImmutable($payload['createdAt']))->setTimezone(new \DateTimeZone('UTC'));
        } catch (\Exception) {
            throw new HttpException(400, 'La fecha de la grabación no es válida.');
        }
        $webUrl = filter_var($payload['webUrl'], FILTER_VALIDATE_URL);
        $urlParts = is_string($webUrl) ? parse_url($webUrl) : false;
        $host = is_array($urlParts) ? strtolower((string) ($urlParts['host'] ?? '')) : '';
        if (!is_array($urlParts)
            || ($urlParts['scheme'] ?? '') !== 'https'
            || isset($urlParts['user'])
            || isset($urlParts['pass'])
            || !(str_ends_with($host, '.sharepoint.com')
                || $host === '1drv.ms'
                || $host === 'onedrive.live.com')) {
            throw new HttpException(400, 'El enlace de la grabación no es un enlace seguro de Microsoft.');
        }
        $published = $repository->publish($contextId, [
            'sessionId' => trim($payload['sessionId']),
            'driveItemId' => trim($payload['driveItemId']),
            'title' => trim($payload['title']),
            'createdAt' => $createdAt->format(DATE_ATOM),
            'webUrl' => $webUrl,
        ]);
        respondJson(201, ['recording' => $published]);
    }

    if ($requestPath === '/lti/logout') {
        requireMethod('POST');
        $origin = $_SERVER['HTTP_ORIGIN'] ?? '';
        if (!is_string($origin) || !hash_equals(getOrigin($baseUrl), $origin)) {
            throw new HttpException(403, 'Origen de solicitud no permitido.');
        }
        $_SESSION = [];
        if (ini_get('session.use_cookies')) {
            $params = session_get_cookie_params();
            setcookie(session_name(), '', [
                'expires' => time() - 42000,
                'path' => $params['path'],
                'domain' => $params['domain'],
                'secure' => $params['secure'],
                'httponly' => $params['httponly'],
                'samesite' => $params['samesite'] ?? 'None',
            ]);
        }
        session_destroy();
        respondJson(200, ['authenticated' => false]);
    }

    if (($_SERVER['REQUEST_METHOD'] ?? '') === 'GET'
        && $requestPath !== '/api'
        && !str_starts_with($requestPath, '/api/')
        && $requestPath !== '/lti'
        && !str_starts_with($requestPath, '/lti/')
        && is_file(__DIR__.'/index.html')) {
        header('Content-Type: text/html; charset=utf-8');
        readfile(__DIR__.'/index.html');
        exit;
    }

    throw new HttpException(404, 'Ruta no encontrada.');
} catch (HttpException $exception) {
    respondJson($exception->statusCode, ['error' => $exception->getMessage()]);
} catch (OidcException $exception) {
    error_log('Rejected LTI login request: '.$exception::class);
    respondJson(400, ['error' => 'La solicitud de inicio de Brightspace no es válida.']);
} catch (LtiException $exception) {
    error_log('Rejected LTI launch: '.$exception::class);
    respondJson(403, ['error' => 'No se pudo validar el lanzamiento de Brightspace.']);
} catch (Throwable $exception) {
    error_log(sprintf('LTI backend error (%s): %s', $exception::class, $exception->getMessage()));
    respondJson(500, ['error' => 'No se pudo completar la solicitud de Brightspace.']);
}
