<?php

declare(strict_types=1);

namespace TeamsBrightspace;

final class Config
{
    private array $values;

    public function __construct(?string $envFile = null)
    {
        $this->values = self::readEnvFile($envFile ?? dirname(__DIR__).'/.env');
    }

    public function get(string $key, string $default = ''): string
    {
        $value = getenv($key);
        if ($value !== false) {
            return trim($value);
        }

        return trim((string) ($this->values[$key] ?? $default));
    }

    public function isProduction(): bool
    {
        return $this->get('APP_ENV', 'local') === 'production';
    }

    public function appBaseUrl(): string
    {
        $url = rtrim($this->get('APP_BASE_URL'), '/');
        $parts = parse_url($url);
        if (!is_array($parts) || !isset($parts['scheme'], $parts['host'])) {
            throw new \RuntimeException('APP_BASE_URL debe ser una URL absoluta.');
        }
        if ($this->isProduction() && $parts['scheme'] !== 'https') {
            throw new \RuntimeException('APP_BASE_URL debe usar HTTPS en producción.');
        }

        return $url;
    }

    public function ltiRegistration(): array
    {
        return [
            'issuer' => $this->get('LTI_PLATFORM_ISSUER'),
            'clientId' => $this->get('LTI_PLATFORM_CLIENT_ID'),
            'keySetUrl' => $this->get('LTI_PLATFORM_JWKS_URL'),
            'authTokenUrl' => $this->get('LTI_PLATFORM_TOKEN_URL'),
            'authLoginUrl' => $this->get('LTI_PLATFORM_AUTH_URL'),
            'authServer' => $this->get('LTI_PLATFORM_TOKEN_AUDIENCE'),
            'deploymentId' => $this->get('LTI_DEPLOYMENT_ID'),
            'toolKeyId' => $this->get('LTI_TOOL_KEY_ID'),
            'toolPrivateKeyPath' => $this->get('LTI_TOOL_PRIVATE_KEY_PATH', 'storage/keys/lti-private.pem'),
        ];
    }

    public function ltiReady(): bool
    {
        $registration = $this->ltiRegistration();

        return $this->get('APP_BASE_URL') !== ''
            && $registration['issuer'] !== ''
            && $registration['clientId'] !== ''
            && $registration['keySetUrl'] !== ''
            && $registration['authTokenUrl'] !== ''
            && $registration['authLoginUrl'] !== ''
            && $registration['authServer'] !== ''
            && $registration['deploymentId'] !== ''
            && $registration['toolKeyId'] !== ''
            && is_file($this->privateKeyPath());
    }

    public function privateKeyPath(): string
    {
        $path = $this->get('LTI_TOOL_PRIVATE_KEY_PATH', 'storage/keys/lti-private.pem');
        return str_starts_with($path, DIRECTORY_SEPARATOR)
            || preg_match('/^[A-Za-z]:\\\\/', $path) === 1
            ? $path
            : dirname(__DIR__).DIRECTORY_SEPARATOR.str_replace(['/', '\\'], DIRECTORY_SEPARATOR, $path);
    }

    public function ltiPrivateKey(): string
    {
        $path = $this->privateKeyPath();
        if (!is_file($path) || !is_readable($path)) {
            throw new \RuntimeException('No está configurada la clave privada LTI.');
        }

        $key = file_get_contents($path);
        if ($key === false || openssl_pkey_get_private($key) === false) {
            throw new \RuntimeException('La clave privada LTI no es válida.');
        }

        return $key;
    }

    public function courseNrcClaim(): string
    {
        return $this->get('LTI_COURSE_NRC_CLAIM', 'nrc');
    }

    public function coursePeriodClaim(): string
    {
        return $this->get('LTI_COURSE_PERIOD_CLAIM', 'academic_period');
    }

    private static function readEnvFile(string $path): array
    {
        if (!is_file($path) || !is_readable($path)) {
            return [];
        }

        $lines = file($path, FILE_IGNORE_NEW_LINES | FILE_SKIP_EMPTY_LINES);
        if ($lines === false) {
            throw new \RuntimeException('No se pudo leer la configuración local del servidor.');
        }

        $values = [];
        foreach ($lines as $line) {
            $line = trim($line);
            if ($line === '' || str_starts_with($line, '#') || !str_contains($line, '=')) {
                continue;
            }

            [$key, $value] = explode('=', $line, 2);
            $key = trim($key);
            $value = trim($value);
            if (preg_match('/^[A-Z][A-Z0-9_]*$/', $key) !== 1) {
                continue;
            }
            if (strlen($value) >= 2 && $value[0] === $value[strlen($value) - 1]
                && ($value[0] === '"' || $value[0] === "'")) {
                $value = substr($value, 1, -1);
            }
            $values[$key] = $value;
        }

        return $values;
    }
}
