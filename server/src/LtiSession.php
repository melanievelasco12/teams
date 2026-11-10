<?php

declare(strict_types=1);

namespace TeamsBrightspace;

use Packback\Lti1p3\Interfaces\ICache;

final class LtiSession implements ICache
{
    private const CACHE_TTL = 600;
    private const SESSION_TTL = 14_400;

    public function getLaunchData(string $key): ?array
    {
        $entry = $_SESSION['lti_launches'][$key] ?? null;
        if (!is_array($entry) || ($entry['expires_at'] ?? 0) < time()) {
            unset($_SESSION['lti_launches'][$key]);
            return null;
        }

        return is_array($entry['data'] ?? null) ? $entry['data'] : null;
    }

    public function cacheLaunchData(string $key, array $jwtBody): void
    {
        $_SESSION['lti_launches'][$key] = [
            'data' => $jwtBody,
            'expires_at' => time() + self::SESSION_TTL,
        ];
    }

    public function cacheNonce(string $nonce, string $state): void
    {
        $_SESSION['lti_nonces'][$nonce] = [
            'state' => $state,
            'expires_at' => time() + self::CACHE_TTL,
        ];
    }

    public function checkNonceIsValid(string $nonce, string $state): bool
    {
        $entry = $_SESSION['lti_nonces'][$nonce] ?? null;
        unset($_SESSION['lti_nonces'][$nonce]);

        return is_array($entry)
            && ($entry['state'] ?? null) === $state
            && ($entry['expires_at'] ?? 0) >= time();
    }

    public function cacheAccessToken(string $key, string $accessToken): void
    {
        $_SESSION['lti_tokens'][$key] = [
            'token' => $accessToken,
            'expires_at' => time() + self::CACHE_TTL,
        ];
    }

    public function getAccessToken(string $key): ?string
    {
        $entry = $_SESSION['lti_tokens'][$key] ?? null;
        if (!is_array($entry) || ($entry['expires_at'] ?? 0) < time()) {
            unset($_SESSION['lti_tokens'][$key]);
            return null;
        }

        return is_string($entry['token'] ?? null) ? $entry['token'] : null;
    }

    public function clearAccessToken(string $key): void
    {
        unset($_SESSION['lti_tokens'][$key]);
    }
}
