<?php

declare(strict_types=1);

namespace TeamsBrightspace;

use Packback\Lti1p3\Interfaces\ICookie;

final class LtiCookie implements ICookie
{
    public function getCookie(string $name): ?string
    {
        $value = $_COOKIE[$name] ?? null;
        return is_string($value) ? $value : null;
    }

    public function setCookie(string $name, string $value, int $exp = 3600, array $options = []): void
    {
        setcookie($name, $value, [
            'expires' => time() + max(0, $exp),
            'path' => '/',
            'secure' => true,
            'httponly' => true,
            'samesite' => 'None',
        ]);
    }
}
