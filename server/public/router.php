<?php

declare(strict_types=1);

$path = parse_url($_SERVER['REQUEST_URI'] ?? '/', PHP_URL_PATH);
$requestedFile = realpath(__DIR__.DIRECTORY_SEPARATOR.ltrim((string) $path, '/\\'));
$publicDirectory = realpath(__DIR__);
if ($requestedFile !== false && $publicDirectory !== false
    && str_starts_with($requestedFile, $publicDirectory.DIRECTORY_SEPARATOR)
    && is_file($requestedFile)) {
    return false;
}

require __DIR__.'/index.php';
