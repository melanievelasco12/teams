<?php

declare(strict_types=1);

$directories = [dirname(__DIR__).'/src', dirname(__DIR__).'/public', __DIR__];
foreach ($directories as $directory) {
    $files = new RecursiveIteratorIterator(new RecursiveDirectoryIterator($directory));
    foreach ($files as $file) {
        if ($file->isFile() && $file->getExtension() === 'php') {
            passthru(PHP_BINARY.' -l '.escapeshellarg($file->getPathname()), $exitCode);
            if ($exitCode !== 0) {
                exit($exitCode);
            }
        }
    }
}

exit(0);
