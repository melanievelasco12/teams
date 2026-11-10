<?php

declare(strict_types=1);

namespace TeamsBrightspace;

use RuntimeException;

final class HttpException extends RuntimeException
{
    public function __construct(
        public readonly int $statusCode,
        string $message
    ) {
        parent::__construct($message);
    }
}
