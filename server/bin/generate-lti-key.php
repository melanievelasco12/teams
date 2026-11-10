<?php

declare(strict_types=1);

use TeamsBrightspace\Config;

require dirname(__DIR__).'/vendor/autoload.php';

$config = new Config();
$keyPath = $config->privateKeyPath();
$keyDirectory = dirname($keyPath);
if (is_file($keyPath)) {
    fwrite(STDERR, "La clave ya existe; no se sobrescribió.\n");
    exit(1);
}
if (!is_dir($keyDirectory) && !mkdir($keyDirectory, 0700, true) && !is_dir($keyDirectory)) {
    fwrite(STDERR, "No se pudo crear la carpeta de claves.\n");
    exit(1);
}

$key = openssl_pkey_new([
    'private_key_type' => OPENSSL_KEYTYPE_RSA,
    'private_key_bits' => 2048,
]);
if ($key === false || !openssl_pkey_export($key, $privateKey)) {
    $errors = [];
    while (($error = openssl_error_string()) !== false) {
        $errors[] = $error;
    }
    fwrite(STDERR, "No se pudo generar la clave privada LTI.");
    if ($errors !== []) {
        fwrite(STDERR, ' '.implode('; ', $errors));
    }
    fwrite(STDERR, "\nEn Windows, configura OPENSSL_CONF antes de ejecutar este comando si OpenSSL no encuentra su archivo openssl.cnf.\n");
    exit(1);
}
if (file_put_contents($keyPath, $privateKey, LOCK_EX) !== strlen($privateKey)) {
    if (is_file($keyPath)) {
        unlink($keyPath);
    }
    fwrite(STDERR, "No se pudo guardar la clave privada LTI.\n");
    exit(1);
}
if (!chmod($keyPath, 0600)) {
    unlink($keyPath);
    fwrite(STDERR, "No se pudieron restringir los permisos de la clave privada LTI.\n");
    exit(1);
}

fwrite(STDOUT, "Clave privada guardada en: {$keyPath}\n");
fwrite(STDOUT, "Configura este identificador como LTI_TOOL_KEY_ID: ".bin2hex(random_bytes(12))."\n");
