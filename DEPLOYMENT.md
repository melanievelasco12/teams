# Despliegue en alojamiento privado

Esta guía describe cómo publicar Microsoft Teams en un alojamiento administrado por la institución y conservar la interfaz, el inicio LTI de Brightspace y las sesiones publicadas.

## Antes de contratar o preparar el alojamiento

Confirma con el proveedor o el equipo de TI que el alojamiento tenga:

- PHP 8.4 (PHP 8.3 o superior también satisface `server/composer.json`).
- Extensiones PHP `json`, `openssl`, `PDO`, `pdo_sqlite` y `session`.
- Composer 2 y acceso para instalar las dependencias de `server/composer.lock`.
- HTTPS válido para el dominio final.
- Un proceso PHP persistente (Apache con PHP o Nginx con PHP-FPM).
- Un directorio persistente y escribible por PHP para `server/storage/`.
- Posibilidad de configurar como raíz web `server/public/`, o una configuración equivalente que impida publicar el resto del proyecto.

La aplicación debe poder llegar por HTTPS desde Brightspace y desde los navegadores de los usuarios para completar LTI 1.3. El servidor y la base de datos pueden ser privados y administrados por la institución, pero un sitio accesible solo desde la red local o protegido íntegramente por otra pantalla de inicio de sesión no permitirá que Brightspace complete el lanzamiento dentro de un marco. El acceso a cursos y roles lo controla la sesión LTI validada por el servidor PHP. No expongas `server/`, `.env`, claves privadas, dependencias ni almacenamiento como contenido web.

El servidor usa SQLite en `server/storage/course-sessions.sqlite`. Este despliegue está pensado para una sola instancia de la aplicación con almacenamiento local persistente. No uses réplicas o varios servidores con discos independientes: cada instancia vería una base distinta. Para distribuir la carga entre varios servidores, primero habría que migrar el almacenamiento a una base de datos compartida.

## 1. Preparar el dominio y compilar la interfaz

Define un dominio institucional, por ejemplo `https://teams.tu-institucion.edu`, con certificado HTTPS. Usa la dirección final desde el inicio; cambiarla después requiere actualizar también Brightspace, Entra y el servidor.

En una copia local del proyecto, configura las variables públicas de Entra para la aplicación web de una sola página en `.env.production.local` (en la carpeta `teams`, al mismo nivel que `package.json`):

```dotenv
VITE_ENTRA_CLIENT_ID=ID_PUBLICO_DE_LA_APLICACION
VITE_ENTRA_TENANT_ID=ID_DEL_TENANT
VITE_ENTRA_REDIRECT_URI=https://teams.tu-institucion.edu
```

El identificador de cliente y el identificador del inquilino son datos públicos de la aplicación web, no secretos. No pongas secretos de cliente, contraseñas ni credenciales de acceso en variables `VITE_*`: Vite las incorpora al JavaScript que reciben los usuarios.

Genera la versión publicada con Node.js 20.19 o posterior, o 22.12 o posterior, y npm:

```powershell
npm ci
npm run build
```

El resultado se genera en `server/public/`. Sube ese contenido junto con el servidor PHP siguiendo la estructura del siguiente paso. No publiques el servidor de desarrollo de Vite ni uses ngrok como alojamiento de producción.

## 2. Publicar los archivos y las dependencias PHP

Sube el código del proyecto al servidor usando Git o un canal institucional seguro. En producción, la raíz pública del dominio (la carpeta de publicación o `DocumentRoot`) debe apuntar exactamente a:

```text
/ruta/privada/teams/server/public
```

Mantén el resto del proyecto fuera de la raíz pública. En particular, no configures como carpeta de publicación `teams/` ni `teams/server/`.

Desde `teams/server`, instala las dependencias bloqueadas por Composer:

```sh
composer install --no-dev --prefer-dist --optimize-autoloader --no-interaction
```

Ejecuta el comando con la misma versión de PHP que utilizará el servidor web. Confirma los requisitos con:

```sh
php -v
php -m
composer check-platform-reqs --no-dev
```

La lista debe incluir `json`, `openssl`, `PDO`, `pdo_sqlite` y `session`; la línea `PDO drivers` debe incluir `sqlite`. La extensión independiente `sqlite3` no es necesaria para este proyecto. Si el alojamiento no ofrece estas extensiones o no permite un directorio escribible y persistente, solicita al proveedor que lo habilite antes de continuar.

### Apache

Apunta el `DocumentRoot` a `server/public`, habilita `mod_rewrite` y permite las reglas `.htaccess` (por ejemplo, `AllowOverride All`). El archivo `server/public/.htaccess` ya dirige las solicitudes de `/api` y `/lti` hacia el servidor PHP.

### Nginx con PHP-FPM

Nginx no procesa `.htaccess`. Configura el sitio para servir archivos existentes desde `server/public` y enviar las demás rutas a `server/public/index.php`. El equipo de TI debe adaptar este bloque básico a su versión de PHP-FPM:

```nginx
root /ruta/privada/teams/server/public;
index index.php index.html;

location / {
    try_files $uri $uri/ /index.php?$query_string;
}

location ~ \.php$ {
    try_files $uri =404;
    include fastcgi_params;
    fastcgi_param SCRIPT_FILENAME $document_root$fastcgi_script_name;
    fastcgi_pass unix:/run/php/php-fpm.sock;
}

location ~ /\.(?!well-known) {
    deny all;
}
```

Sustituye la ruta de conexión de ejemplo por la que proporcione PHP-FPM. La configuración final debe impedir que los archivos PHP o de configuración se descarguen como texto y solo debe exponer `server/public`.

## 3. Configurar el servidor y proteger sus datos

En el servidor, copia `server/.env.example` a `server/.env` y edita el nuevo archivo. No compartas ni subas el `.env` real al repositorio:

```sh
cd /ruta/privada/teams/server
cp .env.example .env
```

Establece como mínimo:

```dotenv
APP_ENV=production
APP_BASE_URL=https://teams.tu-institucion.edu
LTI_DEPLOYMENT_ID=ID_DE_IMPLEMENTACION_DE_BRIGHTSPACE
LTI_TOOL_KEY_ID=IDENTIFICADOR_IMPRESO_POR_EL_GENERADOR
LTI_TOOL_PRIVATE_KEY_PATH=storage/keys/lti-private.pem
```

Revisa y confirma también en ese archivo `LTI_PLATFORM_ISSUER`, `LTI_PLATFORM_CLIENT_ID`, `LTI_PLATFORM_AUTH_URL`, `LTI_PLATFORM_TOKEN_URL`, `LTI_PLATFORM_TOKEN_AUDIENCE` y `LTI_PLATFORM_JWKS_URL` con los datos del registro LTI de la institución. Define `LTI_DEPLOYMENT_ID` cuando Brightspace genere el identificador de implementación.

Genera la clave de la herramienta una sola vez, después de crear `.env` e instalar Composer:

```sh
php bin/generate-lti-key.php
```

Copia exactamente el identificador que imprime el comando a `LTI_TOOL_KEY_ID` y conserva `server/storage/keys/lti-private.pem` de forma privada. No regeneres la clave en cada publicación: el identificador y la clave pública registrada en Brightspace deben seguir correspondiendo a la clave privada activa. Restringe los permisos del `.env`, de la clave y del directorio de almacenamiento al usuario del servicio y a los administradores autorizados.

El usuario que ejecuta PHP necesita permiso de escritura en `server/storage/` para crear y actualizar SQLite. En Linux, configura propietario/grupo y permisos según el usuario real de Apache o PHP-FPM; no uses permisos globales de escritura como `777`. Mantén las sesiones PHP en un almacén persistente del servidor y no borres sus archivos mientras haya usuarios con una sesión activa.

Haz copias de seguridad periódicas de:

- `server/storage/course-sessions.sqlite` y, si existen, sus archivos SQLite `-wal` y `-shm`.
- `server/storage/keys/lti-private.pem`.
- `server/.env`, en un gestor de secretos o una copia de seguridad cifrada y con acceso restringido.

Usa una herramienta de copia de seguridad compatible con SQLite o detén brevemente la aplicación antes de copiar la base. No copies únicamente el archivo principal mientras SQLite esté escribiendo en el registro WAL. Prueba la restauración en un entorno separado. La clave privada y el `.env` nunca deben incluirse en copias públicas.

## 4. Registrar las URL finales en Entra y Brightspace

En el registro de aplicación de Microsoft Entra, agrega `https://teams.tu-institucion.edu` como URI de retorno de la aplicación web de una sola página. Debe coincidir exactamente con `VITE_ENTRA_REDIRECT_URI`. Conserva los permisos delegados de Microsoft Graph que utiliza el proyecto y solicita la aprobación institucional requerida. No crees un secreto de cliente para esta aplicación.

En Brightspace, registra el dominio final y configura las URL LTI 1.3:

| Configuración | URL |
|---|---|
| Inicio de sesión / OIDC | `https://teams.tu-institucion.edu/lti/login` |
| URI de redirección / lanzamiento | `https://teams.tu-institucion.edu/lti/launch` |
| URL de claves públicas | `https://teams.tu-institucion.edu/lti/jwks` |

Usa el emisor, el identificador de cliente y el identificador de implementación correspondientes a ese registro de Brightspace. La institución debe confirmar que el token de identidad del lanzamiento incluye los atributos requeridos por la aplicación (correo, rol y contexto del curso). Prueba el lanzamiento desde un curso real con una cuenta docente y otra de estudiante. No tomes como señal de fallo un `401` al abrir `/api/course/sessions` directamente en una ventana: ese servicio requiere la cookie de una sesión LTI iniciada desde el curso.

## 5. Verificación y mantenimiento

Antes de habilitar a los usuarios:

1. Confirma que el dominio presenta un certificado HTTPS válido y que redirige HTTP a HTTPS.
2. Abre `https://teams.tu-institucion.edu/` y verifica que carga la interfaz compilada.
3. Abre `https://teams.tu-institucion.edu/lti/jwks`; con la clave configurada debe responder JSON con la clave pública. Este endpoint no debe revelar la clave privada.
4. Inicia la herramienta desde Brightspace y comprueba la sesión y el curso correctos.
5. Desde una cuenta docente, publica una reunión de prueba y confirma que aparece para usuarios del mismo curso después de abrir LTI. Comprueba que otra clase no la ve.
6. Desde la lista de reuniones, conecta la cuenta Microsoft que creó el evento y prueba **Eliminar** con una sesión de prueba. Confirma que se cancela en el calendario y desaparece del curso. Microsoft puede notificar la cancelación a las personas invitadas; el archivo y el vínculo de OneDrive no se revocan.
7. Reinicia el servicio y confirma que las demás reuniones siguen presentes. Esto valida que SQLite reside en almacenamiento persistente.
8. Revisa los registros del servidor y de PHP-FPM/Apache para detectar errores; no registres ni compartas cookies, credenciales de acceso, claves ni el contenido del `.env`.

Para actualizar la aplicación, conserva primero las copias de seguridad. Publica la nueva versión, vuelve a generar la interfaz con la configuración de producción, ejecuta `composer install --no-dev --prefer-dist --optimize-autoloader --no-interaction`, conserva sin cambios el `.env`, la clave LTI y la base SQLite, y reinicia PHP-FPM/Apache si el proveedor lo requiere. Verifica de nuevo un lanzamiento LTI y la persistencia de las sesiones.

## Consideraciones sobre privacidad

- Usa alojamiento, copias de seguridad, registros y acceso administrativo aprobados por la institución.
- Da al usuario de PHP solo los permisos necesarios; mantén `.env`, `vendor/`, `src/`, `storage/` y las claves fuera del área pública.
- El lanzamiento LTI autentica el curso y el rol; no reemplaces esta validación por una dirección secreta o una pantalla de contraseña externa que bloquee las respuestas de retorno de Brightspace.
- Los enlaces de grabación que la aplicación genera con alcance `organization` pueden abrirse con una cuenta de TESA que los obtenga; no quedan limitados a la matrícula del curso.
- Mantén PHP, Composer y las dependencias actualizados dentro de un ciclo probado, y define con TI responsables de monitorización, retención y restauración.
