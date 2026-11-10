# 📚 Microsoft Teams TESA

Aplicación web para planificar reuniones de Teams desde cursos de D2L Brightspace. Los docentes crean y publican sesiones en el calendario Microsoft 365; los estudiantes consultan las reuniones y grabaciones publicadas para su curso.

---

## 🌐 URL de producción

Este proyecto no define una URL de producción. Para las pruebas institucionales se ha usado el túnel de desarrollo `https://spousal-engraved-ability.ngrok-free.dev`; no lo uses como alojamiento permanente. Para publicar la aplicación en un dominio institucional, sigue la guía [DEPLOYMENT.md](./DEPLOYMENT.md).

---

## ✨ Funciones principales

- **Planificación docente:** crea reuniones individuales o sesiones semanales para días seleccionados, con horario, zona horaria, descripción, ubicación e invitados.
- **Microsoft Teams:** crea eventos reales en el calendario Microsoft 365 del docente y publica los enlaces de reunión para el curso.
- **Espacio del curso:** los docentes administran las reuniones publicadas y los estudiantes consultan las sesiones del curso LTI actual.
- **Eliminación:** cancela el evento de Microsoft Graph y después quita la sesión del curso y la grabación vinculada. El archivo de OneDrive no se elimina ni se revoca su vínculo.
- **Grabaciones:** detecta grabaciones de la carpeta personal `Recordings` o `Grabaciones` del docente, las asocia con sesiones y comparte un enlace de visualización institucional.
- **Brightspace LTI 1.3:** valida los lanzamientos en PHP y limita lecturas y cambios al contexto del curso verificado.
- **Español y móvil:** interfaz en español con vistas adaptadas para docentes y estudiantes.

## 🚀 Inicio rápido en Windows

1. Instala Node.js 20.19 o posterior (o 22.12 o posterior), npm, PHP 8.3 o posterior y Composer 2. PHP debe incluir `json`, `openssl`, `PDO`, `pdo_sqlite` y `session`.
2. Desde la carpeta del proyecto, instala las dependencias de la interfaz y del backend:

   ```powershell
   npm ci
   Push-Location .\server
   composer install
   Copy-Item .env.example .env
   Pop-Location
   ```

3. Configura `server/.env` con el dominio o URL del entorno, el registro LTI de Brightspace y la clave privada de la herramienta. Genera la clave con:

   ```powershell
   Push-Location .\server
   php bin/generate-lti-key.php
   Pop-Location
   ```

   Guarda el `LTI_TOOL_KEY_ID` que imprime el comando en `server/.env`. No compartas ni subas el archivo `.env` ni la clave privada.
4. Configura `.env.local` en la raíz con los valores públicos de Microsoft Entra:

   ```dotenv
   VITE_ENTRA_CLIENT_ID=ID_PUBLICO_DE_LA_APLICACION
   VITE_ENTRA_TENANT_ID=ID_DEL_TENANT
   VITE_ENTRA_REDIRECT_URI=http://localhost:5173
   VITE_LTI_BACKEND_URL=http://127.0.0.1:8000
   ```

   Las variables `VITE_*` se incluyen en el JavaScript del navegador; nunca pongas secretos ni contraseñas allí.
5. Inicia el servidor PHP desde la raíz del proyecto:

   ```powershell
   php -S 127.0.0.1:8000 -t server/public server/public/router.php
   ```

6. En otra terminal, inicia Vite:

   ```powershell
   npm run dev
   ```

   Abre `http://localhost:5173`. El inicio LTI real desde Brightspace requiere una URL HTTPS públicamente accesible y la configuración correspondiente en Brightspace y Entra. Para publicar en un hosting, sigue [DEPLOYMENT.md](./DEPLOYMENT.md).

## 🔐 Autenticación y permisos

La aplicación combina dos identidades independientes:

- **Brightspace LTI 1.3:** PHP valida la firma y los claims del lanzamiento, y mantiene el rol y el curso en una sesión del servidor. El backend clasifica como estudiante los correos `@estud.tesa.edu.ec` y como docente los correos `@tesa.edu.ec`; también requiere un rol LTI válido y exige el rol LTI de instructor para las cuentas docentes. Los usuarios de otros dominios o con claims insuficientes son rechazados.
- **Microsoft Entra / Graph:** el docente inicia sesión con MSAL y permisos delegados para consultar su perfil y calendario, y leer sus grabaciones de OneDrive. La conexión se selecciona para cada curso Brightspace; la aplicación no adopta automáticamente la primera cuenta guardada en el navegador y solicita elegir una cuenta Microsoft al conectarse. La aplicación no usa un secreto de cliente en el navegador.

Un docente debe conectar la cuenta Microsoft que crea las reuniones para poder administrarlas o cancelarlas. El inicio de sesión de Brightspace no autentica automáticamente al usuario en Microsoft. Estudiantes y docentes solo acceden a datos publicados bajo el contexto Brightspace validado para el curso actual.

## 🚦 Funciones disponibles

- Crear una reunión puntual después de revisar y confirmar sus datos.
- Crear horarios semanales como eventos individuales, hasta 120 sesiones por envío.
- Reintentar publicaciones fallidas sin duplicar reuniones ya publicadas.
- Consultar las reuniones publicadas del curso en las vistas de docente y estudiante.
- Cancelar una reunión desde el calendario Microsoft 365 y quitar su publicación del curso.
- Consultar y publicar grabaciones detectadas en `Recordings` o `Grabaciones` del OneDrive del docente.
- Publicar enlaces de grabación con alcance institucional (`organization`), no anónimos.
- Mantener sesiones y grabaciones del curso en SQLite, aisladas por el identificador de contexto LTI.

La asociación automática de grabaciones se basa en la fecha del archivo y el horario de las sesiones; si hay más de una coincidencia, el archivo no se publica automáticamente. La detección ocurre mientras el docente mantiene abierta la herramienta. Un enlace institucional de OneDrive puede abrirlo cualquier cuenta de TESA que lo obtenga; no queda restringido automáticamente a estudiantes matriculados en el curso.

## ⚙️ Requisitos

- Node.js 20.19+ o 22.12+ y npm.
- PHP 8.3+; se recomienda PHP 8.4 en una versión con soporte de seguridad.
- Extensiones PHP `json`, `openssl`, `pdo`, `pdo_sqlite` y `session`.
- Composer 2 y los paquetes bloqueados en `server/composer.lock`.
- Registro LTI 1.3 de Brightspace y registro de aplicación SPA en Microsoft Entra.
- En producción: dominio HTTPS y almacenamiento persistente y escribible para `server/storage/`.

## 🏗️ Estructura del proyecto

```text
teams/
├── src/                         # Interfaz React y lógica de Microsoft Graph
├── scripts/                     # Limpieza segura de recursos de compilación antiguos
├── server/
│   ├── public/                  # Raíz web, aplicación compilada y rutas PHP
│   ├── src/                     # Configuración, sesión LTI y repositorios SQLite
│   ├── bin/                     # Generación de clave y validación de sintaxis PHP
│   ├── storage/                 # Base SQLite y clave privada; no publicar
│   ├── composer.json
│   └── .env.example
├── DEPLOYMENT.md                # Publicación en hosting privado
├── package.json
└── vite.config.ts
```

En producción, configura `server/public/` como raíz web. No publiques `server/`, `.env`, claves privadas ni la base de datos.

## 📡 Integración con Brightspace y Microsoft

### D2L Brightspace LTI 1.3

Registra en Brightspace las URL del dominio HTTPS definitivo:

- Inicio de sesión: `https://TU-DOMINIO/lti/login`
- Redirección: `https://TU-DOMINIO/lti/launch`
- Claves públicas: `https://TU-DOMINIO/lti/jwks`

El issuer, client ID, URL de autenticación, URL JWKS, audiencia y deployment deben coincidir con `server/.env`. Los datos del curso y el rol se derivan de claims validados; no se aceptan parámetros del navegador como prueba de identidad o curso.

### Microsoft Entra y Graph

En el registro de Entra, agrega una plataforma **Aplicación de página única (SPA)** y registra la URI de retorno exacta de cada entorno. Habilita los permisos delegados:

- `User.Read` para consultar la cuenta conectada.
- `Calendars.ReadWrite` para crear y cancelar reuniones.
- `Files.Read` para leer la carpeta personal de grabaciones.
- `Files.ReadWrite` para crear enlaces institucionales de visualización.

Solicita consentimiento administrativo si lo exige la política del tenant. La licencia, las políticas de Teams y las reglas de grabación de la organización también aplican. No agregues secretos a variables `VITE_*`.

## 🔄 Flujo de operación

1. Brightspace inicia la herramienta y el backend PHP valida el lanzamiento LTI.
2. El backend crea una sesión segura con el rol y el contexto del curso.
3. En cada curso, el docente conecta la cuenta Microsoft correspondiente mediante MSAL; la selección queda asociada al contexto de ese curso en ese navegador.
4. Al publicar una sesión, la interfaz crea el evento con Microsoft Graph y registra sus datos en SQLite bajo el curso LTI actual.
5. La vista del curso consulta únicamente las sesiones y grabaciones asociadas a ese contexto.
6. Al eliminar una sesión, la aplicación cancela primero el evento en Graph y después lo quita del curso. Si falla la limpieza del curso, se puede reintentar sin volver a cancelar el evento.

## 🔒 Variables de entorno y políticas de producción

- `server/.env` contiene la URL pública, los datos confidenciales y la configuración del proveedor LTI; mantenlo fuera del control de versiones y de la raíz web.
- `.env.local` y `.env.production.local` contienen configuración de desarrollo o valores públicos de Entra. No pongas secretos en `VITE_*`.
- La clave privada LTI debe permanecer fuera de `server/public/`, con permisos de lectura limitados al proceso PHP.
- Usa HTTPS; la sesión LTI requiere cookies `Secure`, `HttpOnly` y `SameSite=None` para funcionar dentro del marco de Brightspace.
- Conserva `server/storage/` en un volumen persistente con permisos limitados. Esta implementación usa SQLite local y está destinada a una sola instancia de aplicación, no a varios servidores con discos independientes.
- Usa `ngrok` solo para pruebas autorizadas; no es alojamiento de producción ni mantiene activos los procesos locales.

## 🧩 Componentes principales

- **`src/App.tsx`:** vistas de docente y estudiante, planificación, publicación, eliminación y consulta del curso.
- **`src/graph.ts`:** autenticación MSAL, solicitudes a Microsoft Graph, eventos y lectura de grabaciones.
- **`server/public/index.php`:** rutas LTI y API, validación de sesión, rol, origen y solicitudes.
- **`server/src/Config.php`:** carga y validación de configuración del backend.
- **`server/src/CourseSessionRepository.php`:** sesiones publicadas, zona horaria y persistencia SQLite.
- **`server/src/CourseRecordingRepository.php`:** grabaciones asociadas a sesiones y cursos.
- **`scripts/clean-build-assets.cjs`:** elimina únicamente recursos hashados antiguos tras una compilación correcta.

## 🗄️ Base de datos

Las sesiones y grabaciones se guardan en `server/storage/course-sessions.sqlite`. Las tablas relacionan cada dato con el contexto LTI y las grabaciones se vinculan a una sesión; al eliminarla, SQLite elimina la referencia de la grabación del curso. El archivo de video y su enlace de OneDrive no se eliminan.

Haz copias de seguridad de `server/storage/` y prueba su restauración. No coloques esta carpeta bajo la raíz pública ni uses discos locales separados detrás de un balanceador.

## ⚙️ Configuración

Consulta [`server/.env.example`](./server/.env.example) para las variables del backend y [DEPLOYMENT.md](./DEPLOYMENT.md) para el procedimiento de publicación, HTTPS, permisos, Apache/Nginx y almacenamiento persistente.

La aplicación sirve su interfaz desde `server/public/index.html`. En desarrollo, Vite puede reenviar `/api` y `/lti` al backend si se configura `VITE_LTI_BACKEND_URL`.

## 🔧 Actualizar el sistema

1. Conserva una copia de seguridad de `server/storage/` y de la configuración privada.
2. Actualiza el código usando el repositorio y las instrucciones institucionales.
3. Instala dependencias respetando los archivos de bloqueo: `npm ci` y `composer install --no-dev`.
4. Compila la interfaz con `npm run build`.
5. Verifica los requisitos PHP con `composer check-platform-reqs --no-dev` y revisa la aplicación antes de reabrir el acceso.

## ▶️ Ejecución y pruebas

Comandos de desarrollo desde la raíz del proyecto:

```powershell
npm run dev
npm run build
```

Para comprobar la sintaxis PHP, desde `server/`:

```powershell
php bin/lint.php
```

Para revisar las dependencias:

```powershell
npm audit
Push-Location .\server
composer audit
Pop-Location
```

El proyecto no declara una suite automatizada de pruebas de interfaz. Antes de publicar, verifica el inicio LTI desde Brightspace, el rol docente y estudiante, el aislamiento entre cursos, la creación y cancelación de reuniones, y la publicación y consulta de grabaciones.

## 📋 Análisis técnico y pendientes

### Límites operativos

- La detección de grabaciones requiere que el docente abra la herramienta con su cuenta Microsoft conectada.
- Los enlaces de grabación tienen alcance institucional y no se limitan a la matrícula del curso.
- SQLite requiere almacenamiento persistente y una sola instancia de aplicación con disco compartido local.
- Brightspace debe entregar los claims LTI de correo, rol y contexto esperados por la institución.

### Antes de producción

- Sustituye la URL de ngrok por un dominio institucional HTTPS.
- Confirma con TI el formato real de los claims LTI y las políticas de cookies dentro del iframe.
- Valida los permisos Graph y las políticas de grabación con el administrador de Microsoft 365.
- Prueba copias de seguridad, restauración y permisos del almacenamiento persistente.

## ✅ Reglas de seguridad

- No subas `.env`, claves privadas, tokens, contraseñas ni datos reales al repositorio.
- No confíes en curso, rol o identidad enviados por la interfaz: el backend debe derivarlos de la sesión LTI validada.
- Mantén `server/public/` como única raíz web pública.
- No configures secretos en variables `VITE_*`; forman parte de los archivos enviados al navegador.
- Limita los permisos Graph y el acceso al almacenamiento a lo necesario.
- Usa HTTPS en producción y cierra las sesiones Microsoft en dispositivos compartidos.

## 📞 Contacto del proyecto

- **Desarrolladora de la interfaz:** Melanie Velasco
- **Correo:** [velascomelanie063@gmail.com](mailto:velascomelanie063@gmail.com)

Para habilitar el registro LTI, los permisos de Microsoft Entra/Graph, el hosting o los claims institucionales, contacta al administrador de Brightspace, al equipo de TI o al administrador de Microsoft 365 de TESA.
