# Asistencia · Laboratorio Impro

## ⚠️ Antes de nada: lee esto

Este proyecto necesita una base de datos real (Firebase, gratis) para que
los datos se compartan entre todos los profesores en tiempo real, desde
cualquier móvil u ordenador. Sin ese paso, la app funciona igual, pero cada
navegador guarda sus datos por su cuenta, sin compartirlos con nadie más
(verás un aviso arriba de la app recordándotelo mientras no lo hagas).

Las instrucciones para configurarlo están dentro de `src/storage.js`, en
forma de comentario al principio del archivo.

## Cómo subir esto a Netlify (paso a paso, sin usar la terminal)

No hace falta instalar nada en tu ordenador. Se hace todo desde el
navegador, en dos webs: GitHub (donde "vive" el proyecto) y Netlify (que lo
publica).

### Paso 1 — Sube esta carpeta a GitHub

1. Ve a [github.com](https://github.com) y créate una cuenta gratis si no
   tienes (con tu email, como cualquier registro).
2. Una vez dentro, botón verde "New" (o el símbolo "+" arriba a la derecha
   → "New repository").
3. Ponle un nombre, por ejemplo `asistencia-impro`. Déjalo en "Public" o
   "Private" (da igual). Pulsa "Create repository".
4. En la página del repositorio recién creado, busca el enlace que dice
   algo como "uploading an existing file" (o el botón "Add file" →
   "Upload files").
5. Arrastra ahí **todos los archivos y carpetas** que te he dado en el
   zip (`index.html`, `netlify.toml`, `package.json`, `vite.config.js`,
   `README.md` y la carpeta `src` completa con sus 3 archivos dentro).
6. Baja del todo y pulsa "Commit changes" (el botón verde) para guardarlo.

### Paso 2 — Conecta ese repositorio con Netlify

1. Ve a [app.netlify.com](https://app.netlify.com) y créate una cuenta
   gratis (puedes entrar directamente con tu cuenta de GitHub, es lo más
   rápido).
2. Botón "Add new site" → "Import an existing project".
3. Elige "Deploy with GitHub" y selecciona el repositorio que acabas de
   crear (`asistencia-impro`).
4. Netlify ya sabe cómo construirlo solo, porque lo dice el archivo
   `netlify.toml` que subiste: no tienes que rellenar nada, simplemente
   pulsa "Deploy site".
5. Espera un par de minutos. Cuando termine, te da una dirección web (algo
   como `nombre-al-azar.netlify.app`) — esa es la que compartes con tus
   compañeros.
6. **A partir de ahora**: cada vez que quieras cambiar algo del código (por
   ejemplo, rellenar los datos de Firebase), solo tienes que volver a
   GitHub, editar o resubir ese archivo, y Netlify reconstruye la web sola
   en un par de minutos — no hay que repetir estos pasos.

## Contraseñas

- La contraseña de administrador está dentro de `src/App.jsx`, cerca del
  principio del archivo, en una línea que pone
  `const ADMIN_MASTER_PASSWORD = "..."`. Cámbiala por la que quieras antes
  de compartir el enlace con nadie.
- La contraseña de cada profesor se pone **desde dentro de la propia app**
  (no hay que tocar código): pestaña Profesores → clicar el nombre del
  profesor → campo "Contraseña personal".

## Copia de seguridad

Desde la app, como administrador, puedes descargar en cualquier momento:

- Excel de Asistencia (pestaña Alumnos)
- Excel de Pagos (pestaña Alumnos)
- Excel de Dinámicas (pestaña Dinámicas)

Guárdalos donde quieras (tu Drive, tu ordenador...) de vez en cuando, como
respaldo aparte de lo que ya guarda Firebase.
