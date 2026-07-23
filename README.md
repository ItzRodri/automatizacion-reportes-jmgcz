# Consolidador de Solicitudes — Subalcaldías (app web)

Reemplaza el script de terminal por una página web: subes la Matriz y las
plantillas llenadas, apretás un botón, y descargás el Excel consolidado.

## Archivos de esta carpeta

- `app.py` — la pantalla de la web (lo que ves en el navegador).
- `consolidador.py` — la lógica que ya probamos (leer plantillas, evitar duplicados, etc.).
- `xlsx_dropdown_fix.py` — arregla los desplegables de Excel que se pierden al editar con Python.
- `requirements.txt` — la lista de programas que necesita el servidor para correr la app.

No necesitas modificar nada de estos archivos para publicarla.

## Cómo publicarla gratis (10-15 minutos, una sola vez)

### Paso 1 — Crea una cuenta de GitHub (si no tienes)
Ve a [github.com](https://github.com) → "Sign up". Es gratis.

### Paso 2 — Sube esta carpeta a un repositorio
1. En GitHub, botón verde **"New"** (repositorio nuevo). Ponle un nombre, por ejemplo `consolidador-subalcaldias`. Puede ser **privado** (no necesita ser público).
2. Dentro del repositorio recién creado, usa **"Add file" → "Upload files"** y arrastra los 4 archivos de esta carpeta (`app.py`, `consolidador.py`, `xlsx_dropdown_fix.py`, `requirements.txt`).
3. Dale a **"Commit changes"** (guardar).

### Paso 3 — Despliega en Streamlit Community Cloud (gratis)
1. Ve a [share.streamlit.io](https://share.streamlit.io) y entra con tu cuenta de GitHub.
2. Botón **"New app"**.
3. Elige el repositorio que acabas de crear, la rama `main`, y en "Main file path" escribe `app.py`.
4. Dale a **"Deploy"**. Espera 1-2 minutos mientras instala todo.

Listo — te da un link parecido a `https://consolidador-subalcaldias.streamlit.app`. Ese es tu link permanente: cada vez que lo abras, la app está lista, sin instalar nada.

### Paso 4 (opcional) — Actualizaciones futuras
Si más adelante quieres cambiar algo del código, solo tienes que subir el archivo actualizado a GitHub (Paso 2.2) y la app en Streamlit se actualiza sola en un par de minutos.

## Notas importantes

- **No hay base de datos.** Cada vez que abres la app, no "recuerda" nada de la vez anterior — tú subes la Matriz y las plantillas, y descargas el resultado. La Matriz consolidada que descargas es la que debes guardar y volver a subir la próxima vez.
- **Privacidad:** los archivos que subes solo existen mientras tienes la pestaña abierta procesando; no quedan guardados en ningún servidor.
- **Límite del plan gratis de Streamlit Cloud:** la app se "duerme" si nadie la usa por varios días, y tarda ~20-30 segundos en despertar la primera vez que alguien entra después de eso. Es normal, no es un error.
