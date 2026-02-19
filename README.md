# Convertidor Ejecutivo | Magaya HTML a Excel

Herramienta web que convierte reportes HTML exportados de Magaya a archivos Excel con formato profesional (estilos, colores, anchos de columna). Todo el procesamiento se hace en el navegador; no se envían datos a ningún servidor.

## Uso

1. Abre la aplicación (localmente o en la URL publicada).
2. Arrastra un archivo `.html` o `.htm` de Magaya a la zona de carga, o haz clic para seleccionarlo.
3. Revisa el resultado y descarga el Excel generado.

## Publicar en GitHub (GitHub Pages)

### Opción A: Repositorio público

1. Crea un repositorio en GitHub (ej. `Generador-de-Relaciones`).
2. Sube el contenido del proyecto (carpetas `css/`, `js/`, `scripts/`, archivos `index.html`, `404.html`, `README.md`, etc.).
3. En el repositorio: **Settings** → **Pages**.
4. En **Source** elige **Deploy from a branch**.
5. En **Branch** elige `main` (o `master`) y carpeta **/ (root)**.
6. Guarda. En unos minutos la web estará en:
   - `https://<tu-usuario>.github.io/<nombre-repo>/`

### Opción B: Usar solo la rama `gh-pages`

1. Crea una rama llamada `gh-pages`.
2. Sube ahí `index.html`, `404.html`, y las carpetas `css/`, `js/`.
3. En **Settings** → **Pages** selecciona la rama `gh-pages` y raíz.
4. La URL será la misma que arriba.

## Estructura del proyecto

```
├── index.html      # Página principal
├── 404.html        # Página de error (redirige a la app)
├── css/
│   └── styles.css
├── js/
│   └── app.js
├── scripts/        # Solo para desarrollo local (abrir HTML)
├── .gitignore
└── README.md
```

## Requisitos

- Navegador moderno con JavaScript activado.
- Archivos HTML generados por Magaya (compatibles con la estructura esperada por la herramienta).

## Privacidad y seguridad

- El procesamiento es **100% local** en tu navegador.
- No se envían archivos ni datos a la nube.
- Los reportes se leen con codificación ISO-8859-1 para compatibilidad con Magaya.

## Licencia

Uso interno / proyecto propio. Ajusta según tu política.
