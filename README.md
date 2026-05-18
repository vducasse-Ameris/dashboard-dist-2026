# Dashboard Reuniones — Distribución Ameris 2026

Dashboard estático para el seguimiento de reuniones con clientes del área
de Distribución. **El procesamiento del Excel ocurre 100% en el navegador**
(via Pyodide — Python compilado a WebAssembly). Nada se sube a un servidor.

## Link público

Una vez activado GitHub Pages (Settings → Pages → Source: `main` / root):

```
https://<org>.github.io/<repo>/
```

## Cómo se usa

1. Abrir el link en cualquier navegador moderno.
2. Esperar ~15-30 s la primera vez mientras descarga Pyodide + pandas
   + openpyxl (~30 MB). Después queda en cache del browser.
3. Click **↑ Subir Excel** → seleccionar `Follow Up clientes_2026.xlsm`
   (cerrar el archivo en Excel antes — Windows lo bloquea si está abierto).
4. El dashboard se rellena en ~3-5 s. Listo.

Cada vez que quieran ver datos actualizados, vuelven a subir el Excel.
El archivo nunca sale del navegador.

## Privacidad

- El repo (público) **no contiene data de clientes**. Sólo el código del
  dashboard y el pipeline Python.
- El Excel se procesa en memoria del navegador. Cuando cerrás la pestaña,
  los datos se borran.
- Nadie tercero ve la data: no hay backend, no hay logs, no hay telemetría.

## Estructura

```
dashboard-pages/
├── index.html         ← dashboard (cargado por el browser)
├── pipeline.py        ← procesa el Excel con pandas (corre en Pyodide)
├── README.md
└── .gitignore
```

## Cómo regenerar `index.html`

El `index.html` se genera transformando el output de `generar_dashboard.py`
(en la carpeta hermana `../dashboard/`):

```powershell
cd ../dashboard
.\.venv\Scripts\python.exe build_pyodide.py
```

Esto toma `dashboard_reuniones actualizado YYYY-MM-DD.html`, le saca toda
la data inline, le inyecta el cargador Pyodide + UI de upload, y escribe el
resultado en `dashboard-pages/index.html`.

## Stack

- **Pyodide v0.26.2** — Python 3.11 en WebAssembly.
- **pandas + openpyxl** — lee y procesa el Excel.
- **Chart.js 4.4** — gráficos.
- **HTML/CSS plano** — sin frameworks JS.

## Limitaciones conocidas

- Las tablas "Activación Q1 2026" y "Áreas / Productos" del Resumen
  quedan vacías — no son derivables del Excel automáticamente. Si las
  queremos, hay que extender `pipeline.py`.
- Primera carga lenta (~15-30 s) por el peso de Pyodide. Después es
  instantáneo (browser cachea).
