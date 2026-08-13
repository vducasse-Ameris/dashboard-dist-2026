# Dashboard Reuniones — Distribución Ameris

Dashboard estático del seguimiento de reuniones con clientes del área de
Distribución. **El procesamiento del Excel ocurre 100% en el navegador**
(via Pyodide — Python compilado a WebAssembly). Nada se sube a un servidor.

## Link público

```
https://<org>.github.io/<repo>/
```

## Cómo se usa (todos los usuarios)

1. Abrir el link en cualquier navegador moderno.
2. Esperar ~15-30 s la primera vez mientras descarga Pyodide + pandas +
   openpyxl (~30 MB). Después queda en cache.
3. Click **↑ Subir Excel** → seleccionar `Follow Up clientes_2026.xlsm`
   (cerrar el archivo en Excel antes — Windows lo bloquea si está abierto).
4. El dashboard se rellena en ~3-5 s.

Para actualizar datos, vuelve a subir el Excel.

## Privacidad

- El repo **no contiene data de clientes** — sólo el código.
- El Excel se procesa en memoria del navegador. Al cerrar la pestaña, los
  datos se borran.
- No hay backend, logs ni telemetría.

## Estructura del repo

```
dashboard-pages/                ← repo de GitHub Pages (self-contained)
├── index.html                  ← dashboard (cargado por el browser)
├── pipeline.py                 ← procesa el Excel (corre en Pyodide)
├── README.md
├── .gitignore
└── build/                      ← herramientas de construcción
    ├── build_pyodide.py        ← regenera index.html
    └── source.html             ← borrador del diseño (fuente de verdad visual)
```

## Cómo regenerar `index.html` (mantenimiento)

Si cambia el diseño o el código JS embebido, regenera el `index.html` así:

```powershell
# Desde la raíz del repo (cualquier persona con Python 3)
python build/build_pyodide.py
```

Solo necesita Python 3 estándar (sin pandas, sin librerías externas).
Lee `build/source.html`, le inyecta toda la lógica Pyodide + Q-rolling +
multi-año, y escribe `index.html` en la raíz.

Después: `git add . && git commit -m "..." && git push`.

## Validación automática del Excel

Cuando subes un Excel, el pipeline valida primero la estructura:

- **Errores críticos** (suben un mensaje rojo y detienen el procesamiento):
  - Falta toda hoja `Clientes foto YYYY`
  - Falta toda hoja `Apuntes YYYY`
  - La hoja del año más reciente no tiene columna `Cliente`
  - Apuntes no tiene columnas `Fecha`, `Empresa/Cliente` o `Subtema`

- **Advertencias** (se muestran en banner amarillo después de cargar):
  - No existe `Clientes foto {año actual}` — usa el más reciente
  - Falta columna `Prioridad`
  - Faltan filas separadoras `Distribuidores` / `Institucionales`

## Diseño para que perdure en el tiempo (2026 → 2027 → 2028…)

El dashboard auto-detecta años desde el Excel y rola sus labels solo.

### Lo que rola automáticamente al subir un Excel nuevo

| Componente | Cómo rola |
|---|---|
| KPIs de Q (Reuniones Q, Top Clientes Q, Activación Q…) | Pipeline usa el **trimestre en curso** del calendario, cortado en el último mes con datos |
| YTD del año actual | Suma hasta el mes en curso, acotado al último mes con reuniones cargadas |
| Comparaciones vs año anterior | Pipeline lee la hoja `Clientes foto YYYY-1` |
| Tabs de año (Vista mensual / Subtemas) | Reconstruidos desde la lista de hojas `Clientes foto YYYY` |
| Heatmap / gráfico de línea / pares YoY | Usan todos los años con hoja `Clientes foto` |
| Nuevos distribuidores | Compara `Clientes foto cur_year` vs `…cur_year-1` (sección Distribuidores). Si no hay nuevos, esconde la sección. |
| Header / Footer / Consolidado / Semáforo / setYoy | Labels Q-relativos rolan via JS |

### Lo que hay que hacer en 2027 (sin tocar código)

1. **Agregar al Excel las hojas nuevas** (manteniendo el formato exacto):
   - `Clientes foto 2027` (con filas "Distribuidores" / "Institucionales"
     como separadores)
   - `Apuntes  2027` (con dos espacios entre "Apuntes" y el año, como las
     demás)
2. Subir el Excel al dashboard. Todo rola solo:
   - Año actual pasa a 2027
   - Top Clientes Q1 2026 → Top Clientes QX 2027 (según el trimestre)
   - Tabs muestran 2027 y los 3 años anteriores
   - Comparativas YoY usan 2026 vs 2027

### Si algún día hay que cambiar el diseño del dashboard

El "borrador" (`../dashboard_reuniones actualizado 13 mayo 2026.html`) es
la fuente del diseño. Para cambios visuales:

1. Modificar `../dashboard/generar_dashboard.py` (o el HTML borrador directo).
2. Regenerar el output.
3. Correr `build_pyodide.py` para que tome la nueva versión.
4. Push.

## Stack

- **Pyodide v0.26.2** — Python 3.11 en WebAssembly.
- **pandas + openpyxl** — leer y procesar el Excel.
- **Chart.js 4.4** — gráficos.
- **HTML/CSS plano** — sin frameworks JS.

## Limitaciones conocidas

- Primera carga del browser ~15-30s. Después es instantáneo (cache).
- Si el Excel cambia su estructura (renombrar columnas críticas como
  `Cliente`, `Prioridad`, `Fecha`, `Empresa/Cliente`, `Subtema`), hay que
  actualizar `pipeline.py`.
- El detector de Distribuidor vs Institucional depende de las filas
  separadoras "Distribuidores" y "Institucionales" en la columna `Cliente`.
  Si se renombran o eliminan, todos quedan clasificados como "Dist".
