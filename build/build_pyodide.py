"""
build_pyodide.py
────────────────
Genera el index.html del dashboard (Pyodide-powered) a partir de
`build/source.html` (el "borrador" del diseño).

- Quita TODA la data sensible inline (clientes, meetings, etc.)
- Inyecta Pyodide loader + UI de upload
- Inyecta applyData(d) que toma el dict del pipeline y re-renderiza todo
- Inyecta automatización Q-rolling y multi-año
- Output: ../index.html (público en GitHub Pages)

Este script vive dentro del repo (no depende de archivos externos),
así que cualquier persona puede clonarlo y regenerar el dashboard:

    python build/build_pyodide.py

Requisitos: solo Python 3 (sin librerías externas).
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

HERE        = Path(__file__).resolve().parent       # .../dashboard-pages/build
REPO_DIR    = HERE.parent                            # .../dashboard-pages
SOURCE_HTML = HERE / "source.html"                   # borrador del diseño
OUT_HTML    = REPO_DIR / "index.html"


# ════════════════════════════════════════════════════════════════════════════
# Bloques que inyectamos en el HTML
# ════════════════════════════════════════════════════════════════════════════

PYODIDE_HEAD = """
<!-- ──────────────── Pyodide (motor Python en el browser) ──────────────── -->
<script src="https://cdn.jsdelivr.net/pyodide/v0.26.2/full/pyodide.js"></script>
<style>
  /* Barra de upload */
  #upload-bar {
    background: #FFFFFF; border-bottom: 1px solid var(--border);
    padding: 10px 2rem; display: flex; align-items: center; gap: 14px;
    font-size: 12px; flex-wrap: wrap;
  }
  #upload-bar.empty { background: linear-gradient(180deg,#FAFBFD 0%,#FFFFFF 100%); padding: 22px 2rem; }
  #upload-bar .ub-status { display:flex; align-items:center; gap:8px; color:var(--text2); }
  #upload-bar .ub-spinner {
    width:14px; height:14px; border:2px solid var(--border2); border-top-color: var(--a1);
    border-radius:50%; animation: spin 0.8s linear infinite;
  }
  @keyframes spin { to { transform: rotate(360deg); } }
  #upload-bar .ub-check  { color: var(--green); font-weight:600; }
  #upload-bar .ub-error  { color: var(--red);   font-weight:600; }
  #upload-bar .ub-info   { color: var(--text3); font-size:11px; }
  #upload-bar .ub-meta   { color: var(--text3); font-size:11px; font-family:'IBM Plex Mono',monospace; }
  #upload-bar button {
    background: var(--a1); color:#fff; border:0; border-radius:6px;
    padding: 7px 14px; font-size: 12px; font-weight:500; cursor: pointer;
    display:flex; align-items:center; gap:6px;
  }
  #upload-bar button:hover { background: var(--a2); }
  #upload-bar button:disabled { background: var(--text3); cursor: not-allowed; }
  #upload-bar .ub-secondary { background: var(--card2); color: var(--text2); border:1px solid var(--border); }
  #upload-bar .ub-secondary:hover { background: var(--border); color: var(--text); }
  /* Estado "esperando upload" — esconde paneles para que se vea limpio */
  body.no-data .panel { opacity: 0.35; pointer-events: none; }
  body.no-data #upload-bar { background: #FFF3CD; border-bottom-color:#E5C97D; }
  body.no-data #upload-bar .ub-info { color: #6B5100; font-weight:500; }

  /* ── Cruces por distribuidor (saldo/gasto vs reuniones) ── */
  .xhead { display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;
    font-size:10px; color:var(--text3); text-transform:uppercase; letter-spacing:.05em; font-weight:600; }
  .xrow { display:flex; align-items:center; gap:12px; padding:10px 8px;
    border-bottom:1px solid var(--border); transition:background .15s; }
  .xrow:last-child { border-bottom:0; }
  .xrow:hover { background:var(--card2); }
  .xrank { display:inline-flex; align-items:center; justify-content:center; min-width:24px; height:24px;
    padding:0 6px; border-radius:7px; font-family:'IBM Plex Mono',monospace; font-size:11px; font-weight:700;
    background:#EBF0F8; color:#1B4B9B; flex-shrink:0; }
  .xrank.top { background:linear-gradient(135deg,#2E6BAF,#1B4B9B); color:#fff;
    box-shadow:0 1px 3px rgba(27,75,155,.3); }
  .xname { font-size:12px; font-weight:600; color:var(--text);
    white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
  .xval { font-family:'IBM Plex Mono',monospace; font-size:13px; font-weight:700; color:var(--text); white-space:nowrap; }
  .xbar-track { height:5px; background:var(--border); border-radius:3px; margin-top:6px; overflow:hidden; }
  .xbar-fill { height:100%; border-radius:3px; }
  .xpill { display:inline-flex; align-items:baseline; gap:3px; padding:3px 10px; border-radius:999px;
    font-family:'IBM Plex Mono',monospace; font-size:12px; font-weight:700; background:#EBF0F8; color:#1B4B9B; flex-shrink:0; }
  .xpill small { font-size:9px; font-weight:600; opacity:.7; }
  .xpill.zero { background:#FAE9E7; color:#B33A2E; }
  .xsub { font-size:10px; font-family:'IBM Plex Mono',monospace; margin-top:3px; white-space:nowrap; }
  .xscroll { max-height:520px; overflow-y:auto; }
  .xscroll::-webkit-scrollbar { width:8px; }
  .xscroll::-webkit-scrollbar-thumb { background:var(--border2); border-radius:4px; }
</style>
"""


UPLOAD_BAR_HTML = """
<!-- ── Upload bar (Pyodide-powered) ────────────────────────────────────── -->
<div id="upload-bar" class="empty">
  <input type="file" id="xlsm-input" accept=".xlsm,.xlsx" style="display:none;">
  <button id="upload-btn" disabled onclick="document.getElementById('xlsm-input').click();">
    <span id="upload-btn-icon">⏳</span> <span id="upload-btn-text">Cargando motor…</span>
  </button>
  <input type="file" id="metas-input" accept=".xlsx,.xlsm" style="display:none;">
  <button id="metas-btn" class="ub-secondary" disabled onclick="document.getElementById('metas-input').click();"
          title="Opcional: archivo de resultados de metas (avance AUM)">
    <span id="metas-btn-text">↑ Resultados metas</span>
  </button>
  <div class="ub-status" id="upload-status">
    <span class="ub-spinner" id="upload-spinner"></span>
    <span id="upload-status-text">Iniciando Python en el navegador (primera vez ~15 s, después instantáneo)…</span>
  </div>
  <div class="ub-info" style="margin-left:auto;">
    Los datos viven sólo en tu navegador — nada se sube a un servidor.
  </div>
</div>
"""


# applyData() y handlers Pyodide. Va al final del <script>, después de las
# render functions existentes pero antes del </script> de cierre.
APPLY_DATA_JS = r"""

// ══════════════════════════════════════════════════════════════════════════
// PYODIDE — motor Python en browser para procesar el Excel
// ══════════════════════════════════════════════════════════════════════════

let pyodideReady = null;

async function bootPyodide() {
  const statusText = document.getElementById('upload-status-text');
  const spinner    = document.getElementById('upload-spinner');
  const btn        = document.getElementById('upload-btn');
  const btnIcon    = document.getElementById('upload-btn-icon');
  const btnText    = document.getElementById('upload-btn-text');

  try {
    statusText.textContent = 'Cargando Pyodide…';
    const py = await loadPyodide({ indexURL: 'https://cdn.jsdelivr.net/pyodide/v0.26.2/full/' });
    statusText.textContent = 'Cargando pandas…';
    await py.loadPackage(['pandas', 'micropip']);
    statusText.textContent = 'Instalando openpyxl (pip)…';
    await py.runPythonAsync(`
import micropip
await micropip.install('openpyxl')
`);
    statusText.textContent = 'Cargando pipeline…';
    // cache-bust: el browser cacheaba pipeline.py y nos quedaba con una versión vieja.
    const pipelineSrc = await (await fetch('./pipeline.py?v=' + Date.now(), { cache: 'no-store' })).text();
    py.FS.writeFile('pipeline.py', pipelineSrc);
    py.runPython('import pipeline');

    // Listo
    document.body.classList.add('no-data');
    spinner.style.display = 'none';
    statusText.textContent = '';
    statusText.innerHTML = '<span class="ub-info">Sube el Excel Follow Up clientes para ver los datos.</span>';
    btn.disabled = false;
    btnIcon.textContent = '↑';
    btnText.textContent = 'Subir Excel';
    var mbtn = document.getElementById('metas-btn');
    if (mbtn) mbtn.disabled = false;
    return py;
  } catch (err) {
    spinner.style.display = 'none';
    statusText.innerHTML = '<span class="ub-error">Error iniciando Pyodide: ' + (err.message || err) + '</span>';
    throw err;
  }
}

// Extrae mensaje limpio del error Python en Pyodide (descarta el traceback)
function _extractPyError(err) {
  const msg = String(err && err.message ? err.message : err);
  // Si es ExcelValidationError, viene como 'ExcelValidationError: msg\n' o anidado en traceback
  let m = msg.match(/ExcelValidationError:\s*(.+?)(?:\n|$)/);
  if (m) return { kind: 'validation', text: m[1].trim() };
  // Otros errores Python: tomar la última línea no vacía
  const lines = msg.split('\n').filter(l => l.trim());
  if (lines.length) {
    let last = lines[lines.length - 1];
    let n = last.match(/^[A-Z]\w+(?:Error|Exception):\s*(.+)/);
    if (n) return { kind: 'error', text: n[1].trim() };
    return { kind: 'error', text: last.trim() };
  }
  return { kind: 'error', text: msg };
}

function _renderWarnings(warnings) {
  const old = document.getElementById('ub-warnings');
  if (old) old.remove();
  if (!warnings || !warnings.length) return;
  const bar = document.getElementById('upload-bar');
  if (!bar) return;
  const div = document.createElement('div');
  div.id = 'ub-warnings';
  div.style.cssText = 'width:100%;margin-top:8px;padding:8px 12px;background:#FEF3E2;border:1px solid #E8C77D;border-radius:6px;font-size:11px;color:#6B5100;';
  div.innerHTML = '<strong>⚠ Advertencias del Excel:</strong><ul style="margin:4px 0 0 18px;padding:0;">'
    + warnings.map(w => '<li>' + w + '</li>').join('')
    + '</ul>';
  bar.appendChild(div);
}

// ── Resultados de metas (archivo separado, opcional) ────────────────────────
async function processMetasResults(file) {
  const py = await pyodideReady;
  const statusText = document.getElementById('upload-status-text');
  const input = document.getElementById('metas-input');
  try {
    const buf = await file.arrayBuffer();
    py.globals.set('metas_bytes_js', new Uint8Array(buf));
    const jsonStr = py.runPython(`
import json
metas_bytes = metas_bytes_js.to_bytes()
json.dumps(pipeline.compute_metas_results(metas_bytes), ensure_ascii=False, default=str)
`);
    const res = JSON.parse(jsonStr);
    window.__METAS_RESULTS__ = res;
    renderMetasAvance(res);
    statusText.innerHTML = '<span class="ub-check">✓ Resultados de metas cargados (' + file.name + ')</span>';
  } catch (err) {
    var info = _extractPyError(err);
    statusText.innerHTML = '<span class="ub-error">Error en resultados de metas: ' + info.text + '</span>';
    console.error(err);
  } finally {
    if (input) input.value = '';
  }
}

function _fmtB(v) {
  var b = v / 1e9;
  if (Math.abs(b) >= 1) return 'CLP ' + b.toFixed(2) + 'B';
  if (Math.abs(v) >= 1e6) return 'CLP ' + Math.round(v / 1e6) + 'M';
  return 'CLP ' + Math.round(v).toLocaleString('es-CL');
}

// Formato USD (los valores AUM vienen en USD)
function _fmtUSD(v) {
  var a = Math.abs(v);
  if (a >= 1e6) return 'US$ ' + (v / 1e6).toFixed(2) + 'M';
  if (a >= 1e3) return 'US$ ' + (v / 1e3).toFixed(0) + 'k';
  return 'US$ ' + Math.round(v).toLocaleString('en-US');
}

// Devuelve {txt, color} para un logro dado su target (maneja negativos)
function _logroDisplay(logro, target) {
  var neg = logro < 0;
  var pct = target > 0 ? Math.round(logro / target * 100) : 0;
  if (neg) return { txt: _fmtB(logro) + ' · rescate neto', color: '#B33A2E', pct: pct, neg: true };
  return {
    txt: _fmtB(logro) + ' · ' + pct + '%',
    color: pct >= 100 ? '#0E7A4E' : 'var(--text2)',
    pct: pct, neg: false,
  };
}

// Actualiza el trimestre `q` (ej "1T26") con metas + cumplimiento en USD.
// Lee todo de window.__METAS_RESULTS__.byPeriodo (no usa aumData CLP).
function _updateAumLogrosQ(q, targets) {
  var res = window.__METAS_RESULTS__;
  var metaIds = { deuda: 'q-deuda-meta', inm: 'q-inm-meta', inter: 'q-inter-meta' };
  var logIds  = { deuda: 'q-deuda-logro', inm: 'q-inm-logro', inter: 'q-inter-logro' };
  var bars    = { deuda: 'q-deuda-bar',   inm: 'q-inm-bar',   inter: 'q-inter-bar' };
  var pondIds = { deuda: 'q-deuda-pond', inm: 'q-inm-pond', inter: 'q-inter-pond' };
  var baseColor = { deuda: '#1B4B9B', inm: '#2E6BAF', inter: '#5A93CC' };
  var totalEl = document.getElementById('q-total-logro');
  var totalMetaEl = document.getElementById('q-total-meta');

  // Ocultar la card de Notas Estructuradas (ya no se mide)
  var notasMeta = document.getElementById('q-notas-meta');
  if (notasMeta) { var nc = notasMeta.closest('.aum-card'); if (nc) nc.style.display = res ? 'none' : ''; }

  var data = res && res.byPeriodo && res.byPeriodo[q];
  if (!data) {
    // Sin resultados cargados → dejar 'pendiente' y barras vacías
    ['deuda', 'inm', 'inter'].forEach(function(k) {
      var el = document.getElementById(logIds[k]);
      if (el) { el.textContent = 'pendiente'; el.style.color = 'var(--text2)'; }
      var bar = document.getElementById(bars[k]);
      if (bar) { bar.style.width = '0%'; bar.style.background = baseColor[k]; }
    });
    if (totalEl) { totalEl.textContent = '— pendiente'; totalEl.style.color = 'var(--text3)'; }
    return;
  }

  var totalMeta = 0, totalLogro = 0, anyLogro = false, sumCump = 0, sumPond = 0;
  ['deuda', 'inm', 'inter'].forEach(function(k) {
    var d = data[k]; if (!d) return;
    totalMeta += (d.target || 0);
    sumPond += (d.pond || 0);
    // Meta card en USD
    var me = document.getElementById(metaIds[k]);
    if (me) me.textContent = _fmtUSD(d.target || 0);
    // Badge de ponderación
    var pe = document.getElementById(pondIds[k]);
    if (pe && d.pond) pe.textContent = (d.pond * 100).toFixed(2) + '%';
    // Logro + barra
    var el = document.getElementById(logIds[k]);
    var bar = document.getElementById(bars[k]);
    if (d.logro === null || d.logro === undefined) {
      if (el) { el.textContent = 'corte pendiente'; el.style.color = 'var(--text2)'; }
      if (bar) { bar.style.width = '0%'; bar.style.background = baseColor[k]; }
      return;
    }
    anyLogro = true;
    totalLogro += d.logro;
    sumCump += (d.cump_pond || 0);
    var pct = Math.round((d.ratio || 0) * 100);
    var neg = d.logro < 0;
    if (el) {
      el.textContent = _fmtUSD(d.logro) + ' · ' + (neg ? 'rescate' : pct + '%');
      el.style.color = neg ? '#B33A2E' : (pct >= 100 ? '#0E7A4E' : 'var(--text2)');
    }
    if (bar) {
      var w = neg ? 0 : Math.max(0, Math.min(100, pct));
      bar.style.width = w + '%';
      bar.style.background = neg ? '#B33A2E' : (pct >= 100 ? '#0E7A4E' : baseColor[k]);
    }
    // Línea extra: variación vs corte anterior (momentum) + brecha vs meta
    var extraEl = document.getElementById('q-' + k + '-extra');
    if (extraEl) {
      var parts = [];
      if (d.delta_prev !== null && d.delta_prev !== undefined) {
        var dp = d.delta_prev;
        var arrow = dp > 0 ? '▲' : (dp < 0 ? '▼' : '–');
        var dpCol = dp > 0 ? '#0E7A4E' : (dp < 0 ? '#B33A2E' : 'var(--text3)');
        parts.push('<span style="color:' + dpCol + ';">' + arrow + ' ' + (dp >= 0 ? '+' : '') + _fmtUSD(dp) + ' vs corte ant.</span>');
      }
      if (d.gap !== null && d.gap !== undefined) {
        if (d.gap > 0) parts.push('<span style="color:var(--text3);">faltan ' + _fmtUSD(d.gap) + '</span>');
        else parts.push('<span style="color:#0E7A4E;">+' + _fmtUSD(-d.gap) + ' sobre meta</span>');
      }
      extraEl.innerHTML = parts.join(' &nbsp;·&nbsp; ');
      extraEl.style.display = parts.length ? 'block' : 'none';
    }
  });

  if (totalMetaEl) totalMetaEl.textContent = _fmtUSD(totalMeta);
  if (totalEl) {
    if (!anyLogro) { totalEl.textContent = 'corte pendiente'; totalEl.style.color = 'var(--text3)'; }
    else { totalEl.textContent = _fmtUSD(totalLogro); totalEl.style.color = totalLogro < 0 ? '#B33A2E' : (totalLogro >= totalMeta ? '#0E7A4E' : 'var(--text)'); }
  }
  var pctEl = document.getElementById('q-total-pct');
  if (pctEl) {
    if (!anyLogro) { pctEl.textContent = 'Corte del trimestre aún no disponible'; }
    else {
      var pctMeta = totalMeta ? Math.round(totalLogro / totalMeta * 100) : 0;
      var cumpPts = sumCump * 100;
      pctEl.innerHTML = (totalLogro < 0 ? 'rescate neto' : pctMeta + '% de la meta')
        + ' · aporte ponderado: <strong style="color:' + (cumpPts < 0 ? '#B33A2E' : '#0E7A4E') + ';">'
        + (cumpPts >= 0 ? '+' : '') + cumpPts.toFixed(2) + ' pts</strong>';
    }
  }
  var pondValEl = document.getElementById('aum-pond-value');
  var cumpSubEl = document.getElementById('aum-cump-sub');
  if (pondValEl) pondValEl.textContent = (sumPond * 100).toFixed(1) + '%';
  if (cumpSubEl) {
    if (!anyLogro) { cumpSubEl.textContent = 'Del scorecard total distribución'; }
    else {
      var cp = sumCump * 100;
      cumpSubEl.innerHTML = 'Del scorecard · cumplió <strong style="color:' + (cp < 0 ? '#B33A2E' : '#0E7A4E') + ';">'
        + (cp >= 0 ? '+' : '') + cp.toFixed(2) + ' pts</strong> en ' + q;
    }
  }
}

// Reconstruye los 2 charts AUM (barras trimestrales + línea acumulada) en USD
function _buildAumChartsUSD(res) {
  if (typeof Chart === 'undefined' || !res || !res.byPeriodo) return;
  var bp = res.byPeriodo;
  var cy = (window.__CURRENT_Q__ && window.__CURRENT_Q__.year) || 2026;
  var yy = String(cy).slice(-2);
  var pers = ['1T' + yy, '2T' + yy, '3T' + yy, '4T' + yy];
  var qLabels = ['Q1 ' + cy, 'Q2 ' + cy, 'Q3 ' + cy, 'Q4 ' + cy];
  var metas = ['deuda', 'inm', 'inter'];
  var mLabel = { deuda: 'Deuda Privada', inm: 'Inmobiliario', inter: 'Internacionales' };
  var mColor = { deuda: 'rgba(27,75,155,0.85)', inm: 'rgba(46,107,175,0.85)', inter: 'rgba(90,147,204,0.85)' };

  // Barras: meta USD (millones) por trimestre por instrumento (stacked)
  var elQ = document.getElementById('aumQChart');
  if (elQ) {
    var ex = Chart.getChart(elQ); if (ex) ex.destroy();
    var datasets = metas.map(function(k, i) {
      return {
        label: mLabel[k],
        data: pers.map(function(per) { var d = bp[per] && bp[per][k]; return d ? (d.target || 0) / 1e6 : 0; }),
        backgroundColor: mColor[k], stack: 's', borderSkipped: false,
        borderRadius: i === metas.length - 1 ? [3, 3, 0, 0] : 0
      };
    });
    window.aumQChartInst = new Chart(elQ, {
      type: 'bar',
      data: { labels: qLabels, datasets: datasets },
      options: {
        responsive: true, maintainAspectRatio: false, animation: { duration: 300 },
        plugins: { legend: { display: false }, tooltip: { mode: 'index', intersect: false, callbacks: {
          label: function(ctx) { return ' ' + ctx.dataset.label + ': US$ ' + ctx.parsed.y.toFixed(2) + 'M'; },
          footer: function(ctx) { return 'Total: US$ ' + ctx.reduce(function(a, c) { return a + c.parsed.y; }, 0).toFixed(2) + 'M'; }
        }}},
        scales: { x: Object.assign({ stacked: true }, xScale()), y: Object.assign({ stacked: true }, yScale(), {
          title: { display: true, text: 'US$ millones', color: TICK_C, font: { size: 9 } },
          ticks: { callback: function(v) { return 'US$' + v + 'M'; }, color: TICK_C, font: TICK_FONT } }) }
      }
    });
    var leg = document.getElementById('aumq-legend');
    if (leg) { leg.innerHTML = ''; metas.forEach(function(k) { leg.innerHTML += '<span class="legend-item"><span class="legend-dot" style="background:' + mColor[k] + '"></span>' + mLabel[k] + '</span>'; }); }
  }

  // Línea: meta acumulada vs logro acumulado (USD millones)
  var elY = document.getElementById('aumYtdChart');
  if (elY) {
    var ex2 = Chart.getChart(elY); if (ex2) ex2.destroy();
    var metaCum = [0], logroCum = [0], mc = 0, lc = 0, lValid = true;
    pers.forEach(function(per) {
      var dsum = 0, lsum = 0, hasL = false;
      metas.forEach(function(k) { var d = bp[per] && bp[per][k]; if (d) { dsum += (d.target || 0); if (d.logro != null) { lsum += d.logro; hasL = true; } } });
      mc += dsum; metaCum.push(mc / 1e6);
      if (hasL && lValid) { lc += lsum; logroCum.push(lc / 1e6); } else { lValid = false; logroCum.push(null); }
    });
    new Chart(elY, {
      type: 'line',
      data: { labels: ['Inicio', 'Cierre Q1', 'Cierre Q2', 'Cierre Q3', 'Cierre Q4'], datasets: [
        { label: 'Hito acumulado meta', data: metaCum, borderColor: '#1B4B9B', backgroundColor: 'rgba(27,75,155,0.08)', borderWidth: 2.5, fill: true, tension: 0.2, pointRadius: 5, pointBackgroundColor: '#1B4B9B', pointBorderColor: '#fff', pointBorderWidth: 2 },
        { label: 'Logro YTD', data: logroCum, borderColor: '#0E7A4E', backgroundColor: 'rgba(14,122,78,0.08)', borderWidth: 2.5, borderDash: [5, 3], fill: true, tension: 0.2, pointRadius: 5, pointBackgroundColor: '#0E7A4E', pointBorderColor: '#fff', pointBorderWidth: 2 }
      ]},
      options: {
        responsive: true, maintainAspectRatio: false, animation: { duration: 300 },
        plugins: { legend: { display: false }, tooltip: { callbacks: { label: function(ctx) { return ctx.parsed.y === null ? ctx.dataset.label + ': pendiente' : ' ' + ctx.dataset.label + ': US$ ' + ctx.parsed.y.toFixed(2) + 'M'; } } } },
        scales: { x: xScale(), y: Object.assign(yScale(), {
          title: { display: true, text: 'US$ millones (acum.)', color: TICK_C, font: { size: 9 } },
          ticks: { callback: function(v) { return 'US$' + v + 'M'; }, color: TICK_C, font: TICK_FONT } }) }
      }
    });
  }
}

// Actualiza el sub-panel anual (YTD): KPI + celdas de la tabla (USD)
function _updateAumYTD() {
  var res = window.__METAS_RESULTS__;
  var data = res && res.byPeriodo && res.byPeriodo['2026'];
  var ids = { deuda: 'ytd-deuda-logro', inm: 'ytd-inm-logro', inter: 'ytd-inter-logro' };
  var kpiEl = document.getElementById('ytd-kpi-logro');
  var totalEl = document.getElementById('ytd-total-logro');
  // Ocultar fila de Notas en la tabla YTD
  var notasCell = document.getElementById('ytd-notas-logro');
  if (notasCell && res) { var tr = notasCell.closest('tr'); if (tr) tr.style.display = 'none'; }
  if (!data) return;
  var totalLogro = 0;
  ['deuda', 'inm', 'inter'].forEach(function(k) {
    var el = document.getElementById(ids[k]);
    var d = data[k];
    if (d && d.logro != null) totalLogro += d.logro;
    if (!el || !d) return;
    if (d.logro == null) { el.textContent = 'corte pendiente'; el.style.fontStyle = 'normal'; return; }
    var pct = Math.round((d.ratio || 0) * 100);
    var neg = d.logro < 0;
    el.textContent = _fmtUSD(d.logro) + ' · ' + (neg ? 'rescate' : pct + '%');
    el.style.color = neg ? '#B33A2E' : (pct >= 100 ? '#0E7A4E' : 'var(--text2)');
    el.style.fontStyle = 'normal';
  });
  if (totalEl) {
    totalEl.textContent = _fmtUSD(totalLogro);
    totalEl.style.color = totalLogro < 0 ? '#B33A2E' : 'var(--text)';
    totalEl.style.fontStyle = 'normal';
  }
  if (kpiEl) {
    kpiEl.textContent = _fmtUSD(totalLogro);
    kpiEl.style.color = totalLogro < 0 ? '#B33A2E' : 'var(--green)';
  }
}

// Refresca todo el tab AUM con los resultados cargados
function renderMetasAvance(res) {
  // Re-disparar el trimestre activo (para actualizar las cards trimestrales)
  var activeQ = document.querySelector('[id^="qsel-"].active');
  if (activeQ && typeof selectQuarter === 'function') {
    var qm = activeQ.getAttribute('onclick').match(/selectQuarter\('([^']+)'/);
    if (qm) selectQuarter(qm[1], activeQ);
  }
  _updateAumYTD();
  // Cruce saldo vs reuniones (necesita saldoPorAdvisor de este archivo +
  // reunPorEntidad del Excel principal ya cargado).
  if (typeof renderCrossSaldoReun === 'function') renderCrossSaldoReun();
  // Reconstruir los charts AUM en USD. Si el tab está visible, ya; si no,
  // se reconstruyen al entrar (initAumCharts chequea __METAS_RESULTS__).
  window.aumChartsInited = false;
  var aumPanel = document.getElementById('tab-metas-aum');
  if (aumPanel && aumPanel.offsetParent !== null && typeof initAumCharts === 'function') {
    initAumCharts();
  }
  // Actualizar el banner de aviso para reflejar que ya hay datos cargados
  var banner = document.getElementById('aum-aviso-text');
  if (banner && res && res.corte) {
    banner.innerHTML = '✓ <strong>Resultados reales cargados</strong> (fuente Britech, corte ' + res.corte +
      '). Las barras y los logros reflejan el avance real. Valores negativos = rescate neto. ' +
      'La data vive sólo en tu navegador.';
  }
}

async function processExcel(file) {
  const py = await pyodideReady;
  const statusText = document.getElementById('upload-status-text');
  const spinner    = document.getElementById('upload-spinner');
  const btn        = document.getElementById('upload-btn');
  const btnIcon    = document.getElementById('upload-btn-icon');
  const btnText    = document.getElementById('upload-btn-text');
  const input      = document.getElementById('xlsm-input');
  spinner.style.display = '';
  btn.disabled = true;
  statusText.innerHTML = '<span class="ub-info">Procesando ' + file.name + '…</span>';
  _renderWarnings([]); // limpiar warnings previos

  try {
    const buf = await file.arrayBuffer();
    py.globals.set('xlsm_bytes_js', new Uint8Array(buf));
    const jsonStr = py.runPython(`
import json
xlsm_bytes = xlsm_bytes_js.to_bytes()
data = pipeline.compute_data(xlsm_bytes)
json.dumps(data, ensure_ascii=False, default=str)
`);
    const data = JSON.parse(jsonStr);
    applyData(data);
    document.body.classList.remove('no-data');
    spinner.style.display = 'none';
    const sizeKB = (file.size / 1024).toFixed(0);
    const now = new Date();
    const hhmm = String(now.getHours()).padStart(2,'0') + ':' + String(now.getMinutes()).padStart(2,'0');
    statusText.innerHTML =
      '<span class="ub-check">✓ ' + file.name + '</span>' +
      ' <span class="ub-meta">' + sizeKB + ' KB · ' + (data.meta && data.meta.n_clientes || '?') + ' clientes · corte ' + (data.meta && data.meta.fecha_corte || '?') + ' · cargado ' + hhmm + '</span>';
    btnIcon.textContent = '↻';
    btnText.textContent = 'Actualizar Excel';
    _renderWarnings(data.meta && data.meta.warnings);
  } catch (err) {
    spinner.style.display = 'none';
    const info = _extractPyError(err);
    if (info.kind === 'validation') {
      statusText.innerHTML =
        '<span class="ub-error">✗ El Excel no tiene la estructura esperada</span> '
        + '<span class="ub-info" style="display:block;margin-top:4px;">' + info.text + '</span>';
    } else {
      statusText.innerHTML = '<span class="ub-error">Error procesando Excel: ' + info.text + '</span>';
    }
    console.error(err);
  } finally {
    btn.disabled = false;
    if (input) input.value = '';
  }
}

// ══════════════════════════════════════════════════════════════════════════
// applyData(d) — toma el dict del pipeline e inyecta TODO en el dashboard
// ══════════════════════════════════════════════════════════════════════════

function applyData(d) {
  // 1) Globales que el resto del JS consume
  // `data` se setea como objeto con keys numéricas para compat con código viejo (data[2026]).
  var dataNumKeys = {};
  Object.keys(d.data || {}).forEach(function(k){ dataNumKeys[k] = d.data[k]; dataNumKeys[parseInt(k,10)] = d.data[k]; });
  window.data              = dataNumKeys;
  window.prioData          = d.prioData;
  window.clientsData       = d.clientsData;
  window.recontactData     = d.recontactData;
  window.riskAll           = d.riskAll;
  window.criticalList      = d.criticalList;
  window.monthlyDetailData = d.monthlyDetailData;
  window.subtemaMoData     = d.subtemaMoData;
  window.subteamAnnual     = d.subteamAnnual;
  window.subtemasColors    = d.subtemasColors;
  window.subtemasClientsM  = d.subtemasClientsM;
  window.subtemasClientsA  = d.subtemasClientsA;
  window.riskData2026      = d.riskData2026;
  window.activationData    = d.activationData    || {};
  window.areasData         = d.areasData         || [];
  window.__REUN_POR_ENTIDAD__ = d.reunPorEntidad || {};
  window.fotoYTD           = d.fotoYTD || { year:0, prevYear:0, month:0, monthNames:[], rows:[] };
  // Estos quedan idénticos a los inline (no son data sensible — son aliases JS):
  window.data2024 = d.data['2024']; window.data2025 = d.data['2025']; window.data2026 = d.data['2026'];

  // 2) Re-renderizar el line chart usando los 3 años más recientes disponibles
  // (en 2027 esto pasará a mostrar 2025/2026/2027 automáticamente).
  try {
    if (typeof lineChart !== 'undefined' && lineChart && lineChart.data && d.years) {
      var yrs3 = d.years.slice(-3);
      for (var i = 0; i < 3; i++) {
        var yr = yrs3[i];
        if (yr == null) continue;
        var ds = lineChart.data.datasets[i];
        if (!ds) continue;
        ds.label = String(yr);
        ds.data = d.data[String(yr)] || [];
      }
      lineChart.update();
    }
  } catch (e) { console.warn('lineChart update fallo', e); }

  // 3) KPIs + labels Q-rolling
  try {
    var k = d.kpi || {};
    var cq = d.currentQ || { label: 'Q1', year: 2026, q: 1, label_year_prev: 'Q1' };
    var cy = d.currentYear || 2026;
    // Guardar currentQ globalmente para que renderMetasAvance lo use.
    window.__CURRENT_Q__ = cq;
    // Si ya había resultados de metas cargados, re-renderizar con el Q actual.
    if (window.__METAS_RESULTS__ && typeof renderMetasAvance === 'function') {
      try { renderMetasAvance(window.__METAS_RESULTS__); } catch (e) {}
    }

    // KPI principal del trimestre: muestra el total del Q actual completo
    var elQ1 = document.querySelector('[data-kpi="q1-value"]');
    if (elQ1) elQ1.textContent = (cq.reuniones != null ? cq.reuniones : k.q1);

    // KPI Reuniones 2026 YTD: número
    var elYtd = document.querySelector('[data-kpi="ytd-value"]');
    if (elYtd) elYtd.textContent = k.ytd;

    // Sub del YTD: muestra hasta el mes anterior con su valor
    var elYtdSub = document.querySelector('[data-kpi="ytd-sub"]');
    if (elYtdSub) {
      var MES_ABBR = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
      var dCur = d.data && d.data[String(cy)] ? d.data[String(cy)] : [];
      var parts = [];
      // Mostrar Q del trimestre completo + 1 o 2 meses posteriores con datos
      parts.push(cq.label.split(' ')[0] + ': ' + (cq.reuniones || 0));
      var startM = cq.q * 3;  // primer mes después del Q completo (idx 0-based en MES_ABBR)
      for (var i = startM; i < 12 && dCur[i] > 0; i++) {
        parts.push(MES_ABBR[i] + ': ' + dCur[i]);
      }
      elYtdSub.innerHTML = parts.join(' &nbsp;&middot;&nbsp; ');
    }

    // Labels que rolan con el trimestre
    var setText = function(sel, text) {
      document.querySelectorAll(sel).forEach(function(el) { el.textContent = text; });
    };
    setText('[data-q-section="vision-general"]', 'Visión general ' + cq.label);
    setText('[data-q-text="reuniones-q-label"]', 'Reuniones ' + cq.label);
    setText('[data-q-text="reuniones-ytd-label"]', 'Reuniones ' + cy + ' YTD');
    setText('[data-q-text="top-clientes-q-card"]', 'Top clientes ' + cq.label + ' por segmento');

    // Sub "vs N en QX YYYY-1 (+P%)"
    var elVs = document.querySelector('[data-q-text="reuniones-q-vs"]');
    if (elVs && cq.reuniones_prev != null) {
      var sign = (cq.pct_change || 0) >= 0 ? '+' : '';
      elVs.textContent = 'vs ' + cq.reuniones_prev + ' en ' + cq.label_year_prev + ' (' + sign + (cq.pct_change || 0) + '%)';
    }

    // Header badge: "Q1 2026 · Mar 2026" → "Q? YYYY · Mes YYYY"
    var elBadge = document.getElementById('update-badge');
    if (elBadge && d.meta && d.meta.fecha_corte) {
      var MES_FULL = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
      var parts2 = d.meta.fecha_corte.split('-');
      var mesIdx = parseInt(parts2[1], 10) - 1;
      elBadge.textContent = cq.label + ' · ' + MES_FULL[mesIdx] + ' ' + parts2[0];
    }

    // Header brand "Dashboard de seguimiento · YYYY – YYYY"
    var yrsAsc = (d.years || [2024,2025,2026]).slice();
    setText('[data-q-text="header-years"]', 'Dashboard de seguimiento · ' + yrsAsc[0] + ' – ' + yrsAsc[yrsAsc.length - 1]);

    // Recontact section label: prev_year → current_year YTD
    setText('[data-q-text="recontact-section"]', 'Tasa de recontacto ' + (cy - 1) + ' → ' + cy + ' YTD · click para ver detalle');

    // AUM data + qMetas (de la hoja Metas{cur_year}). Reset del flag para
    // que initAumCharts() reconstruya con la data nueva en próxima vista.
    if (d.aumData)  window.aumData  = d.aumData;
    if (d.qMetas)   window.qMetas   = d.qMetas;
    window.aumChartsInited = false;

    // Nuevos distribuidores
    var nd = d.nuevosDistribuidores || [];
    var ndGoal = (d.metasGoals && d.metasGoals.nuevos_distribuidores_anual) || null;
    var ndKpiVal = document.querySelector('[data-q-text="nuevos-kpi-value"]');
    if (ndKpiVal) ndKpiVal.textContent = nd.length;
    setText('[data-q-text="nuevos-kpi-label"]', 'Nuevos distribuidores ' + cy);
    var ndSubEl = document.querySelector('[data-q-text="nuevos-kpi-sub"]');
    if (ndSubEl) {
      var names = nd.slice(0, 4).map(function(n){return n.name;}).join(' · ');
      var goalTxt = ndGoal ? ' · Meta anual: ' + ndGoal : '';
      ndSubEl.textContent = (names || 'sin nuevos a la fecha') + goalTxt;
    }
    setText('[data-q-text="nuevos-section"]', 'Nuevos distribuidores ' + cy + ' — detalle');
    setText('[data-q-text="nuevos-footer"]',
      '* Cliente nuevo = distribuidor presente en el foto ' + cy + ' que no estaba en ' + (cy - 1) +
      (ndGoal ? ' · Meta anual: ' + ndGoal + ' nuevos (' + Math.round(ndGoal / 4) + ' por trimestre)' : ''));
    var ndTbody = document.getElementById('nuevos-tbody');
    if (ndTbody) {
      if (nd.length === 0) {
        // Sin nuevos: esconder la sección entera
        var sectLbl = document.getElementById('nuevos-section-label');
        if (sectLbl) sectLbl.style.display = 'none';
        var card = ndTbody.closest('.card');
        if (card) card.style.display = 'none';
      } else {
        var sectLbl2 = document.getElementById('nuevos-section-label');
        if (sectLbl2) sectLbl2.style.display = '';
        var card2 = ndTbody.closest('.card');
        if (card2) card2.style.display = '';
        ndTbody.innerHTML = nd.map(function(n, i) {
          var isLast = (i === nd.length - 1);
          return '<tr' + (isLast ? '' : ' style="border-bottom:1px solid var(--border);"') + '>'
            + '<td style="padding:8px 12px;font-weight:500;color:var(--text);">' + n.name + '</td>'
            + '<td style="text-align:center;padding:8px 12px;"><span style="background:#EBF0F8;color:#1B4B9B;font-size:10px;font-weight:600;padding:2px 8px;border-radius:4px;">'
            + (n.fondo || '—') + '</span></td>'
            + '<td style="text-align:center;padding:8px 12px;font-family:\'IBM Plex Mono\',monospace;color:#0E7A4E;font-weight:600;">'
            + (n.q || '—') + '</td>'
            + '</tr>';
        }).join('');
      }
    }

    // Labels Q/año restantes (Consolidado, Semáforo, Footer, AUM, setYoy)
    setText('[data-q-text="nuevos-card-title"]', 'Clientes nuevos incorporados en ' + cy);
    setText('[data-q-text="comparativa-section"]', 'Comparativa ' + cq.label + ' vs ' + cq.label_year_prev + ' — por prioridad');
    setText('[data-q-text="crosssell-section"]', 'Oportunidades de cross-sell — alto contacto, pocos productos');
    setText('[data-q-text="crosssell-title"]', 'Clientes con ≥ 3 reuniones en ' + cy + ' y ≤ 1 producto distinto discutido (fuera de Follow Up)');
    setText('[data-q-text="consolidado-section"]', 'Consolidado ' + cq.label + ' — progreso del año por segmento');
    setText('[data-q-text="semaforo-title"]', 'Semáforo de contacto ' + cy + ' YTD — por segmento');
    setText('[data-q-text="kpi-q-current"]', cq.label);
    setText('[data-q-text="footer-date"]', cq.label + ' · ' + yrsAsc[0] + '–' + yrsAsc[yrsAsc.length - 1]);

    // Tabla Consolidado: marcar Q completados con ✓, Q en curso con (YTD)
    for (var qi = 1; qi <= 4; qi++) {
      var el = document.querySelector('[data-q-th="' + qi + '"]');
      if (!el) continue;
      var lbl = 'Q' + qi + ' ' + cy;
      if (qi < cq.q || (qi === cq.q && cq.q < 4)) lbl += ' ✓';
      else if (qi === cq.q + 1) lbl += ' (YTD)';
      el.textContent = lbl;
    }

    // AUM: actualizar quarters dinámicos para charts y labels visibles
    var yy = String(cy).slice(-2);
    window.AUM_QUARTERS = ['Q1 ' + cy, 'Q2 ' + cy, 'Q3 ' + cy, 'Q4 ' + cy];
    var qCurEl = document.getElementById('q-current-label');
    if (qCurEl) qCurEl.textContent = cq.label;
    var qSecEl = document.getElementById('q-section-label');
    if (qSecEl) qSecEl.textContent = 'Desglose por instrumento · ' + cq.label;

    // setYoy buttons: rebuild dinámico con los últimos 3 años → 2 pares
    var yoyContainer = document.querySelector('[data-q-buttons="yoy-pairs"]');
    if (yoyContainer && yrsAsc.length >= 2) {
      var last3 = yrsAsc.slice(-3);
      var pairs = [];
      if (last3.length >= 3) pairs.push([last3[0], last3[1]]);
      pairs.push([last3[last3.length - 2], last3[last3.length - 1]]);
      yoyContainer.innerHTML = pairs.map(function(p, i) {
        var active = (i === pairs.length - 1) ? ' active' : '';
        var key = p[0] + '-' + p[1];
        return '<button class="ftab' + active + '" onclick="setYoy(\'' + key + '\',this)">' + p[0] + ' vs ' + p[1] + '</button>';
      }).join('');
      // Disparar render inicial con el par más reciente
      var activeYoyBtn = yoyContainer.querySelector('.active');
      if (activeYoyBtn && typeof setYoy === 'function') {
        var k = activeYoyBtn.getAttribute('onclick').match(/setYoy\('([^']+)'/);
        if (k) setYoy(k[1], activeYoyBtn);
      }
    }

    // Activación de cartera (sección + 4 KPIs)
    setText('[data-q-text="activacion-section"]', 'Activación de cartera ' + cq.label + ' · click para ver detalle');
    setText('[data-q-text="reuniones-area-section"]', 'Reuniones ' + cq.label + ' por área / producto · click para ver clientes');
    if (d.activationData) {
      ['nunca','nuevos','reactiv','inactivos'].forEach(function(t){
        var info = d.activationData[t]; if (!info) return;
        setText('[data-act-label="' + t + '"]', info.kpi_label || '');
        var elV = document.querySelector('[data-act-value="' + t + '"]');
        if (elV) elV.textContent = (info.items || []).length;
        setText('[data-act-sub="' + t + '"]', info.kpi_sub || '');
      });
    }
  } catch (e) { console.warn('KPI/Q-label update fallo', e); }

  // 4) Recontact bars (texto + barra + %)
  try {
    ['hp', 'mp', 'lp'].forEach(function (pl) {
      var sub = d.recontactData[pl];
      if (!sub) return;
      var yes = sub.yes.length, no = sub.no.length, tot = yes + no;
      var pct = tot > 0 ? Math.round(yes / tot * 100) : 0;
      var elText = document.querySelector('[data-recontact-text="' + pl + '"]');
      var elBar  = document.querySelector('[data-recontact-bar="' + pl + '"]');
      var elPct  = document.querySelector('[data-recontact-pct="' + pl + '"]');
      if (elText) elText.textContent = yes + ' de ' + tot + ' recontactados';
      if (elBar)  elBar.style.width = pct + '%';
      if (elPct)  elPct.textContent = pct + '%';
    });
  } catch (e) { console.warn('recontact update fallo', e); }

  // 5) Tab Riesgo — KPIs y títulos de columna (HP/MP/LP — N clientes)
  try {
    var r = d.kpi.riesgo_totals || {}, p = d.kpi.riesgo_by_prio || {};
    var setText = function (sel, val) { var el = document.querySelector(sel); if (el) el.textContent = val; };
    setText('[data-riesgo="activo-value"]',    r.activo);
    setText('[data-riesgo="pendiente-value"]', r.pendiente);
    setText('[data-riesgo="inactivo-value"]',  r.inactivo);
    setText('[data-riesgo="total-value"]',     r.total);
    setText('[data-riesgo="activo-sub"]',      'HP:' + (p.HP||{}).activo    + ' · MP:' + (p.MP||{}).activo    + ' · LP:' + (p.LP||{}).activo);
    setText('[data-riesgo="pendiente-sub"]',   'HP:' + (p.HP||{}).pendiente + ' · MP:' + (p.MP||{}).pendiente + ' · LP:' + (p.LP||{}).pendiente);
    setText('[data-riesgo="inactivo-sub"]',    'HP:' + (p.HP||{}).inactivo  + ' · MP:' + (p.MP||{}).inactivo  + ' · LP:' + (p.LP||{}).inactivo);
    setText('[data-riesgo="total-sub"]',       'HP:' + (p.HP||{}).total     + ' · MP:' + (p.MP||{}).total     + ' · LP:' + (p.LP||{}).total);
    setText('[data-riesgo-col="HP"]', 'HP — ' + (p.HP||{}).total + ' clientes');
    setText('[data-riesgo-col="MP"]', 'MP — ' + (p.MP||{}).total + ' clientes');
    setText('[data-riesgo-col="LP"]', 'LP — ' + (p.LP||{}).total + ' clientes');
  } catch (e) { console.warn('riesgo update fallo', e); }

  // 6) Charts anónimos del Resumen — destroy + recreate
  try {
    if (typeof Chart !== 'undefined') {
      rebuildSemaforoChart(d.kpi.semaforo);
      rebuildPeakChart(d.recentLabels, d.recentVals);
      rebuildReunYtdChart(d.kpi.reun_ytd_by_prio, d.kpi.riesgo_by_prio);
    }
  } catch (e) { console.warn('rebuild charts fallo', e); }

  // 7) Re-renderizar las secciones que leen las variables
  try {
    var btnAll = document.querySelector('[onclick*="setClients(\'all\'"]')
              || document.querySelector('[onclick*="setClients"]');
    if (btnAll && typeof setClients === 'function') setClients('all', btnAll);

    var btnHP = document.querySelector('[onclick*="setPrioClients(\'HP\'"]')
             || document.querySelector('[onclick*="setPrioClients"]');
    if (btnHP && typeof setPrioClients === 'function') setPrioClients('HP', btnHP);

    // Reconstruir tabs de año dinámicamente según data.years (multi-año)
    rebuildYearTabs(d);

    if (typeof _renderCriticalList === 'function') _renderCriticalList();
    if (typeof _renderAreas === 'function')        _renderAreas();
    renderComparativaQ(d);
    renderCrossSell(d);
    renderCrossGastoReun(d);
    renderCrossSaldoReun();
    if (typeof renderRiesgo2026 === 'function')    renderRiesgo2026();
    if (typeof renderMonthly === 'function')       renderMonthly();
    if (typeof renderSubtemas === 'function')      renderSubtemas();
    if (typeof renderFoto === 'function')          renderFoto();
    if (typeof buildHeatmap === 'function')        buildHeatmap();
  } catch (e) { console.warn('re-render fallo', e); }
}

// ══════════════════════════════════════════════════════════════════════════
// Comparativa Q YoY por prioridad
// ══════════════════════════════════════════════════════════════════════════

function renderComparativaQ(d) {
  var el = document.getElementById('comparativa-q-yoy');
  if (!el) return;
  var c = d.qComparativa || {};
  el.innerHTML = ['HP', 'MP', 'LP'].map(function(p) {
    var v = c[p] || { cur: 0, prev: 0, pct: 0 };
    var dir = v.pct > 0 ? '↑' : (v.pct < 0 ? '↓' : '−');
    var col = v.pct > 0 ? '#0E7A4E' : (v.pct < 0 ? '#B33A2E' : '#7D93B5');
    var bg  = v.pct > 0 ? '#E0F4EC' : (v.pct < 0 ? '#FAE9E7' : 'var(--card2)');
    var sign = v.pct > 0 ? '+' : '';
    var pCol = (p === 'HP') ? '#1B4B9B' : (p === 'MP' ? '#2E6BAF' : '#5A93CC');
    return '<div style="padding:14px;border-radius:8px;background:var(--card2);border-left:3px solid ' + pCol + ';">'
      + '<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px;">'
      +   '<span style="font-size:11px;font-weight:700;color:' + pCol + ';text-transform:uppercase;letter-spacing:.06em;">' + p + '</span>'
      +   '<span style="font-size:12px;font-weight:700;color:' + col + ';padding:2px 8px;border-radius:4px;background:' + bg + ';">' + dir + ' ' + sign + v.pct + '%</span>'
      + '</div>'
      + '<div style="display:flex;align-items:baseline;gap:8px;">'
      +   '<span style="font-family:monospace;font-size:24px;font-weight:700;color:var(--text);">' + v.cur + '</span>'
      +   '<span style="font-size:11px;color:var(--text3);">reuniones</span>'
      + '</div>'
      + '<div style="font-size:10px;color:var(--text3);margin-top:4px;">vs <strong>' + v.prev + '</strong> año anterior</div>'
      + '</div>';
  }).join('');
}

// ══════════════════════════════════════════════════════════════════════════
// Cross-sell radar (alto contacto, pocos productos)
// ══════════════════════════════════════════════════════════════════════════

function renderCrossSell(d) {
  var el = document.getElementById('cross-sell-list');
  if (!el) return;
  var list = d.crossSellRadar || [];
  if (!list.length) {
    el.innerHTML = '<div style="padding:14px;text-align:center;color:var(--text3);font-size:11px;">Sin oportunidades detectadas — todos los clientes con ≥3 reuniones ya discutieron varios productos</div>';
    return;
  }
  var pC = { HP: '#1B4B9B', MP: '#2E6BAF', LP: '#5A93CC' };
  el.innerHTML = list.map(function(c, i) {
    var col = pC[c.prio] || '#7D93B5';
    var prods = c.products && c.products.length ? c.products.join(' · ') : '<em style="color:var(--text3);">solo Follow Up / Catch Up</em>';
    return '<div style="display:flex;align-items:center;gap:10px;padding:8px 0;border-bottom:1px solid var(--border);">'
      + '<span style="font-family:monospace;font-size:10px;color:var(--text3);min-width:20px;text-align:right;">' + (i + 1) + '</span>'
      + '<span style="font-size:9px;font-weight:600;padding:2px 5px;border-radius:3px;background:' + col + '18;color:' + col + ';min-width:26px;text-align:center;">' + c.prio + '</span>'
      + '<div style="flex:1;min-width:0;">'
      +   '<div style="font-size:12px;font-weight:500;color:var(--text);">' + c.name + '</div>'
      +   '<div style="font-size:10px;color:var(--text3);margin-top:2px;">Producto: ' + prods + '</div>'
      + '</div>'
      + '<div style="text-align:right;flex-shrink:0;min-width:60px;">'
      +   '<span style="font-family:monospace;font-size:14px;font-weight:700;color:var(--text);">' + c.reun + '</span>'
      +   '<div style="font-size:9px;color:var(--text3);">reuniones</div>'
      + '</div>'
      + '</div>';
  }).join('');
}

// ══════════════════════════════════════════════════════════════════════════
// Cruces por distribuidor (saldo vs reuniones · gasto vs reuniones)
// ══════════════════════════════════════════════════════════════════════════

// Formato CLP compacto (B / M / k)
function _fmtCLP(v) {
  var a = Math.abs(v);
  if (a >= 1e9) return 'CLP ' + (v / 1e9).toFixed(2) + 'B';
  if (a >= 1e6) return 'CLP ' + (v / 1e6).toFixed(1) + 'M';
  if (a >= 1e3) return 'CLP ' + Math.round(v / 1e3) + 'k';
  return 'CLP ' + Math.round(v).toLocaleString('es-CL');
}

function _xEmpty(msg) {
  return '<div style="padding:18px 14px;text-align:center;color:var(--text3);font-size:11px;line-height:1.5;">' + msg + '</div>';
}

// Saldo (MarketValue total) vs reuniones — join saldoPorAdvisor x reunPorEntidad.
function renderCrossSaldoReun() {
  var el = document.getElementById('cross-saldo-reun');
  if (!el) return;
  var res  = window.__METAS_RESULTS__;
  var reun = window.__REUN_POR_ENTIDAD__ || {};
  var saldo = res && res.saldoPorAdvisor;
  if (!saldo || !Object.keys(saldo).length || !Object.keys(reun).length) {
    el.innerHTML = _xEmpty('Carga el Excel <strong>Follow Up clientes</strong> y el archivo de <strong>resultados de metas</strong> para ver el cruce.');
    return;
  }
  var rows = [];
  Object.keys(saldo).forEach(function(n) {
    if (reun[n]) rows.push({ nombre: reun[n].nombre, saldo: saldo[n].saldoUSD, reun: reun[n].reun });
  });
  rows.sort(function(a, b) { return b.saldo - a.saldo; });
  if (!rows.length) {
    el.innerHTML = _xEmpty('Ningún distribuidor está simultáneamente en saldos por corte y con reuniones registradas.');
    return;
  }
  var maxV = rows[0].saldo || 1;
  var corte = res.saldoCorteLabel || '';
  var head = '<div class="xhead"><span>' + rows.length + ' distribuidores</span>'
    + (corte ? '<span>MarketValue total · corte ' + corte + '</span>' : '') + '</div>';
  var body = rows.map(function(c, i) {
    var w = Math.max(3, Math.round(c.saldo / maxV * 100));
    return '<div class="xrow">'
      + '<span class="xrank' + (i === 0 ? ' top' : '') + '">' + (i + 1) + '</span>'
      + '<div style="flex:1;min-width:0;">'
      +   '<div style="display:flex;justify-content:space-between;align-items:baseline;gap:10px;">'
      +     '<span class="xname">' + c.nombre + '</span>'
      +     '<span class="xval">' + _fmtUSD(c.saldo) + '</span>'
      +   '</div>'
      +   '<div class="xbar-track"><div class="xbar-fill" style="width:' + w + '%;background:linear-gradient(90deg,#5A93CC,#1B4B9B);"></div></div>'
      + '</div>'
      + '<span class="xpill">' + c.reun + ' <small>reun</small></span>'
      + '</div>';
  }).join('');
  el.innerHTML = head + body;
}

// Gasto vs reuniones por distribuidor (todo computado en compute_data).
function renderCrossGastoReun(d) {
  var el = document.getElementById('cross-gasto-reun');
  if (!el) return;
  var list = (d && d.crossGastoReun) || (window.__CROSS_GASTO_REUN__ || []);
  window.__CROSS_GASTO_REUN__ = list;
  if (!list.length) {
    el.innerHTML = _xEmpty('Carga el Excel <strong>Follow Up clientes</strong> para ver el gasto vs reuniones por distribuidor.');
    return;
  }
  var maxV = Math.max.apply(null, list.map(function(r) { return r.gasto; }).concat([1]));
  var head = '<div class="xhead"><span>' + list.length + ' distribuidores</span>'
    + '<span>Gasto · reuniones · costo por reunión</span></div>';
  var body = list.map(function(c, i) {
    var w = Math.max(3, Math.round(c.gasto / maxV * 100));
    var reunPill = c.reun > 0
      ? '<span class="xpill">' + c.reun + ' <small>reun</small></span>'
      : '<span class="xpill zero">0 <small>reun</small></span>';
    var cpr;
    if (c.gastoPorReun == null) {
      cpr = '<span class="xsub" style="color:var(--text3);">sin reuniones</span>';
    } else {
      var v = c.gastoPorReun;
      var col = v < 30000 ? '#0E7A4E' : (v < 100000 ? '#D4890A' : '#B33A2E');
      cpr = '<span class="xsub" style="color:' + col + ';font-weight:700;">' + _fmtCLP(v) + ' c/u</span>';
    }
    return '<div class="xrow">'
      + '<span class="xrank' + (i === 0 ? ' top' : '') + '">' + (i + 1) + '</span>'
      + '<div style="flex:1;min-width:0;">'
      +   '<div style="display:flex;justify-content:space-between;align-items:baseline;gap:10px;">'
      +     '<span class="xname">' + c.nombre + '</span>'
      +     '<span class="xval">' + _fmtCLP(c.gasto) + '</span>'
      +   '</div>'
      +   '<div class="xbar-track"><div class="xbar-fill" style="width:' + w + '%;background:linear-gradient(90deg,#E0A23A,#D4890A);"></div></div>'
      + '</div>'
      + '<div style="display:flex;flex-direction:column;align-items:flex-end;gap:2px;flex-shrink:0;min-width:86px;">'
      +   reunPill + cpr
      + '</div>'
      + '</div>';
  }).join('');
  el.innerHTML = head + '<div class="xscroll">' + body + '</div>';
}

// ══════════════════════════════════════════════════════════════════════════
// Rebuild de tabs de año (Vista mensual + Subtemas) — multi-año
// ══════════════════════════════════════════════════════════════════════════

function rebuildYearTabs(d) {
  var yrs = (d.years || [2024, 2025, 2026]).slice().reverse(); // más reciente primero
  var cy = d.currentYear || yrs[0];

  // Vista mensual — año
  var vm = document.querySelector('[data-q-buttons="vista-year"]');
  if (vm) {
    vm.innerHTML = yrs.map(function(yr) {
      var active = (yr === cy) ? ' active' : '';
      return '<button class="ftab' + active + '" data-yr="' + yr + '" '
           + 'onclick="setMonthYear(\'' + yr + '\',this)">' + yr + '</button>';
    }).join('');
  }

  // Subtemas — año (con "Todos" al final)
  var st = document.querySelector('[data-q-buttons="subtema-year"]');
  if (st) {
    var btns = yrs.map(function(yr) {
      var active = (yr === cy) ? ' active' : '';
      return '<button class="ftab' + active + '" data-sub-yr="' + yr + '" '
           + 'onclick="setSubYear(\'' + yr + '\',this)">' + yr + '</button>';
    });
    btns.push('<button class="ftab" data-sub-yr="ALL" onclick="setSubYear(\'ALL\',this)">Todos</button>');
    st.innerHTML = btns.join('');
  }

  // Heatmap: 3 años más recientes (ascendente)
  window.HEATMAP_YEARS = yrs.slice(0, 3).reverse();

  // Disparar render inicial en el año actual
  var elM = document.querySelector('[data-yr="' + cy + '"]');
  if (elM && typeof setMonthYear === 'function') setMonthYear(String(cy), elM);
  var elS = document.querySelector('[data-sub-yr="' + cy + '"]');
  if (elS && typeof setSubYear === 'function') setSubYear(String(cy), elS);
}

// ══════════════════════════════════════════════════════════════════════════
// Rebuild de charts anónimos (Resumen)
// ══════════════════════════════════════════════════════════════════════════

function rebuildSemaforoChart(s) {
  var el = document.getElementById('semaforoChart');
  if (!el || typeof Chart === 'undefined') return;
  var existing = Chart.getChart(el);
  if (existing) existing.destroy();
  new Chart(el, {
    type: 'bar',
    data: {
      labels: ['HP (' + s.HP.total + ')', 'MP (' + s.MP.total + ')', 'LP (' + s.LP.total + ')'],
      datasets: [
        { label: '>= meta',      data: [s.HP.en_meta,  s.MP.en_meta,  s.LP.en_meta],  backgroundColor: '#0E7A4E', borderSkipped: false },
        { label: 'Parcial',      data: [s.HP.parcial,  s.MP.parcial,  s.LP.parcial],  backgroundColor: '#D4890A', borderSkipped: false },
        { label: 'Sin contacto', data: [s.HP.sin_cont, s.MP.sin_cont, s.LP.sin_cont], backgroundColor: '#B33A2E', borderSkipped: false }
      ]
    },
    options: {
      responsive: true, maintainAspectRatio: false, animation: { duration: 300 },
      plugins: { legend: { display: false }, tooltip: { mode: 'index', intersect: false, callbacks: { label: function(ctx) { return ' ' + ctx.dataset.label + ': ' + ctx.parsed.y + ' clientes'; }}}},
      scales: { x: Object.assign({ stacked: true }, xScale()), y: Object.assign({ stacked: true }, yScale()) }
    }
  });
}

function rebuildPeakChart(labels, vals) {
  var el = document.getElementById('peakChart');
  if (!el || typeof Chart === 'undefined') return;
  var existing = Chart.getChart(el);
  if (existing) existing.destroy();
  var maxV = Math.max.apply(null, vals.concat([1]));
  var colors = vals.map(function(v) {
    if (v === maxV) return 'rgba(13,45,107,0.95)';
    if (v >= maxV * 0.7) return 'rgba(27,75,155,0.85)';
    return 'rgba(90,147,204,0.65)';
  });
  var vLabels = { id: 'vLabels', afterDatasetsDraw: function(ch) {
    var ctx = ch.ctx;
    ch.data.datasets.forEach(function(ds, i) {
      ch.getDatasetMeta(i).data.forEach(function(bar, j) {
        ctx.save(); ctx.font = '600 13px IBM Plex Mono,monospace';
        ctx.fillStyle = '#0B1E3D'; ctx.textAlign = 'center'; ctx.textBaseline = 'bottom';
        ctx.fillText(ds.data[j], bar.x, bar.y - 4); ctx.restore();
      });
    });
  }};
  new Chart(el, {
    type: 'bar',
    data: { labels: labels, datasets: [{ data: vals, backgroundColor: colors, borderRadius: 5, borderSkipped: false }] },
    options: {
      responsive: true, maintainAspectRatio: false, animation: { duration: 300 }, layout: { padding: { top: 22 } },
      plugins: { legend: { display: false }, tooltip: { callbacks: { label: function(ctx) { return ' ' + ctx.parsed.y + ' reuniones'; }}}},
      scales: { x: xScale(), y: Object.assign(yScale(Math.ceil(maxV * 1.2)), { ticks: Object.assign({}, yScale().ticks, { callback: function(v) { return Number.isInteger(v) ? v : ''; }})})}
    },
    plugins: [vLabels]
  });
}

function rebuildReunYtdChart(byPrio, byPrioStats) {
  var el = document.getElementById('reunYtdChart');
  if (!el || typeof Chart === 'undefined') return;
  var existing = Chart.getChart(el);
  if (existing) existing.destroy();
  var total = (byPrio.HP || 0) + (byPrio.MP || 0) + (byPrio.LP || 0);
  var pctOf = function(v) { return total > 0 ? Math.round(v / total * 1000) / 10 : 0; };
  new Chart(el, {
    type: 'bar',
    data: {
      labels: [
        'HP (' + byPrioStats.HP.total + ' clientes)',
        'MP (' + byPrioStats.MP.total + ' clientes)',
        'LP (' + byPrioStats.LP.total + ' clientes)'
      ],
      datasets: [{ label: 'Reuniones YTD',
        data: [byPrio.HP || 0, byPrio.MP || 0, byPrio.LP || 0],
        backgroundColor: ['#1B4B9B','#2E6BAF','#5A93CC'], borderSkipped: false }]
    },
    options: {
      responsive: true, maintainAspectRatio: false, animation: { duration: 300 },
      plugins: { legend: { display: false }, tooltip: { callbacks: { label: function(ctx) {
        return ' ' + ctx.parsed.y + ' reuniones (' + pctOf(ctx.parsed.y) + '% del total)';
      }}}},
      scales: { x: xScale(), y: Object.assign(yScale(), { ticks: { callback: function(v) { return Number.isInteger(v) ? v : ''; }, color: TICK_C, font: TICK_FONT }})}
    }
  });
  var leg = document.getElementById('reun-ytd-leg');
  if (leg) {
    leg.innerHTML = [
      ['#1B4B9B', 'HP · ' + (byPrio.HP || 0) + ' reun. · ' + byPrioStats.HP.total + ' clientes'],
      ['#2E6BAF', 'MP · ' + (byPrio.MP || 0) + ' reun. · ' + byPrioStats.MP.total + ' clientes'],
      ['#5A93CC', 'LP · ' + (byPrio.LP || 0) + ' reun. · ' + byPrioStats.LP.total + ' clientes']
    ].map(function(it) { return '<span class="legend-item"><span class="legend-dot" style="background:' + it[0] + '"></span>' + it[1] + '</span>'; }).join('');
  }
}

// ── Setup ──────────────────────────────────────────────────────────────────
function _pyodideSetup() {
  document.body.classList.add('no-data');
  var input = document.getElementById('xlsm-input');
  if (input) input.addEventListener('change', function (e) {
    var f = e.target.files && e.target.files[0];
    if (f) processExcel(f);
  });
  var metasInput = document.getElementById('metas-input');
  if (metasInput) metasInput.addEventListener('change', function (e) {
    var f = e.target.files && e.target.files[0];
    if (f) processMetasResults(f);
  });
  pyodideReady = bootPyodide();
}
// Este <script> está al final del <body>, así que DOMContentLoaded ya
// disparó. Corremos al toque, pero por seguridad chequeamos readyState.
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', _pyodideSetup);
} else {
  _pyodideSetup();
}
"""


# ════════════════════════════════════════════════════════════════════════════
# Transformaciones
# ════════════════════════════════════════════════════════════════════════════

def strip_var_assignment(html: str, var_name: str, new_value: str) -> str:
    """
    Encuentra `var <name> = ...;` con un valor que puede ser literal o
    multi-línea ({…} ó […]) y lo reemplaza por `var <name> = <new_value>;`.

    Maneja balanceo de llaves/corchetes y respeta strings.
    """
    pattern = r"var\s+" + re.escape(var_name) + r"\s*="
    m = re.search(pattern, html)
    if not m:
        print(f"  WARN: var {var_name} no encontrado")
        return html
    # Buscar el primer { ó [ a partir del =
    i = m.end()
    while i < len(html) and html[i] in " \t\r\n":
        i += 1
    if i >= len(html) or html[i] not in "{[":
        print(f"  WARN: var {var_name} no parece estructura — saltado")
        return html
    opener = html[i]
    closer = "}" if opener == "{" else "]"
    depth = 0
    in_str = None
    j = i
    while j < len(html):
        ch = html[j]
        if in_str:
            if ch == "\\":
                j += 2
                continue
            if ch == in_str:
                in_str = None
        elif ch in ('"', "'"):
            in_str = ch
        elif ch == opener:
            depth += 1
        elif ch == closer:
            depth -= 1
            if depth == 0:
                end = j + 1
                if end < len(html) and html[end] == ";":
                    end += 1
                break
        j += 1
    else:
        print(f"  WARN: var {var_name} sin cierre — saltado")
        return html
    return html[: m.start()] + f"var {var_name} = {new_value};" + html[end:]


# Defaults vacíos por variable
EMPTY_DEFAULTS = {
    "data":              "{ 2024:[0,0,0,0,0,0,0,0,0,0,0,0], 2025:[0,0,0,0,0,0,0,0,0,0,0,0], 2026:[0,0,0,0,0,0,0,0,0,0,0,0] }",
    "prioData":          "{ HP:[], MP:[], LP:[] }",
    "clientsData":       "{ all:[], 2024:[], 2025:[], 2026:[] }",
    "recontactData":     "{ hp:{yes:[],no:[]}, mp:{yes:[],no:[]}, lp:{yes:[],no:[]} }",
    "activationData":    "{ nunca:{label:'',color:'#B33A2E',items:[]}, nuevos:{label:'',color:'#0E7A4E',items:[]}, reactiv:{label:'',color:'#0E7A4E',items:[]}, inactivos:{label:'',color:'#B33A2E',items:[]} }",
    "areasData":         "[]",
    "criticalList":      "[]",
    "riskAll":           "{ HP:[], MP:[], LP:[] }",
    "monthlyDetailData": "{}",
    "subtemaMoData":     "{}",
    "subteamAnnual":     "{}",
    "subtemasColors":    "{}",
    "subtemasClientsM":  "{}",
    "subtemasClientsA":  "{}",
    "riskData2026":      "{ HP:[], MP:[], LP:[] }",
}


def add_activation_attrs(html: str) -> str:
    """Inyecta data-attributes en las 4 KPI cards de Activación de cartera."""
    # Section label
    html = html.replace(
        '<div class="section-label">Activación de cartera Q1 2026 · click para ver detalle</div>',
        '<div class="section-label" data-q-text="activacion-section">Activación de cartera Q1 2026 · click para ver detalle</div>',
        1,
    )
    # Reuniones por área section label
    html = html.replace(
        '<div class="section-label">Reuniones Q1 2026 por área / producto · click para ver clientes</div>',
        '<div class="section-label" data-q-text="reuniones-area-section">Reuniones Q1 2026 por área / producto · click para ver clientes</div>',
        1,
    )

    # Las 4 KPI cards de activación (nunca / nuevos / reactiv / inactivos)
    # Cada una sigue el patrón: kpi-label, kpi-value, kpi-sub
    # Las identificamos por el onclick="toggleActivation('X')"
    for tipo in ['nunca', 'nuevos', 'reactiv', 'inactivos']:
        # Anchor por el onclick para ubicar la card y agregar attrs
        anchor = f"toggleActivation('{tipo}')"
        idx = html.find(anchor)
        if idx < 0:
            print(f"  WARN: no encontré card activación {tipo}")
            continue
        # Buscar la card y agregar attrs a label/value/sub
        # Buscamos los próximos kpi-label, kpi-value, kpi-sub dentro de ~600 chars
        end = idx + 800
        chunk = html[idx:end]
        chunk = re.sub(
            r'(<div class="kpi-label")(>[^<]+</div>)',
            rf'\1 data-act-label="{tipo}"\2', chunk, count=1,
        )
        chunk = re.sub(
            r'(<div class="kpi-value(?:\s+\w+)?")((?:\s+style="[^"]*")?)(>)(\d+)',
            rf'\1\2 data-act-value="{tipo}"\3\4', chunk, count=1,
        )
        chunk = re.sub(
            r'(<div class="kpi-sub")([^>]*)(>[^<]+</div>)',
            rf'\1\2 data-act-sub="{tipo}"\3', chunk, count=1,
        )
        html = html[:idx] + chunk + html[end:]
    return html


def add_quarter_attrs(html: str) -> str:
    """
    Inyecta data-attributes en labels que deben rolar trimestralmente
    (Q1 → Q2 → Q3 → Q4) y anualmente (año actual). applyData() los
    reemplaza usando data.currentQ y data.currentYear.
    """
    # Header badge (ya tiene id="update-badge", no necesita attr extra)
    # "Visión general Q1 2026"
    html = html.replace(
        '<div class="section-label">Visión general Q1 2026</div>',
        '<div class="section-label" data-q-section="vision-general">Visión general Q1 2026</div>',
        1,
    )
    # "Reuniones Q1 2026" KPI label
    html = html.replace(
        '<div class="kpi-label">Reuniones Q1 2026</div>',
        '<div class="kpi-label" data-q-text="reuniones-q-label">Reuniones Q1 2026</div>',
        1,
    )
    # "vs 98 en Q1 2025 (+11%)" sub
    html = re.sub(
        r'<div class="kpi-sub">vs \d+ en Q1 2025 \([\+\-]?\d+%\)</div>',
        '<div class="kpi-sub" data-q-text="reuniones-q-vs">vs 98 en Q1 2025 (+11%)</div>',
        html, count=1,
    )
    # "Reuniones 2026 YTD"
    html = html.replace(
        '<div class="kpi-label">Reuniones 2026 YTD</div>',
        '<div class="kpi-label" data-q-text="reuniones-ytd-label">Reuniones 2026 YTD</div>',
        1,
    )
    # "Top clientes Q1 2026 por segmento"
    html = html.replace(
        '<div class="card-title">Top clientes Q1 2026 por segmento</div>',
        '<div class="card-title" data-q-text="top-clientes-q-card">Top clientes Q1 2026 por segmento</div>',
        1,
    )
    return html


def add_kpi_attrs(html: str) -> str:
    """
    Inyecta data-attributes en los elementos KPI hardcodeados para que
    applyData() pueda actualizarlos via querySelector.
    """
    # Q1 value (en tab-resumen)
    html = re.sub(
        r'(<div class="kpi-label">Reuniones Q1 2026</div>\s*<div class="kpi-value")>',
        r'\1 data-kpi="q1-value">', html, count=1
    )
    # YTD value + sub
    html = re.sub(
        r'(<div class="kpi-label">Reuniones 2026 YTD</div>\s*<div class="kpi-value")(\s+style="[^"]*")?>',
        r'\1\2 data-kpi="ytd-value">', html, count=1
    )
    html = re.sub(
        r'(<div class="kpi-sub pos">)(Q1: \d+ &nbsp;&middot;&nbsp; Abr: \d+(?: &nbsp;&middot;&nbsp; May: \d+)?)(</div>)',
        r'<div class="kpi-sub pos" data-kpi="ytd-sub">\2\3', html, count=1
    )

    # Recontact bars — texto "X de Y recontactados", barra (.progress-bar) y % en el span
    # El bloque está cerca de id="recontact-detail-hp/mp/lp"
    for pl in ['hp', 'mp', 'lp']:
        # Encontrar el ancla y la región anterior (~700 chars antes)
        idx = html.find(f'id="recontact-detail-{pl}"')
        if idx < 0:
            print(f"  WARN: recontact-detail-{pl} no encontrado")
            continue
        bs, be = max(0, idx - 800), idx + 100
        block = html[bs:be]
        # Texto "X de Y recontactados"
        new_block, n1 = re.subn(
            r'>(\d+ de \d+ recontactados)<',
            lambda mm: f' data-recontact-text="{pl}">{mm.group(1)}<',
            block, count=1
        )
        # Barra de progreso
        new_block, n2 = re.subn(
            r'(progress-bar[^"]*"[^>]*style="[^"]*width:)\d+(%[^"]*")',
            lambda mm: mm.group(1) + '0' + mm.group(2) + f' data-recontact-bar="{pl}"',
            new_block, count=1
        )
        # % numérico
        new_block, n3 = re.subn(
            r'(font-weight:600;color:#\w+;")(\s*>)(\d+)(%</span>)',
            lambda mm: mm.group(1) + f' data-recontact-pct="{pl}"' + mm.group(2) + mm.group(3) + mm.group(4),
            new_block, count=1
        )
        if not (n1 and n2 and n3):
            print(f"  WARN: recontact-{pl} parcial: text={n1} bar={n2} pct={n3}")
        html = html[:bs] + new_block + html[be:]

    # Tab Riesgo — los 4 KPIs (activo/pendiente/inactivo/total) y los títulos de columna
    # Las 4 KPI cards están dentro del primer kpi-row-4 de #tab-riesgo
    rt_start = html.find('id="tab-riesgo"')
    if rt_start >= 0:
        rt_end = html.find('</div>\n\n  <div class="three-col">', rt_start)
        if rt_end < 0:
            rt_end = rt_start + 4000
        block = html[rt_start:rt_end]
        # 4 kpi cards en orden: activo, pendiente, inactivo, total
        # Cada una tiene <div class="kpi-value..."> N </div><div class="kpi-sub..."> HP:x...
        # Reemplazo por posición ordinal usando un counter
        labels_order = ['activo', 'pendiente', 'inactivo', 'total']
        i_label = [0]
        def repl_val(mm):
            if i_label[0] >= len(labels_order):
                return mm.group(0)
            tag = labels_order[i_label[0]]
            return mm.group(1) + f' data-riesgo="{tag}-value"' + mm.group(2) + mm.group(3) + mm.group(4)
        block_new = re.sub(
            r'(<div class="kpi-value)((?:\s+[\w-]+)?(?:\s+style="[^"]*")?)(>)(\d+)',
            repl_val, block
        )
        i_label = [0]
        def repl_sub(mm):
            if i_label[0] >= len(labels_order):
                return mm.group(0)
            tag = labels_order[i_label[0]]
            i_label[0] += 1
            return mm.group(1) + f' data-riesgo="{tag}-sub"' + mm.group(2) + mm.group(3)
        block_new = re.sub(
            r'(<div class="kpi-sub)((?:\s+[\w-]+)?(?:\s+style="[^"]*")?)(>HP:)',
            repl_sub, block_new
        )
        # Títulos de columna: <span>HP — N clientes</span>
        for prio in ['HP', 'MP', 'LP']:
            block_new = re.sub(
                r'<span>' + prio + r'\s*&mdash;\s*\d+\s*clientes</span>',
                f'<span data-riesgo-col="{prio}">{prio} &mdash; 0 clientes</span>',
                block_new, count=1
            )
        html = html[:rt_start] + block_new + html[rt_end:]

    return html


def transform(html: str) -> str:
    # 1) Empty defaults para las variables sensibles
    for var, val in EMPTY_DEFAULTS.items():
        html = strip_var_assignment(html, var, val)

    # 1b) Guard en setClients para no crashear con lista vacía (corre al page load)
    html = html.replace(
        "var list=clientsData[year], max=list[0][1];",
        "var list=clientsData[year]; if (!list || !list.length) { "
        "document.getElementById('clients-container').innerHTML="
        "'<div style=\"padding:20px;text-align:center;color:var(--text3);font-size:11px;\">"
        "Sube el Excel para ver datos</div>'; return; } var max=list[0][1];",
        1,
    )

    # 1.E) Contraste en KPI cards — borde más fuerte + sombra suave + hover
    kpi_css_old = (
        ".kpi {\n"
        "  background:var(--card); border:1px solid var(--border);\n"
        "  border-radius:var(--r); padding:16px 18px;\n"
        "}"
    )
    kpi_css_new = (
        ".kpi {\n"
        "  background:var(--card); border:1px solid var(--border2);\n"
        "  border-radius:var(--r); padding:16px 18px;\n"
        "  box-shadow:0 1px 4px rgba(11,30,61,0.06);\n"
        "  transition:box-shadow .2s, transform .2s;\n"
        "}\n"
        ".kpi:hover { box-shadow:0 6px 18px rgba(11,30,61,0.1); transform:translateY(-1px); }"
    )
    if kpi_css_old in html:
        html = html.replace(kpi_css_old, kpi_css_new, 1)
    else:
        print("  WARN: no encontré el CSS de .kpi")

    # También fortalecer .card (cards más grandes)
    card_css_old = (
        ".card {\n"
        "  background:var(--card); border:1px solid var(--border);\n"
        "  border-radius:var(--rl); padding:18px 20px;"
    )
    card_css_new = (
        ".card {\n"
        "  background:var(--card); border:1px solid var(--border2);\n"
        "  border-radius:var(--rl); padding:18px 20px;\n"
        "  box-shadow:0 1px 4px rgba(11,30,61,0.05);"
    )
    if card_css_old in html:
        html = html.replace(card_css_old, card_css_new, 1)
    else:
        print("  WARN: no encontré el CSS de .card")

    # 1a) Heatmap con más contraste — escala sqrt y paleta blue→amber→red para
    # que los meses de alta concentración salten visualmente vs los de baja.
    heatmap_old_color = (
        "function getColor(v){if(v===0)return '#F0F3F8';var t=v/max;"
        "if(t>0.8)return '#051B4A';if(t>0.6)return '#0D2D6B';if(t>0.4)return '#1B4B9B';"
        "if(t>0.22)return '#3A7BC8';if(t>0.1)return '#7FB3E8';return '#C8DDF4';}"
    )
    heatmap_new_color = (
        "function getColor(v){if(v===0)return '#F5F7FA';var t=Math.sqrt(v/max);"
        "if(t>0.85)return '#B33A2E';"          # hot red
        "if(t>0.70)return '#D4890A';"          # amber
        "if(t>0.55)return '#0D2D6B';"          # dark navy
        "if(t>0.40)return '#1B4B9B';"          # Ameris blue
        "if(t>0.25)return '#3A7BC8';"          # medium blue
        "if(t>0.12)return '#7FB3E8';"          # light blue
        "return '#C8DDF4';}"                   # pale blue
    )
    if heatmap_old_color in html:
        html = html.replace(heatmap_old_color, heatmap_new_color, 1)
    else:
        print("  WARN: no encontré el getColor del heatmap")

    # Texto blanco sobre los colores intensos (navy / amber / red)
    heatmap_old_txt = "function txtCol(v){return v/max>0.35?'#fff':'#1B4B9B';}"
    heatmap_new_txt = "function txtCol(v){return Math.sqrt(v/max)>0.40?'#fff':'#1B4B9B';}"
    if heatmap_old_txt in html:
        html = html.replace(heatmap_old_txt, heatmap_new_txt, 1)
    else:
        print("  WARN: no encontré el txtCol del heatmap")

    # Leyenda: agregar los nuevos colores hot al final
    heatmap_old_legend = "['#F0F3F8','#C8DDF4','#7FB3E8','#3A7BC8','#1B4B9B','#0D2D6B','#051B4A']"
    heatmap_new_legend = "['#F5F7FA','#C8DDF4','#7FB3E8','#3A7BC8','#1B4B9B','#0D2D6B','#D4890A','#B33A2E']"
    if heatmap_old_legend in html:
        html = html.replace(heatmap_old_legend, heatmap_new_legend, 1)
    else:
        print("  WARN: no encontré la leyenda del heatmap")

    # 1.D) Las cards de COMPARATIVA Q YoY y CROSS-SELL RADAR se insertan
    # AL FINAL (después de las eliminaciones de LISTA CRÍTICA y REUNIONES
    # POR ÁREA), para que esos regex no se coman partes del card recién
    # insertado. Ver el bloque marcado "1.D INSERT" más abajo.

    # setYoy: aceptar formato "YYYY-YYYY" además del legacy "2425"/"2526"
    html = html.replace(
        "var years = pair==='2425'?[2024,2025]:[2025,2026];",
        "var years; if (pair.indexOf('-')>=0) { var p=pair.split('-'); years=[parseInt(p[0],10),parseInt(p[1],10)]; } else { years = pair==='2425'?[2024,2025]:[2025,2026]; }",
        1,
    )

    # 1.C) LABELS Q/AÑO RESTANTES (Consolidado, Semáforo, Footer, AUM, setYoy)
    # Card title de detalle nuevos
    html = html.replace(
        '<div class="card-title">Clientes nuevos incorporados en Q1 2026</div>',
        '<div class="card-title" data-q-text="nuevos-card-title">Clientes nuevos incorporados en Q1 2026</div>',
        1,
    )
    # Consolidado section-label
    html = html.replace(
        '<div class="section-label">Consolidado Q1 2026 — progreso del año por segmento</div>',
        '<div class="section-label" data-q-text="consolidado-section">Consolidado Q1 2026 — progreso del año por segmento</div>',
        1,
    )
    # Headers de la tabla Consolidado (4 trimestres) — data-attr para JS
    html = re.sub(
        r'(<th[^>]*colspan="2">)Q1 2026 ✓(</th>)',
        r'\1<span data-q-th="1">Q1 2026 ✓</span>\2', html, count=1,
    )
    html = re.sub(
        r'(<th[^>]*colspan="2">)Q2 2026 \(YTD\)(</th>)',
        r'\1<span data-q-th="2">Q2 2026 (YTD)</span>\2', html, count=1,
    )
    html = re.sub(
        r'(<th[^>]*colspan="2">)Q3 2026(</th>)',
        r'\1<span data-q-th="3">Q3 2026</span>\2', html, count=1,
    )
    html = re.sub(
        r'(<th[^>]*colspan="2">)Q4 2026(</th>)',
        r'\1<span data-q-th="4">Q4 2026</span>\2', html, count=1,
    )
    # Semáforo
    html = html.replace(
        '<div class="card-title">Semáforo de contacto 2026 YTD — por segmento</div>',
        '<div class="card-title" data-q-text="semaforo-title">Semáforo de contacto 2026 YTD — por segmento</div>',
        1,
    )
    # KPI "Q1 2026" estática (en una sub-card, line 549 aprox)
    html = re.sub(
        r'(<div class="kpi-label">)Q1 2026(</div>)',
        r'\1<span data-q-text="kpi-q-current">Q1 2026</span>\2', html, count=1,
    )
    # Footer
    html = re.sub(
        r'<span id="footer-date">[^<]+</span>',
        '<span id="footer-date" data-q-text="footer-date">Q1 2026 · 2024–2026 · Perú incluido</span>',
        html, count=1,
    )
    # setYoy buttons — wrap container con data-q-buttons para rebuild dinámico
    html = re.sub(
        r'(<div class="ftabs[^"]*"[^>]*>)\s*'
        r'(<button class="ftab active" onclick="setYoy\(\'2425\',this\)">2024 vs 2025</button>\s*'
        r'<button class="ftab" onclick="setYoy\(\'2526\',this\)">2025 vs 2026</button>)',
        lambda m: m.group(1).replace('>', ' data-q-buttons="yoy-pairs">', 1) + m.group(2),
        html, count=1,
    )
    # AUM chart quarters hardcoded en initAumCharts
    html = re.sub(
        r"var quarters=\['Q1 2026','Q2 2026','Q3 2026','Q4 2026'\];",
        "var quarters=(window.AUM_QUARTERS||['Q1 2026','Q2 2026','Q3 2026','Q4 2026']);",
        html, count=1,
    )

    # 1.B) NUEVOS DISTRIBUIDORES: data-attrs en KPI + envolver tabla detalle
    # KPI value (Nuevos distribuidores Q1 2026)
    html = re.sub(
        r'(<div class="kpi-label">Nuevos distribuidores Q1 2026</div>\s*<div class="kpi-value pos")(>)\d+',
        r'\1 data-q-text="nuevos-kpi-value"\2 -',
        html, count=1,
    )
    html = html.replace(
        '<div class="kpi-label">Nuevos distribuidores Q1 2026</div>',
        '<div class="kpi-label" data-q-text="nuevos-kpi-label">Nuevos distribuidores Q1 2026</div>',
        1,
    )
    html = html.replace(
        '<div class="kpi-sub">FYNSA AGF · XIM · Investera · Vizcaya · Meta anual: 8</div>',
        '<div class="kpi-sub" data-q-text="nuevos-kpi-sub">FYNSA AGF · XIM · Investera · Vizcaya · Meta anual: 8</div>',
        1,
    )
    # Section label + envolver el card de detalle con id para rebuild dinámico
    html = html.replace(
        '<div class="section-label">Nuevos distribuidores Q1 2026 — detalle</div>',
        '<div class="section-label" data-q-text="nuevos-section" id="nuevos-section-label">Nuevos distribuidores Q1 2026 — detalle</div>',
        1,
    )
    # Reemplazar la tbody hardcodeada con un placeholder (la JS lo llena)
    html = re.sub(
        r'<tbody>\s*<tr style="border-bottom:1px solid var\(--border\);">\s*<td[^>]*>FYNSA AGF</td>.*?</tbody>',
        '<tbody id="nuevos-tbody"></tbody>',
        html, count=1, flags=re.DOTALL,
    )
    # El pie del card también dice "Meta anual: 8" — hacer dinámico
    html = re.sub(
        r'<div style="margin-top:8px;font-size:10px;color:var\(--text3\);">\* Cliente nuevo = distribuidor que ha invertido en Ameris por primera vez en 2026 · Meta anual: 8 nuevos \(2 por trimestre\)</div>',
        '<div style="margin-top:8px;font-size:10px;color:var(--text3);" data-q-text="nuevos-footer">* Cliente nuevo = distribuidor que ha invertido en Ameris por primera vez en 2026 · Meta anual: 8 nuevos (2 por trimestre)</div>',
        html, count=1,
    )

    # 1.A) AUTOMATIZACIÓN MULTI-AÑO ──────────────────────────────────────
    # Marcadores en los contenedores de tabs de año para que applyData
    # los reconstruya dinámicamente desde data.years
    html = html.replace(
        '<div class="ftabs" style="margin:0;">\n        <button class="ftab active" data-yr="2026"',
        '<div class="ftabs" style="margin:0;" data-q-buttons="vista-year">\n        <button class="ftab active" data-yr="2026"',
        1,
    )
    html = html.replace(
        '<div class="ftabs" style="margin:0;">\n        <button class="ftab active" data-sub-yr="2026"',
        '<div class="ftabs" style="margin:0;" data-q-buttons="subtema-year">\n        <button class="ftab active" data-sub-yr="2026"',
        1,
    )

    # Extender botones de mes en Subtemas de 1-6 a 1-12
    sub_mo_old = (
        '<button class="ftab" data-sub-mo="6" onclick="setSubMo(\'6\',this)">Jun</button>\n'
        '      </div>'
    )
    sub_mo_new = (
        '<button class="ftab" data-sub-mo="6" onclick="setSubMo(\'6\',this)">Jun</button>'
        '<button class="ftab" data-sub-mo="7" onclick="setSubMo(\'7\',this)">Jul</button>'
        '<button class="ftab" data-sub-mo="8" onclick="setSubMo(\'8\',this)">Ago</button>'
        '<button class="ftab" data-sub-mo="9" onclick="setSubMo(\'9\',this)">Sep</button>'
        '<button class="ftab" data-sub-mo="10" onclick="setSubMo(\'10\',this)">Oct</button>'
        '<button class="ftab" data-sub-mo="11" onclick="setSubMo(\'11\',this)">Nov</button>'
        '<button class="ftab" data-sub-mo="12" onclick="setSubMo(\'12\',this)">Dic</button>\n'
        '      </div>'
    )
    if sub_mo_old in html:
        html = html.replace(sub_mo_old, sub_mo_new, 1)
    else:
        print("  WARN: no encontré los botones de mes 1-6 en subtemas")

    # En setSubYear, el loop va `mo<=6`. Subirlo a 12 para que los nuevos
    # botones se habiliten/deshabiliten correctamente.
    html = html.replace("for(var mo=1;mo<=6;mo++)", "for(var mo=1;mo<=12;mo++)", 1)

    # Heatmap: usar HEATMAP_YEARS dinámico (default a 3 últimos años)
    html = html.replace(
        "[2024,2025,2026].forEach(function(y){data[y].forEach(function(v){if(v>0)allVals.push(v);});});",
        "(window.HEATMAP_YEARS||[2024,2025,2026]).forEach(function(y){(data[y]||[]).forEach(function(v){if(v>0)allVals.push(v);});});",
        1,
    )
    html = html.replace(
        "[2024,2025,2026].forEach(function(y){\n    html+='<div class=\"hm-label\">'+y+'</div>';\n    data[y].forEach(function(v)",
        "(window.HEATMAP_YEARS||[2024,2025,2026]).forEach(function(y){\n    html+='<div class=\"hm-label\">'+y+'</div>';\n    (data[y]||[]).forEach(function(v)",
        1,
    )

    # Header brand y label de recontact: data-attrs para applyData
    html = html.replace(
        '<p>Dashboard de seguimiento · 2024 – 2026</p>',
        '<p data-q-text="header-years">Dashboard de seguimiento · 2024 – 2026</p>',
        1,
    )
    html = html.replace(
        '<div class="section-label">Tasa de recontacto 2025 → 2026 YTD · click para ver detalle</div>',
        '<div class="section-label" data-q-text="recontact-section">Tasa de recontacto 2025 → 2026 YTD · click para ver detalle</div>',
        1,
    )

    # 1.E) Vista mensual: agregar filtro de Tipo (Dist/Inst) + barra de búsqueda
    vm_filter_old = (
        '<div style="display:flex;gap:6px;align-items:center;flex-wrap:wrap;">\n'
        '      <span style="font-size:10px;font-weight:600;color:var(--text3);min-width:40px;">Prio</span>\n'
        '      <div class="ftabs" style="margin:0;">\n'
        "        <button class=\"ftab active\" data-mprio=\"ALL\" onclick=\"setMonthPrio('ALL',this)\">Todos</button>"
    )
    vm_filter_new = (
        '<div style="display:flex;gap:6px;align-items:center;flex-wrap:wrap;margin-bottom:8px;">\n'
        '      <span style="font-size:10px;font-weight:600;color:var(--text3);min-width:40px;">Tipo</span>\n'
        '      <div class="ftabs" style="margin:0;">\n'
        "        <button class=\"ftab active\" data-mtipo=\"ALL\" onclick=\"setMonthTipo('ALL',this)\">Todos</button>"
        "<button class=\"ftab\" data-mtipo=\"Dist\" onclick=\"setMonthTipo('Dist',this)\">Distribuidor</button>"
        "<button class=\"ftab\" data-mtipo=\"Inst\" onclick=\"setMonthTipo('Inst',this)\">Institucional</button>\n"
        '      </div>\n'
        '    </div>\n'
        '    <div style="display:flex;gap:6px;align-items:center;flex-wrap:wrap;margin-bottom:8px;">\n'
        '      <span style="font-size:10px;font-weight:600;color:var(--text3);min-width:40px;">Buscar</span>\n'
        "      <input type=\"text\" id=\"monthly-search\" placeholder=\"Nombre del cliente...\" oninput=\"setMonthSearch(this.value)\""
        " style=\"flex:1;min-width:200px;max-width:400px;padding:6px 10px;font-size:12px;border:1px solid var(--border);border-radius:6px;background:var(--card);color:var(--text);outline:none;font-family:inherit;\">\n"
        '    </div>\n'
        '    <div style="display:flex;gap:6px;align-items:center;flex-wrap:wrap;">\n'
        '      <span style="font-size:10px;font-weight:600;color:var(--text3);min-width:40px;">Prio</span>\n'
        '      <div class="ftabs" style="margin:0;">\n'
        "        <button class=\"ftab active\" data-mprio=\"ALL\" onclick=\"setMonthPrio('ALL',this)\">Todos</button>"
    )
    if vm_filter_old in html:
        html = html.replace(vm_filter_old, vm_filter_new, 1)
    else:
        print("  WARN: no encontré el filtro de Prio en Vista mensual")

    # Extender el estado _MONTH con tipo y search
    html = html.replace(
        "window._MONTH = window._MONTH || {yr:'2026',mo:null,prio:'ALL'};",
        "window._MONTH = window._MONTH || {yr:'2026',mo:null,prio:'ALL',tipo:'ALL',search:''};",
        1,
    )

    # Agregar setMonthTipo + setMonthSearch como funciones nuevas, justo antes de renderMonthly
    new_setters = (
        "function setMonthTipo(tipo,btn){\n"
        "  document.querySelectorAll('[data-mtipo]').forEach(function(t){t.classList.remove('active');});btn.classList.add('active');\n"
        "  window._MONTH.tipo=tipo; _autoExpandVistaM(); renderMonthly();\n"
        "}\n"
        "function setMonthSearch(val){\n"
        "  window._MONTH.search=(val||'').toLowerCase().trim();\n"
        "  if (window._MONTH.search) _autoExpandVistaM();\n"
        "  renderMonthly();\n"
        "}\n"
        "function _autoExpandVistaM(){\n"
        "  var t=document.getElementById('monthly-table'),b=document.getElementById('vista-m-btn');\n"
        "  if(t&&b&&t.style.display==='none'){t.style.display='';b.innerHTML='&#9650; colapsar';}\n"
        "}\n"
        "function renderMonthly()"
    )
    html = html.replace("function renderMonthly()", new_setters, 1)

    # Extender el filtro de rows en renderMonthly para incluir tipo y search
    html = html.replace(
        "rows=rows.filter(function(r){return st.prio==='ALL'||r.prio===st.prio;});",
        "rows=rows.filter(function(r){"
        "if(st.prio!=='ALL'&&r.prio!==st.prio)return false;"
        "if(st.tipo&&st.tipo!=='ALL'&&r.tipo!==st.tipo)return false;"
        "if(st.search&&r.nombre.toLowerCase().indexOf(st.search)<0)return false;"
        "return true;});",
        1,
    )

    # 1a0) Hacer toda la fila del header clickeable para expandir Vista mensual
    # y Reuniones por subtema (más intuitivo que el botón chico a la derecha).
    vm_old = (
        '<div style="display:flex;justify-content:space-between;align-items:center;'
        'padding:4px 0 10px;border-bottom:1px solid var(--border);margin-bottom:8px;">\n'
        '      <div id="monthly-summary"'
    )
    vm_new = (
        '<div onclick="toggleVistaM(event)" '
        'style="display:flex;justify-content:space-between;align-items:center;'
        'padding:4px 0 10px;border-bottom:1px solid var(--border);margin-bottom:8px;'
        'cursor:pointer;user-select:none;">\n'
        '      <div id="monthly-summary"'
    )
    if vm_old in html:
        html = html.replace(vm_old, vm_new, 1)
    else:
        print("  WARN: no encontré header de Vista mensual")

    sub_hdr_old_clickable = (
        '<div style="display:flex;justify-content:space-between;align-items:center;'
        'padding:4px 0 10px;border-bottom:1px solid var(--border);margin-bottom:8px;">\n'
        '      <div style="display:flex;align-items:center;gap:8px;">\n'
        '        <span style="font-size:12px;font-weight:600;color:var(--text);">Reuniones por subtema</span>'
    )
    sub_hdr_new_clickable = (
        '<div onclick="toggleSubtemaChart(event)" '
        'style="display:flex;justify-content:space-between;align-items:center;'
        'padding:4px 0 10px;border-bottom:1px solid var(--border);margin-bottom:8px;'
        'cursor:pointer;user-select:none;">\n'
        '      <div style="display:flex;align-items:center;gap:8px;">\n'
        '        <span style="font-size:12px;font-weight:600;color:var(--text);">Reuniones por subtema</span>'
    )
    if sub_hdr_old_clickable in html:
        html = html.replace(sub_hdr_old_clickable, sub_hdr_new_clickable, 1)
    else:
        print("  WARN: no encontré header de Reuniones por subtema")

    # 1a1) Eliminar la sección "Reuniones por área / producto" del tab Clientes
    # (la info ya está en el tab Reuniones / Subtemas — evita duplicación).
    new_html, n_areas = re.subn(
        r"<!-- ── REUNIONES POR ÁREA.*?<div id=\"areas-container\"[^>]*></div>\s*</div>",
        "",
        html,
        count=1,
        flags=re.DOTALL,
    )
    if n_areas == 0:
        print("  WARN: no se eliminó la sección Reuniones por área")
    else:
        html = new_html

    # 1a2) Subtemas — hacer el conteo de reuniones por cliente más visible
    # (agrega encabezado "REUN." al header del expand y aumenta tamaño del número)
    sub_hdr_old = (
        "h+='<div style=\"font-size:10px;font-weight:600;color:'+col+';"
        "text-transform:uppercase;letter-spacing:.06em;margin-bottom:8px;\">'+clients.length+' clientes</div>';"
    )
    sub_hdr_new = (
        "h+='<div style=\"font-size:10px;font-weight:600;color:'+col+';"
        "text-transform:uppercase;letter-spacing:.06em;margin-bottom:8px;"
        "display:flex;justify-content:space-between;align-items:center;\">"
        "<span>'+clients.length+' clientes</span>"
        "<span style=\"color:var(--text3);font-size:9px;\">Cantidad de reuniones</span></div>';"
    )
    if sub_hdr_old in html:
        html = html.replace(sub_hdr_old, sub_hdr_new, 1)
    else:
        print("  WARN: no encontré el header del subtema expand")

    sub_count_old = (
        "<span style=\"font-family:monospace;font-size:11px;font-weight:600;color:var(--text);\">'+cl.reun+'</span>"
    )
    sub_count_new = (
        "<span style=\"font-family:monospace;font-size:13px;font-weight:700;"
        "color:'+col+';min-width:28px;text-align:right;\">'+cl.reun+'</span>"
    )
    if sub_count_old in html:
        html = html.replace(sub_count_old, sub_count_new, 1)
    else:
        print("  WARN: no encontré el count de cl.reun en subtemas")

    # 1b1) Bug en renderMonthly al agregar por "todo el año": no copia `tipo` al
    # map de agregación, así que las filas resultantes tienen tipo=undefined y
    # todas caen al fallback "Distribuidor". Inyectamos `tipo:r.tipo` al objeto.
    monthly_agg_old = "if(!map[r.nombre])map[r.nombre]={nombre:r.nombre,prio:r.prio,reun:0,subtemas:[]};"
    monthly_agg_new = "if(!map[r.nombre])map[r.nombre]={nombre:r.nombre,prio:r.prio,tipo:r.tipo,reun:0,subtemas:[]};"
    if monthly_agg_old in html:
        html = html.replace(monthly_agg_old, monthly_agg_new, 1)
    else:
        print("  WARN: no encontré el anchor del agregador de renderMonthly")

    # 1b2) Inyectar etiqueta Dist/Inst en cada fila de la vista mensual.
    # renderMonthly construye filas con statements separados:
    #   h+='<span ...>'+r.prio+'</span>';
    #   h+='<div style="flex:1;...">...nombre...';
    # Insertamos un nuevo statement entre ambos para el tag tipo.
    pill_old = "(r.prio==='?'?'&mdash;':r.prio)+'</span>';\n  h+='<div style=\"flex:1;min-width:0;\">"
    pill_new = (
        "(r.prio==='?'?'&mdash;':r.prio)+'</span>';\n"
        "  h+='<span style=\"font-size:9px;font-weight:700;padding:2px 8px;border-radius:3px;letter-spacing:.05em;"
        "background:'+(r.tipo===\"Inst\"?\"#6B4C9B\":\"#E8ECF4\")"
        "+';color:'+(r.tipo===\"Inst\"?\"#FFFFFF\":\"#3D5278\")"
        "+';text-align:center;flex-shrink:0;margin-top:1px;text-transform:uppercase;white-space:nowrap;\">'"
        "+(r.tipo===\"Inst\"?'Institucional':'Distribuidor')+'</span>';\n"
        "  h+='<div style=\"flex:1;min-width:0;\">"
    )
    if pill_old in html:
        html = html.replace(pill_old, pill_new, 1)
    else:
        print("  WARN: no encontré el anchor para inyectar la etiqueta Dist/Inst")

    # 1b3) Eliminar la sección "Lista crítica" del tab Clientes.
    # Estructura en la fuente:
    #   <!-- ── LISTA CRÍTICA ... -->
    #   <div class="section-label">...</div>
    #   <div class="card"> ... <div id="critical-list" ...></div>
    #   (comentarios huérfanos de gastos)
    #   </div>   ← cierra card
    #   </div>   ← cierra panel (este lo conservamos)
    # Regex: del comentario hasta la primera pareja </div>...</div>, reemplazar
    # por solo el segundo </div> (panel-close).
    new_html, n = re.subn(
        r"<!-- ── LISTA CRÍTICA.*?</div>\s*</div>",
        "</div>",
        html,
        count=1,
        flags=re.DOTALL,
    )
    if n == 0:
        print("  WARN: no se eliminó la lista crítica (regex no matcheó)")
    else:
        html = new_html

    # 1c) IIFE de critical-list — renombrar a función para poder llamarla post-upload
    html = html.replace(
        "(function(){\n  var el=document.getElementById('critical-list'); if(!el) return;",
        "function _renderCriticalList(){\n  var el=document.getElementById('critical-list'); if(!el) return;\n"
        "  if(!criticalList||!criticalList.length){el.innerHTML='<div style=\"padding:20px;text-align:center;color:var(--text3);font-size:11px;\">Sin clientes críticos</div>';return;}",
        1,
    )
    # Cerramos la función en el `})();` que sigue (es el primer `})();` después del nombre)
    # Hacemos un replace puntual del primer match después de la palabra clave del crítico:
    crit_anchor = "function _renderCriticalList(){"
    idx = html.find(crit_anchor)
    if idx >= 0:
        close = html.find("})();", idx)
        if close >= 0:
            html = html[: close] + "} _renderCriticalList();" + html[close + len("})();") :]

    # ─── 1.D INSERT: Cards Comparativa Q YoY + Cross-sell Radar ───────
    # Se hace AQUÍ (después de las eliminaciones de las secciones viejas
    # como LISTA CRÍTICA y REUNIONES POR ÁREA) para evitar que sus regex
    # de eliminación se traguen partes de estos cards nuevos.
    comparativa_card = (
        '\n  <!-- ── COMPARATIVA Q YoY POR PRIORIDAD (auto) ───── -->\n'
        '  <div class="section-label" data-q-text="comparativa-section">Comparativa Q1 2026 vs Q1 2025 — por prioridad</div>\n'
        '  <div class="card" style="margin-bottom:1.25rem;">\n'
        '    <div class="card-title">Reuniones del trimestre completo · año actual vs año anterior</div>\n'
        '    <div class="card-sub">Mismo período (mismos meses) comparado año contra año, separado por segmento de prioridad</div>\n'
        '    <div id="comparativa-q-yoy" style="margin-top:14px;display:grid;grid-template-columns:1fr 1fr 1fr;gap:14px;"></div>\n'
        '  </div>\n\n'
    )
    html = html.replace(
        "</div>\n\n<!-- ═══════════════════════════════════════════════ TAB 2: REUNIONES -->",
        comparativa_card + "</div>\n\n<!-- ═══════════════════════════════════════════════ TAB 2: REUNIONES -->",
        1,
    )

    crosssell_card = (
        '\n  <!-- ── CROSS-SELL RADAR (auto) ───── -->\n'
        '  <div class="section-label" data-q-text="crosssell-section">Oportunidades de cross-sell — alto contacto, pocos productos</div>\n'
        '  <div class="card" style="margin-bottom:1.25rem;">\n'
        '    <div class="card-title" data-q-text="crosssell-title">Clientes con ≥ 3 reuniones en 2026 y ≤ 1 producto distinto discutido (fuera de Follow Up)</div>\n'
        '    <div class="card-sub">Posibles candidatos para ofertas cruzadas — la relación está, falta diversificar el portfolio</div>\n'
        '    <div id="cross-sell-list" style="margin-top:10px;"></div>\n'
        '  </div>\n\n'
    )
    html = html.replace(
        "</div>\n\n<!-- ═══════════════════════════════════════════════ TAB 4: METAS AUM -->",
        crosssell_card + "</div>\n\n<!-- ═══════════════════════════════════════════════ TAB 4: METAS AUM -->",
        1,
    )

    # 1.D2) CRUCES POR DISTRIBUIDOR (saldo vs reuniones · gasto vs reuniones)
    # Se insertan al final del tab Clientes, después del cross-sell (mismo anchor
    # TAB 4, que sigue presente tras la inserción anterior).
    cross_dist_cards = (
        '\n  <!-- ── SALDO vs REUNIONES POR DISTRIBUIDOR (auto) ───── -->\n'
        '  <div class="section-label">Saldo vs reuniones — por distribuidor</div>\n'
        '  <div class="card" style="margin-bottom:1.25rem;">\n'
        '    <div class="card-title">MarketValue total vs reuniones · distribuidores en saldos por corte y con reuniones</div>\n'
        '    <div class="card-sub">Saldo = MarketValue total (todos los fondos, USD) al corte más reciente. Sólo distribuidores presentes en la hoja Saldos por corte <strong>y</strong> con reuniones. Requiere cargar el Excel Follow Up <strong>y</strong> el archivo de resultados de metas.</div>\n'
        '    <div id="cross-saldo-reun" style="margin-top:10px;"></div>\n'
        '  </div>\n\n'
        '  <!-- ── GASTO vs REUNIONES POR DISTRIBUIDOR (auto) ───── -->\n'
        '  <div class="section-label">Gasto vs reuniones — por distribuidor</div>\n'
        '  <div class="card" style="margin-bottom:1.25rem;">\n'
        '    <div class="card-title">Gasto comercial vs reuniones por distribuidor</div>\n'
        '    <div class="card-sub">Gasto acumulado (CLP) y reuniones del año, con costo por reunión. Distribuidores presentes en la hoja de gastos y en el foto. 0 reuniones = gasto sin reunión registrada.</div>\n'
        '    <div id="cross-gasto-reun" style="margin-top:10px;"></div>\n'
        '  </div>\n\n'
    )
    html = html.replace(
        "</div>\n\n<!-- ═══════════════════════════════════════════════ TAB 4: METAS AUM -->",
        cross_dist_cards + "</div>\n\n<!-- ═══════════════════════════════════════════════ TAB 4: METAS AUM -->",
        1,
    )

    # 1.F) AUM: conectar la UI existente (cards trimestrales + tabla YTD) a los
    # logros reales del archivo de resultados. Añadimos IDs a las celdas "Logro
    # YTD" de la tabla anual y al KPI "Logro YTD" para que el JS las actualice.
    # KPI "Logro YTD" (sub-panel anual)
    html = html.replace(
        '<div class="kpi-label">Logro YTD</div>\n'
        '        <div class="kpi-value" style="font-size:18px;color:var(--text3);">— pendiente</div>',
        '<div class="kpi-label">Logro YTD</div>\n'
        '        <div class="kpi-value" id="ytd-kpi-logro" style="font-size:18px;color:var(--text3);">— pendiente</div>',
        1,
    )
    # Celdas "Logro YTD" de la tabla anual: la última <td> de cada fila de
    # instrumento + la del Total. Las marcamos por orden con un contador.
    _ytd_ids = ['ytd-deuda-logro', 'ytd-inm-logro', 'ytd-inter-logro', 'ytd-notas-logro', 'ytd-total-logro']
    _ytd_i = [0]
    def _mark_ytd_cell(m):
        if _ytd_i[0] >= len(_ytd_ids):
            return m.group(0)
        cid = _ytd_ids[_ytd_i[0]]
        _ytd_i[0] += 1
        return m.group(1) + f' id="{cid}"' + m.group(2)
    html = re.sub(
        r'(<td style="text-align:center;padding:8px 12px;color:var\(--text3\);font-style:italic;)(">pendiente</td>)',
        _mark_ytd_cell, html,
    )

    # Marcar el texto del banner de aviso para poder actualizarlo al cargar resultados
    html = html.replace(
        '<div style="font-size:11px;color:var(--text2);line-height:1.6;">Solo se muestran metas de <strong>Aumento de Patrimonio (AUM)</strong>.',
        '<div style="font-size:11px;color:var(--text2);line-height:1.6;" id="aum-aviso-text">Solo se muestran metas de <strong>Aumento de Patrimonio (AUM)</strong>.',
        1,
    )

    # 1.H) Cards: línea extra (variación vs corte anterior + brecha vs meta)
    for k in ["deuda", "inm", "inter"]:
        anchor = (
            '<div style="display:flex;justify-content:space-between;margin-top:5px;">\n'
            '          <span style="font-size:10px;color:var(--text3);">Logro: '
            '<span id="q-' + k + '-logro"'
        )
        extra = (
            '<div id="q-' + k + '-extra" style="font-size:10px;margin-top:4px;display:none;"></div>\n'
            '        ' + anchor
        )
        if anchor in html:
            html = html.replace(anchor, extra, 1)
        else:
            print(f"  WARN: no encontré la fila de logro de {k} para inyectar extra")

    # 1.G) Ponderaciones: IDs en los badges de peso + KPI Ponderación AUM
    html = html.replace(
        'color:#1B4B9B;padding:2px 6px;border-radius:4px;font-weight:600;">17.15%</span>',
        'color:#1B4B9B;padding:2px 6px;border-radius:4px;font-weight:600;" id="q-deuda-pond" title="Peso en el scorecard">17.15%</span>',
        1,
    )
    html = html.replace(
        'color:#2E6BAF;padding:2px 6px;border-radius:4px;font-weight:600;">3.15%</span>',
        'color:#2E6BAF;padding:2px 6px;border-radius:4px;font-weight:600;" id="q-inm-pond" title="Peso en el scorecard">3.15%</span>',
        1,
    )
    html = html.replace(
        'color:#5A93CC;padding:2px 6px;border-radius:4px;font-weight:600;">14.35%</span>',
        'color:#5A93CC;padding:2px 6px;border-radius:4px;font-weight:600;" id="q-inter-pond" title="Peso en el scorecard">14.35%</span>',
        1,
    )
    html = html.replace(
        'color:#7D93B5;padding:2px 6px;border-radius:4px;font-weight:600;">0.35%</span>',
        'color:#7D93B5;padding:2px 6px;border-radius:4px;font-weight:600;" id="q-notas-pond" title="Peso en el scorecard">0.35%</span>',
        1,
    )
    # KPI "Ponderación AUM" — id en valor y sub para mostrar cumplimiento ponderado
    html = html.replace(
        '<div class="kpi-label">Ponderación AUM</div>\n'
        '        <div class="kpi-value">34.5%</div>\n'
        '        <div class="kpi-sub">Del scorecard total distribución</div>',
        '<div class="kpi-label">Ponderación AUM</div>\n'
        '        <div class="kpi-value" id="aum-pond-value">34.5%</div>\n'
        '        <div class="kpi-sub" id="aum-cump-sub">Del scorecard total distribución</div>',
        1,
    )

    # Patch initAumCharts: si hay resultados cargados, construir charts en USD
    html = html.replace(
        "function initAumCharts() {\n  if (aumChartsInited) return;\n  aumChartsInited = true;\n",
        "function initAumCharts() {\n  if (aumChartsInited) return;\n  aumChartsInited = true;\n"
        "  if (window.__METAS_RESULTS__ && typeof _buildAumChartsUSD === 'function') { _buildAumChartsUSD(window.__METAS_RESULTS__); var _b=document.querySelector('[onclick*=\"selectQuarter\"][onclick*=\"1T26\"]'); if(_b) selectQuarter('1T26',_b); return; }\n",
        1,
    )

    # Patch selectQuarter: que actualice los logros trimestrales al cambiar de Q
    html = html.replace(
        "['deuda','inm','inter','notas'].forEach(function(f){document.getElementById('q-'+f+'-meta').textContent=m[f];});",
        "['deuda','inm','inter','notas'].forEach(function(f){document.getElementById('q-'+f+'-meta').textContent=m[f];});"
        " if (typeof _updateAumLogrosQ === 'function') _updateAumLogrosQ(q, d);",
        1,
    )

    # 2) IDs / data-attributes para KPI + Q-rolling + activación + áreas
    html = add_kpi_attrs(html)
    html = add_quarter_attrs(html)
    html = add_activation_attrs(html)

    # 2b) Refactor del IIFE de áreas para que sea llamable post-upload
    areas_iife_start = "(function(){\n  var el=document.getElementById('areas-container'); if(!el) return;"
    if areas_iife_start in html:
        # Renombrar el IIFE a una función nombrada
        html = html.replace(
            areas_iife_start,
            "function _renderAreas(){\n  var el=document.getElementById('areas-container'); if(!el) return;\n"
            "  if (!Array.isArray(areasData) || !areasData.length) { el.innerHTML='<div style=\"padding:20px;text-align:center;color:var(--text3);font-size:11px;\">Sin datos</div>'; return; }",
            1,
        )
        # Cerrar la función: buscar el primer })(); después del anchor
        idx = html.find("function _renderAreas(){")
        close = html.find("})();", idx)
        if close >= 0:
            html = html[:close] + "} _renderAreas();" + html[close + len("})();"):]
        else:
            print("  WARN: no encontré cierre del IIFE de áreas")
    else:
        print("  WARN: no encontré el IIFE de áreas")

    # 3) Inyectar Pyodide head bloque antes de </head>
    html = html.replace("</head>", PYODIDE_HEAD + "\n</head>", 1)

    # 4) Inyectar upload bar después de <body>
    # (el body tiene un <div id="data-banner"> seguido del header; el upload-bar va arriba de todo)
    html = html.replace("<body>", "<body>\n" + UPLOAD_BAR_HTML, 1)

    # 5) Inyectar applyData() + bootPyodide al final del último <script>
    # El último </script> antes de </body> es donde van los render functions
    last_script_close = html.rfind("</script>")
    if last_script_close < 0:
        sys.exit("No encuentro </script> de cierre")
    html = html[:last_script_close] + APPLY_DATA_JS + html[last_script_close:]

    # 6) Tab "Clientes foto" (réplica de la hoja del Follow Up)
    html = add_foto_tab(html)

    return html


# ════════════════════════════════════════════════════════════════════════════
# TAB "CLIENTES FOTO" — réplica de la hoja "Clientes foto YYYY" del Follow Up
# ════════════════════════════════════════════════════════════════════════════

FOTO_PANEL_HTML = r"""<!-- ═══════════════════════════════════════════════ TAB 4: CLIENTES FOTO -->
<div class="panel" id="tab-foto">

  <div class="section-label" data-foto-text="section">Clientes foto — consolidado del año</div>
  <div style="font-size:11px;color:var(--text2);margin-bottom:1.25rem;" data-foto-text="intro">
    Todas las contrapartes del foto con sus reuniones mes a mes, cortado en el mes en curso. Misma estructura que la hoja <em>Clientes foto</em> del Follow Up.
  </div>

  <div class="kpi-row kpi-row-4">
    <div class="kpi accent">
      <div class="kpi-label" data-foto-text="kpi-reun-label">Reuniones YTD</div>
      <div class="kpi-value" data-foto-kpi="reun">—</div>
      <div class="kpi-sub" data-foto-kpi="reun-sub">&nbsp;</div>
    </div>
    <div class="kpi">
      <div class="kpi-label">Contrapartes en el foto</div>
      <div class="kpi-value" data-foto-kpi="total">—</div>
      <div class="kpi-sub" data-foto-kpi="total-sub">&nbsp;</div>
    </div>
    <div class="kpi" style="background:#E0F4EC;border-color:#a8d5bc;">
      <div class="kpi-label" style="color:#0E7A4E;">Con reunión YTD</div>
      <div class="kpi-value" style="color:#0E7A4E;" data-foto-kpi="con">—</div>
      <div class="kpi-sub" style="color:#0E7A4E;" data-foto-kpi="con-sub">&nbsp;</div>
    </div>
    <div class="kpi" style="background:#FAE9E7;border-color:#e8b5af;">
      <div class="kpi-label" style="color:#8B2A1F;">Sin reunión YTD</div>
      <div class="kpi-value neg" data-foto-kpi="sin">—</div>
      <div class="kpi-sub neg" data-foto-kpi="sin-sub">&nbsp;</div>
    </div>
  </div>

  <div style="background:var(--card2);border-radius:var(--rl);padding:14px 16px;margin-bottom:1rem;">
    <div style="display:flex;gap:6px;align-items:center;flex-wrap:wrap;margin-bottom:8px;">
      <span style="font-size:10px;font-weight:600;color:var(--text3);min-width:40px;">Tipo</span>
      <div class="ftabs" style="margin:0;">
        <button class="ftab active" data-ftipo="ALL" onclick="setFotoTipo('ALL',this)">Todos</button><button class="ftab" data-ftipo="Dist" onclick="setFotoTipo('Dist',this)">Distribuidor</button><button class="ftab" data-ftipo="Inst" onclick="setFotoTipo('Inst',this)">Institucional</button>
      </div>
    </div>
    <div style="display:flex;gap:6px;align-items:center;flex-wrap:wrap;margin-bottom:8px;">
      <span style="font-size:10px;font-weight:600;color:var(--text3);min-width:40px;">Prio</span>
      <div class="ftabs" style="margin:0;">
        <button class="ftab active" data-fprio="ALL" onclick="setFotoPrio('ALL',this)">Todos</button><button class="ftab" data-fprio="HP" onclick="setFotoPrio('HP',this)">HP</button><button class="ftab" data-fprio="MP" onclick="setFotoPrio('MP',this)">MP</button><button class="ftab" data-fprio="LP" onclick="setFotoPrio('LP',this)">LP</button>
      </div>
    </div>
    <div style="display:flex;gap:6px;align-items:center;flex-wrap:wrap;margin-bottom:8px;">
      <span style="font-size:10px;font-weight:600;color:var(--text3);min-width:40px;">Estado</span>
      <div class="ftabs" style="margin:0;">
        <button class="ftab active" data-fest="ALL" onclick="setFotoEstado('ALL',this)">Todos</button><button class="ftab" data-fest="CON" onclick="setFotoEstado('CON',this)">Con reunión YTD</button><button class="ftab" data-fest="SIN" onclick="setFotoEstado('SIN',this)">Sin reunión YTD</button>
      </div>
    </div>
    <div style="display:flex;gap:6px;align-items:center;flex-wrap:wrap;">
      <span style="font-size:10px;font-weight:600;color:var(--text3);min-width:40px;">Buscar</span>
      <input type="text" id="foto-search" placeholder="Nombre de la contraparte..." oninput="setFotoSearch(this.value)" style="flex:1;min-width:200px;max-width:400px;padding:6px 10px;font-size:12px;border:1px solid var(--border);border-radius:6px;background:var(--card);color:var(--text);outline:none;font-family:inherit;">
      <button onclick="exportFotoCSV()" style="font-size:11px;font-weight:500;color:var(--a1);padding:6px 12px;border-radius:6px;background:var(--card);border:1px solid var(--a1);cursor:pointer;font-family:inherit;white-space:nowrap;">↓ Descargar CSV</button>
    </div>
  </div>

  <div class="card">
    <div style="display:flex;justify-content:space-between;align-items:center;padding:4px 0 10px;border-bottom:1px solid var(--border);margin-bottom:8px;gap:12px;flex-wrap:wrap;">
      <div id="foto-summary" style="font-size:11px;color:var(--text2);"></div>
      <div style="display:flex;align-items:center;gap:10px;font-size:10px;color:var(--text3);flex-wrap:wrap;">
        <span style="display:flex;align-items:center;gap:4px;"><span style="display:inline-block;width:14px;height:14px;border-radius:2px;background:#E2726A;"></span>0</span>
        <span style="display:flex;align-items:center;gap:4px;"><span style="display:inline-block;width:14px;height:14px;border-radius:2px;background:#FBD5A5;"></span>1</span>
        <span style="display:flex;align-items:center;gap:4px;"><span style="display:inline-block;width:14px;height:14px;border-radius:2px;background:#C9E7CF;"></span>2</span>
        <span style="display:flex;align-items:center;gap:4px;"><span style="display:inline-block;width:14px;height:14px;border-radius:2px;background:#A8CBEA;"></span>3+</span>
        <span style="color:var(--border2);">|</span>
        <span>Click en un encabezado para ordenar</span>
      </div>
    </div>
    <div id="foto-table" style="overflow-x:auto;"></div>
  </div>

  <div style="margin-top:1rem;font-size:10px;color:var(--text3);" data-foto-text="footnote">
    Fuente: hoja «Clientes foto» del Follow Up. Los meses son los del propio foto (no los apuntes), por lo que los totales cuadran exactamente con el Excel.
  </div>

</div>

"""

FOTO_JS = r"""// ── CLIENTES FOTO — consolidado YTD por contraparte ──────────────────────────
// Reproduce la hoja "Clientes foto YYYY" del Follow Up: una fila por contraparte
// con sus reuniones mes a mes hasta el mes en curso. Data: d.fotoYTD del pipeline.
var fotoYTD = { year:0, prevYear:0, month:0, monthNames:[], rows:[] };
// Orden por defecto: alfabético, igual que la hoja del Excel.
window._FOTO = { tipo:'ALL', prio:'ALL', est:'ALL', search:'', sort:'name', dir:1 };

var FOTO_MES_ABBR = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
var FOTO_PRIO_C   = { HP:'#1B4B9B', MP:'#2E6BAF', LP:'#5A93CC', '?':'#8FA3BE' };

function _fEsc(s){ return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }

function _fotoSetBtn(attr, val, btn){
  document.querySelectorAll('['+attr+']').forEach(function(t){ t.classList.remove('active'); });
  if (btn) btn.classList.add('active');
}
function setFotoTipo(v,btn){ _fotoSetBtn('data-ftipo',v,btn); window._FOTO.tipo=v; renderFoto(); }
function setFotoPrio(v,btn){ _fotoSetBtn('data-fprio',v,btn); window._FOTO.prio=v; renderFoto(); }
function setFotoEstado(v,btn){ _fotoSetBtn('data-fest',v,btn); window._FOTO.est=v; renderFoto(); }
function setFotoSearch(v){ window._FOTO.search=(v||'').toLowerCase().trim(); renderFoto(); }
function setFotoSort(key){
  var st=window._FOTO;
  if(st.sort===key){ st.dir=-st.dir; }
  else { st.sort=key; st.dir=(key==='name')?1:-1; }
  renderFoto();
}

// Filas visibles según los filtros activos (compartido por tabla y CSV).
function _fotoVisible(){
  var st=window._FOTO;
  var rows=(fotoYTD.rows||[]).filter(function(r){
    if(st.tipo!=='ALL' && r.tipo!==st.tipo) return false;
    if(st.prio!=='ALL' && r.prio!==st.prio) return false;
    if(st.est==='CON' && !r.ytd) return false;
    if(st.est==='SIN' && r.ytd) return false;
    if(st.search && r.name.toLowerCase().indexOf(st.search)<0) return false;
    return true;
  });
  var key=st.sort, dir=st.dir;
  var pOrd={HP:0,MP:1,LP:2,'?':3};
  rows.sort(function(a,b){
    var va,vb;
    if(key==='name'){ return dir*a.name.localeCompare(b.name,'es'); }
    else if(key==='prio'){ va=pOrd[a.prio]!=null?pOrd[a.prio]:3; vb=pOrd[b.prio]!=null?pOrd[b.prio]:3; }
    else if(key==='tipo'){ va=(a.tipo==='Inst'?1:0); vb=(b.tipo==='Inst'?1:0); }
    else if(key==='prev'){ va=a.prevYtd; vb=b.prevYtd; }
    else if(key==='delta'){ va=a.ytd-a.prevYtd; vb=b.ytd-b.prevYtd; }
    else if(key==='dias'){ va=(a.dias==null?99999:a.dias); vb=(b.dias==null?99999:b.dias); }
    else if(key.charAt(0)==='m'){ var i=parseInt(key.slice(1),10); va=a.m[i]||0; vb=b.m[i]||0; }
    else { va=a.ytd; vb=b.ytd; }
    // Desempate estable: YTD desc y luego nombre.
    return dir*(va-vb) || (b.ytd-a.ytd) || a.name.localeCompare(b.name,'es');
  });
  return rows;
}

// Misma semántica de color que el formato condicional del Excel:
// 0 reuniones = rojo, 1 = naranjo, 2 = verde, 3 o más = azul.
function _fotoColor(v){
  if(!v)      return ['#E2726A','#FFFFFF'];
  if(v===1)   return ['#FBD5A5','#7A4B00'];
  if(v===2)   return ['#C9E7CF','#1B5E33'];
  return             ['#A8CBEA','#0D2D6B'];
}
function _fotoCell(v){
  var c=_fotoColor(v);
  return '<td style="text-align:center;padding:5px 4px;background:'+c[0]+';color:'+c[1]+";font-family:'IBM Plex Mono',monospace;font-size:11px;font-weight:600;border:1px solid rgba(255,255,255,0.55);\">"+v+'</td>';
}

function _fotoTh(label, key, align, extra){
  var st=window._FOTO;
  var arrow = st.sort===key ? (st.dir>0?' ▲':' ▼') : '';
  var col = st.sort===key ? 'var(--a1)' : 'var(--text3)';
  return '<th onclick="setFotoSort(\''+key+'\')" style="position:sticky;top:0;z-index:2;background:var(--card);'+(extra||'')+'text-align:'+(align||'center')+';padding:7px 6px;font-size:9px;font-weight:600;color:'+col+';text-transform:uppercase;letter-spacing:.06em;cursor:pointer;user-select:none;white-space:nowrap;border-bottom:2px solid var(--border2);">'+label+arrow+'</th>';
}

function renderFoto(){
  var el=document.getElementById('foto-table'); if(!el) return;
  var es=document.getElementById('foto-summary');
  var nMonths=(fotoYTD.month||0);
  var rows=_fotoVisible();

  // Labels que rolan con el año / mes de corte.
  if(nMonths){
    var MES_FULL=['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
    var rango = nMonths===1 ? MES_FULL[0] : ('Enero – '+MES_FULL[nMonths-1]);
    var setT=function(k,v){ var e=document.querySelector('[data-foto-text="'+k+'"]'); if(e) e.innerHTML=v; };
    setT('section','Clientes foto '+fotoYTD.year+' — consolidado '+rango);
    setT('kpi-reun-label','Reuniones '+fotoYTD.year+' YTD');
    setT('intro','Todas las contrapartes del foto con sus reuniones mes a mes, de '+rango.toLowerCase()+' de '+fotoYTD.year+'. Misma estructura que la hoja <em>Clientes foto '+fotoYTD.year+'</em> del Follow Up.');
    var orfanas=(fotoYTD.prevYtdTotal||0)-(fotoYTD.prevYtdMatched||0);
    setT('footnote','Fuente: hoja «Clientes foto '+fotoYTD.year+'» del Follow Up. Los meses vienen del propio foto (no de los apuntes), por lo que los totales cuadran exactamente con el Excel. La columna '+fotoYTD.prevYear+' es el mismo tramo Ene–'+FOTO_MES_ABBR[nMonths-1]+' del año anterior.'
      + (orfanas>0 ? ' Ojo: sólo cubre contrapartes que siguen en el foto '+fotoYTD.year+' — quedan '+orfanas+' reuniones de '+fotoYTD.prevYear+' de clientes que ya salieron del foto, por eso la columna suma '+fotoYTD.prevYtdMatched+' y no '+fotoYTD.prevYtdTotal+'.' : ''));
  }

  // KPIs — siempre sobre el universo completo, no sobre el filtro.
  var all=fotoYTD.rows||[];
  var totReun=all.reduce(function(a,r){return a+r.ytd;},0);
  var totPrev=all.reduce(function(a,r){return a+r.prevYtd;},0);
  var con=all.filter(function(r){return r.ytd>0;}).length;
  var sin=all.length-con;
  var dist=all.filter(function(r){return r.tipo!=='Inst';}).length;
  var inst=all.length-dist;
  var setK=function(k,v){ var e=document.querySelector('[data-foto-kpi="'+k+'"]'); if(e) e.innerHTML=v; };
  setK('reun', totReun);
  // Comparación contra el YTD real del año anterior (incluye contrapartes que
  // ya salieron del foto); la columna por fila sólo cubre las que siguen.
  var prevReal=(fotoYTD.prevYtdTotal!=null?fotoYTD.prevYtdTotal:totPrev);
  var dTot=totReun-prevReal;
  setK('reun-sub', prevReal ? ('vs '+prevReal+' en '+fotoYTD.prevYear+' YTD ('+(dTot>=0?'+':'')+dTot+')') : '&nbsp;');
  setK('total', all.length);
  setK('total-sub', dist+' distribuidores &nbsp;·&nbsp; '+inst+' institucionales');
  setK('con', con);
  setK('con-sub', all.length ? Math.round(con/all.length*100)+'% de la cartera' : '&nbsp;');
  setK('sin', sin);
  setK('sin-sub', all.length ? Math.round(sin/all.length*100)+'% de la cartera' : '&nbsp;');

  var visReun=rows.reduce(function(a,r){return a+r.ytd;},0);
  if(es){
    es.innerHTML = '<strong>'+rows.length+' contrapartes</strong> &nbsp;·&nbsp; '+visReun+' reuniones'
      + (rows.length!==all.length ? ' &nbsp;·&nbsp; <span style="color:var(--text3);">de '+all.length+' en el foto</span>' : '');
  }

  if(!nMonths){ el.innerHTML='<div style="padding:20px;text-align:center;color:var(--text3);">Sube el Excel para ver el foto consolidado</div>'; return; }
  if(!rows.length){ el.innerHTML='<div style="padding:20px;text-align:center;color:var(--text3);">Ninguna contraparte cumple los filtros</div>'; return; }

  // Layout = el de la hoja: Cliente · Tipo · Prioridad · meses · Total.
  // Después de una separación vertical van las columnas de análisis.
  var stickyTd='position:sticky;left:0;z-index:1;background:var(--card);';
  var sepL='border-left:3px double var(--border2);';
  var mono="font-family:'IBM Plex Mono',monospace;";
  var h='<table style="width:100%;border-collapse:collapse;font-size:11px;min-width:'+(560+nMonths*54)+'px;">';
  h+='<thead><tr>';
  h+='<th onclick="setFotoSort(\'name\')" style="position:sticky;left:0;top:0;z-index:3;background:var(--card);text-align:left;padding:7px 8px;font-size:9px;font-weight:600;color:'+(window._FOTO.sort==='name'?'var(--a1)':'var(--text3)')+';text-transform:uppercase;letter-spacing:.06em;cursor:pointer;user-select:none;white-space:nowrap;border-bottom:2px solid var(--border2);">Cliente'+(window._FOTO.sort==='name'?(window._FOTO.dir>0?' ▲':' ▼'):'')+'</th>';
  h+=_fotoTh('Tipo','tipo');
  h+=_fotoTh('Prioridad','prio');
  for(var i=0;i<nMonths;i++){ h+=_fotoTh(fotoYTD.monthNames[i]||FOTO_MES_ABBR[i],'m'+i); }
  h+=_fotoTh('Total','ytd');
  h+=_fotoTh(fotoYTD.prevYear+' YTD','prev',null,sepL);
  h+=_fotoTh('Δ','delta');
  h+=_fotoTh('Últ. contacto','dias','right');
  h+='</tr></thead><tbody>';

  rows.forEach(function(r){
    var pc=FOTO_PRIO_C[r.prio]||'#8FA3BE';
    h+='<tr style="border-bottom:1px solid var(--border);">';
    h+='<td style="'+stickyTd+'padding:5px 8px;font-weight:500;color:var(--text);white-space:nowrap;max-width:260px;overflow:hidden;text-overflow:ellipsis;" title="'+_fEsc(r.name)+'">'+_fEsc(r.name)+'</td>';
    h+='<td style="text-align:center;padding:5px 6px;'+mono+'font-size:11px;font-weight:600;color:'+(r.tipo==='Inst'?'#6B4C9B':'#3D5278')+';" title="'+(r.tipo==='Inst'?'Institucional':'Distribuidor')+'">'+(r.tipo==='Inst'?'I':'D')+'</td>';
    h+='<td style="text-align:center;padding:5px 6px;"><span style="font-size:9px;font-weight:600;padding:2px 6px;border-radius:3px;background:'+pc+'18;color:'+pc+';">'+(r.prio==='?'?'—':r.prio)+'</span></td>';
    for(var i=0;i<nMonths;i++){ h+=_fotoCell(r.m[i]||0); }
    h+='<td style="text-align:center;padding:5px 8px;'+mono+'font-size:12px;font-weight:700;color:'+(r.ytd?'var(--text)':'#B33A2E')+';background:var(--card2);">'+r.ytd+'</td>';
    h+='<td style="text-align:center;padding:5px 8px;'+sepL+mono+'font-size:11px;color:var(--text3);">'+r.prevYtd+'</td>';
    var d=r.ytd-r.prevYtd;
    h+='<td style="text-align:center;padding:5px 8px;'+mono+'font-size:11px;font-weight:600;color:'+(d>0?'#0E7A4E':(d<0?'#B33A2E':'var(--text3)'))+';">'+(d>0?'+':'')+(d||'—')+'</td>';
    var lastTxt = r.last ? (r.last+' <span style="color:var(--text3);">('+r.dias+'d)</span>') : '<span style="color:var(--text3);">sin historial</span>';
    h+='<td style="text-align:right;padding:5px 8px;'+mono+'font-size:10px;color:var(--text2);white-space:nowrap;">'+lastTxt+'</td>';
    h+='</tr>';
  });

  // Totales de las filas visibles.
  var tm=[]; for(var i=0;i<nMonths;i++){ tm.push(rows.reduce(function(a,r){return a+(r.m[i]||0);},0)); }
  var tPrev=rows.reduce(function(a,r){return a+r.prevYtd;},0);
  var tD=visReun-tPrev;
  h+='</tbody><tfoot><tr style="border-top:2px solid var(--border2);font-weight:700;">';
  h+='<td style="'+stickyTd+'padding:7px 8px;font-size:10px;text-transform:uppercase;letter-spacing:.06em;color:var(--text2);">Total</td>';
  h+='<td colspan="2"></td>';
  tm.forEach(function(v){ h+='<td style="text-align:center;padding:7px 6px;'+mono+'font-size:11px;color:var(--a1);">'+(v||'—')+'</td>'; });
  h+='<td style="text-align:center;padding:7px 8px;'+mono+'font-size:12px;color:var(--a1);background:var(--card2);">'+visReun+'</td>';
  h+='<td style="text-align:center;padding:7px 8px;'+sepL+mono+'font-size:11px;color:var(--text3);">'+tPrev+'</td>';
  h+='<td style="text-align:center;padding:7px 8px;'+mono+'font-size:11px;color:'+(tD>0?'#0E7A4E':(tD<0?'#B33A2E':'var(--text3)'))+';">'+(tD>0?'+':'')+(tD||'—')+'</td>';
  h+='<td></td></tr></tfoot></table>';
  el.innerHTML=h;
}

function exportFotoCSV(){
  var rows=_fotoVisible(); var n=fotoYTD.month||0;
  if(!rows.length){ return; }
  // Separador ';' — es lo que espera Excel en configuración regional es-CL.
  var head=['Cliente','Tipo','Prioridad'].concat(fotoYTD.monthNames.slice(0,n))
    .concat(['Total '+fotoYTD.year,'YTD '+fotoYTD.prevYear,'Delta','Ultimo contacto','Dias']);
  var lines=[head.join(';')];
  rows.forEach(function(r){
    var cells=['"'+r.name.replace(/"/g,'""')+'"', r.tipo==='Inst'?'I':'D', r.prio==='?'?'':r.prio];
    for(var i=0;i<n;i++){ cells.push(r.m[i]||0); }
    cells.push(r.ytd, r.prevYtd, r.ytd-r.prevYtd, r.last||'', r.dias==null?'':r.dias);
    lines.push(cells.join(';'));
  });
  var blob=new Blob(['﻿'+lines.join('\r\n')],{type:'text/csv;charset=utf-8;'});
  var a=document.createElement('a');
  a.href=URL.createObjectURL(blob);
  a.download='clientes_foto_'+fotoYTD.year+'_YTD_'+FOTO_MES_ABBR[n-1]+'.csv';
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  setTimeout(function(){ URL.revokeObjectURL(a.href); }, 1000);
}

"""


def add_foto_tab(html: str) -> str:
    """Agrega el tab 'Clientes foto': botón, panel y lógica de render."""
    # 1) Botón en la tab-bar, antes de "Metas AUM"
    metas_btn = ('  <button class="tab-btn" onclick="showTab(\'metas-aum\',this)">'
                 '<span class="tab-icon">◆</span>Metas AUM</button>')
    foto_btn = ('  <button class="tab-btn" onclick="showTab(\'foto\',this)">'
                '<span class="tab-icon">▣</span>Clientes foto</button>\n')
    if metas_btn in html:
        html = html.replace(metas_btn, foto_btn + metas_btn, 1)
    else:
        print("  WARN: no encontré el botón 'Metas AUM' para insertar el tab Clientes foto")

    # 2) Con 6 tabs la barra puede no caber en pantallas angostas
    html = html.replace(
        "  display:flex; gap:0; padding: 0 2rem;\n",
        "  display:flex; gap:0; padding: 0 2rem; flex-wrap:wrap;\n",
        1,
    )

    # 3) Panel, antes del de Metas AUM (que pasa a ser el TAB 5)
    metas_panel = "<!-- ═══════════════════════════════════════════════ TAB 4: METAS AUM -->"
    if metas_panel in html:
        html = html.replace(
            metas_panel,
            FOTO_PANEL_HTML + "<!-- ═══════════════════════════════════════════════ TAB 5: METAS AUM -->",
            1,
        )
    else:
        print("  WARN: no encontré el panel 'TAB 4: METAS AUM' para insertar Clientes foto")

    # 4) JS de render, antes de renderRiesgo2026
    if "function renderRiesgo2026()" in html:
        html = html.replace("function renderRiesgo2026()", FOTO_JS + "function renderRiesgo2026()", 1)
    else:
        print("  WARN: no encontré renderRiesgo2026 para insertar el JS de Clientes foto")

    return html


def main():
    if not SOURCE_HTML.exists():
        sys.exit(f"No encuentro source HTML: {SOURCE_HTML}")
    print(f"Leyendo {SOURCE_HTML.name}...")
    html = SOURCE_HTML.read_text(encoding="utf-8")
    print(f"  {len(html)//1024} KB")
    html = transform(html)
    OUT_HTML.write_text(html, encoding="utf-8")
    print(f"  -> {OUT_HTML.relative_to(REPO_DIR)} ({len(html)//1024} KB)")
    print("Listo.")


if __name__ == "__main__":
    main()
