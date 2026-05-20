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
</style>
"""


UPLOAD_BAR_HTML = """
<!-- ── Upload bar (Pyodide-powered) ────────────────────────────────────── -->
<div id="upload-bar" class="empty">
  <input type="file" id="xlsm-input" accept=".xlsm,.xlsx" style="display:none;">
  <button id="upload-btn" disabled onclick="document.getElementById('xlsm-input').click();">
    <span id="upload-btn-icon">⏳</span> <span id="upload-btn-text">Cargando motor…</span>
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
    if (typeof renderRiesgo2026 === 'function')    renderRiesgo2026();
    if (typeof renderMonthly === 'function')       renderMonthly();
    if (typeof renderSubtemas === 'function')      renderSubtemas();
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
