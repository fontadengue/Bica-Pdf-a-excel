import re
import io
import webbrowser
import threading
from flask import Flask, request, send_file, jsonify
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)

DATE_RE = re.compile(r'^\d{2}/\d{2}/\d{4}$')
NUM_RE  = re.compile(r'^\(?[\d,]+\.\d{2}\)?$')

X_COMP_MIN  = 100
X_COMP_MAX  = 175
X_CONC_MIN  = 175
X_CONC_MAX  = 375
X_DEB_MIN   = 375
X_DEB_MAX   = 430
X_CRED_MIN  = 430
X_CRED_MAX  = 490
X_SALDO_MIN = 490

GARBAGE_MARKERS = [
    ' registrado en el Banco si,',
    ' Se presume conformidad',
    ' dentro de los sesenta',
    ' Banco Bica SA',
    'Inscripto - CUIT: 30-71233123',
    ' R Ct Fecha',
    ' en ctas.',
    ' Impuesto computa',
    '01.04.008',
]

def parse_num(s):
    if not s:
        return None
    neg = s.startswith('(') and s.endswith(')')
    s = s.strip('()').replace(',', '')
    try:
        v = float(s)
        return -v if neg else v
    except:
        return None

def clean_concept(text):
    if not text:
        return text
    for marker in GARBAGE_MARKERS:
        idx = text.find(marker)
        if idx != -1:
            text = text[:idx].strip()
    return text

def parse_pdf(file_bytes):
    rows = []
    current = None

    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        total = len(pdf.pages)
        for page in pdf.pages:
            words = page.extract_words(x_tolerance=3, y_tolerance=3)
            lines = {}
            for w in words:
                y = round(w['top'], 1)
                if y not in lines:
                    lines[y] = []
                lines[y].append(w)

            in_table = False
            for y in sorted(lines.keys()):
                lw = sorted(lines[y], key=lambda w: w['x0'])
                full = ' '.join(w['text'] for w in lw)

                if 'COMPROBANTE' in full and 'CONCEPTO' in full and 'DÉBITO' in full:
                    in_table = True
                    continue
                if not in_table:
                    continue

                fecha_ws = [w for w in lw if w['x0'] < X_COMP_MIN]
                comp_ws  = [w for w in lw if X_COMP_MIN <= w['x0'] < X_COMP_MAX]
                conc_ws  = [w for w in lw if X_CONC_MIN <= w['x0'] < X_CONC_MAX]
                deb_ws   = [w for w in lw if X_DEB_MIN  <= w['x0'] < X_DEB_MAX  and NUM_RE.match(w['text'])]
                cred_ws  = [w for w in lw if X_CRED_MIN <= w['x0'] < X_CRED_MAX and NUM_RE.match(w['text'])]
                saldo_ws = [w for w in lw if w['x0'] >= X_SALDO_MIN and NUM_RE.match(w['text'])]
                conc_text = ' '.join(w['text'] for w in conc_ws)

                if fecha_ws and DATE_RE.match(fecha_ws[0]['text']):
                    if current:
                        rows.append(current)
                    current = {
                        'fecha':    fecha_ws[0]['text'],
                        'comp':     ' '.join(w['text'] for w in comp_ws),
                        'concepto': conc_text,
                        'debito':   parse_num(deb_ws[0]['text'])  if deb_ws  else None,
                        'credito':  parse_num(cred_ws[0]['text']) if cred_ws else None,
                        'saldo':    parse_num(saldo_ws[0]['text']) if saldo_ws else None,
                    }
                elif current and conc_text and conc_text != 'TRANSPORTE':
                    current['concepto'] = (current['concepto'] + ' ' + conc_text).strip()

    if current:
        rows.append(current)

    for r in rows:
        r['concepto'] = clean_concept(r['concepto'])

    return rows

def build_excel(rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "Movimientos"

    HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
    HEADER_FONT = Font(bold=True, color="FFFFFF", name="Arial", size=10)
    ALT_FILL    = PatternFill("solid", fgColor="EBF3FB")
    NORM_FILL   = PatternFill("solid", fgColor="FFFFFF")
    DEB_FILL    = PatternFill("solid", fgColor="FFE0E0")
    CRED_FILL   = PatternFill("solid", fgColor="E0F0E0")
    SALDO_FILL  = PatternFill("solid", fgColor="FFF8DC")
    CTAS_FILL   = PatternFill("solid", fgColor="FFFACD")
    thin        = Side(style='thin', color='CCCCCC')
    border      = Border(left=thin, right=thin, top=thin, bottom=thin)

    headers    = ["Fecha", "Comprobante", "Concepto", "Débito", "Crédito", "Saldo", "Cuenta Contable"]
    col_widths = [12, 16, 80, 15, 15, 16, 25]

    for col, (h, w) in enumerate(zip(headers, col_widths), start=1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = border
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.row_dimensions[1].height = 20
    ws.freeze_panes = "A2"

    for i, row in enumerate(rows, start=2):
        fill = ALT_FILL if i % 2 == 0 else NORM_FILL
        data = [row['fecha'], row['comp'], row['concepto'],
                row['debito'], row['credito'], row['saldo'], '']

        for col, val in enumerate(data, start=1):
            cell = ws.cell(row=i, column=col, value=val)
            cell.font = Font(name="Arial", size=9)
            cell.border = border
            if col == 1:
                cell.alignment = Alignment(horizontal='center', vertical='top')
            elif col == 2:
                cell.alignment = Alignment(horizontal='center', vertical='top')
            elif col == 3:
                cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
                cell.fill = fill
            elif col in (4, 5, 6):
                cell.alignment = Alignment(horizontal='right', vertical='top')
                if val is not None:
                    cell.number_format = '#,##0.00'
                if col == 4 and val is not None:
                    cell.fill = DEB_FILL
                elif col == 5 and val is not None:
                    cell.fill = CRED_FILL
                elif col == 6:
                    cell.fill = SALDO_FILL
                    cell.font = Font(name="Arial", size=9, bold=True)
            elif col == 7:
                cell.fill = CTAS_FILL
                cell.alignment = Alignment(horizontal='left', vertical='top')

    ws.auto_filter.ref = f"A1:G{len(rows)+1}"

    ws2 = wb.create_sheet("Resumen")
    ws2['A1'] = "Resumen del Extracto"
    ws2['A1'].font = Font(bold=True, size=14, name="Arial", color="1F4E79")

    total_deb  = sum(r['debito']  for r in rows if r['debito']  is not None)
    total_cred = sum(r['credito'] for r in rows if r['credito'] is not None)
    saldo_ini  = rows[0]['saldo']  if rows else 0
    saldo_fin  = rows[-1]['saldo'] if rows else 0

    summary = [
        ("Total movimientos",       len(rows)),
        ("Movimientos débito",      sum(1 for r in rows if r['debito']  is not None)),
        ("Movimientos crédito",     sum(1 for r in rows if r['credito'] is not None)),
        ("", ""),
        ("Saldo inicial",           saldo_ini),
        ("Total débitos",           total_deb),
        ("Total créditos",          total_cred),
        ("Saldo final",             saldo_fin),
    ]
    for r_idx, (label, val) in enumerate(summary, start=3):
        c1 = ws2.cell(row=r_idx, column=1, value=label)
        c2 = ws2.cell(row=r_idx, column=2, value=val)
        c1.font = Font(name="Arial", size=10, bold=bool(label))
        c2.font = Font(name="Arial", size=10)
        if isinstance(val, float):
            c2.number_format = '#,##0.00'
            c2.alignment = Alignment(horizontal='right')
    ws2.column_dimensions['A'].width = 28
    ws2.column_dimensions['B'].width = 20

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

HTML = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>BICA PDF → Excel</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=Sora:wght@300;400;600&display=swap');

  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

  :root {
    --navy: #0f1c2e;
    --navy2: #1a2f48;
    --accent: #2a9d8f;
    --accent2: #e76f51;
    --light: #f0f4f8;
    --text: #1a2f48;
    --muted: #6b7e96;
    --border: #d0dbe8;
    --white: #ffffff;
    --green: #d1f0eb;
    --green-text: #0f7a6b;
    --red: #fde8e8;
    --red-text: #c0392b;
    --yellow: #fef9e7;
  }

  body {
    font-family: 'Sora', sans-serif;
    background: var(--light);
    color: var(--text);
    min-height: 100vh;
    display: flex;
    flex-direction: column;
    align-items: center;
    padding: 2rem 1rem 4rem;
  }

  header {
    width: 100%;
    max-width: 680px;
    margin-bottom: 2.5rem;
  }

  .logo {
    display: flex;
    align-items: center;
    gap: 12px;
    margin-bottom: 0.4rem;
  }

  .logo-icon {
    width: 40px; height: 40px;
    background: var(--navy);
    border-radius: 10px;
    display: flex; align-items: center; justify-content: center;
  }

  .logo-icon svg { width: 22px; height: 22px; fill: var(--accent); }

  h1 {
    font-size: 1.5rem;
    font-weight: 600;
    color: var(--navy);
    letter-spacing: -0.02em;
  }

  .subtitle {
    font-size: 0.85rem;
    color: var(--muted);
    font-weight: 300;
    margin-top: 2px;
  }

  .card {
    width: 100%;
    max-width: 680px;
    background: var(--white);
    border: 1px solid var(--border);
    border-radius: 16px;
    padding: 2rem;
  }

  .drop-zone {
    border: 2px dashed var(--border);
    border-radius: 12px;
    padding: 3rem 2rem;
    text-align: center;
    cursor: pointer;
    transition: all 0.2s ease;
    position: relative;
    background: var(--light);
  }

  .drop-zone:hover, .drop-zone.drag-over {
    border-color: var(--accent);
    background: #eaf7f5;
  }

  .drop-zone input[type="file"] {
    position: absolute; inset: 0;
    opacity: 0; cursor: pointer; width: 100%; height: 100%;
  }

  .drop-icon {
    width: 52px; height: 52px;
    margin: 0 auto 1rem;
    background: var(--navy);
    border-radius: 12px;
    display: flex; align-items: center; justify-content: center;
  }

  .drop-icon svg { width: 28px; height: 28px; fill: none; stroke: var(--accent); stroke-width: 2; stroke-linecap: round; stroke-linejoin: round; }

  .drop-label {
    font-size: 1rem;
    font-weight: 600;
    color: var(--navy);
    margin-bottom: 0.3rem;
  }

  .drop-hint {
    font-size: 0.8rem;
    color: var(--muted);
    font-family: 'DM Mono', monospace;
  }

  .file-info {
    display: none;
    margin-top: 1.2rem;
    padding: 0.75rem 1rem;
    background: var(--light);
    border-radius: 8px;
    border: 1px solid var(--border);
    display: none;
    align-items: center;
    gap: 10px;
    font-size: 0.85rem;
  }

  .file-info.visible { display: flex; }

  .file-name {
    font-family: 'DM Mono', monospace;
    font-weight: 500;
    color: var(--navy);
    flex: 1;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }

  .file-size {
    color: var(--muted);
    font-size: 0.78rem;
  }

  .btn {
    display: block;
    width: 100%;
    margin-top: 1.5rem;
    padding: 0.9rem;
    background: var(--navy);
    color: var(--white);
    border: none;
    border-radius: 10px;
    font-family: 'Sora', sans-serif;
    font-size: 0.95rem;
    font-weight: 600;
    cursor: pointer;
    letter-spacing: 0.01em;
    transition: background 0.15s ease, transform 0.1s ease;
  }

  .btn:hover { background: var(--navy2); }
  .btn:active { transform: scale(0.99); }
  .btn:disabled { background: var(--border); color: var(--muted); cursor: not-allowed; transform: none; }

  .progress-wrap {
    display: none;
    margin-top: 1.5rem;
  }

  .progress-wrap.visible { display: block; }

  .progress-label {
    font-size: 0.82rem;
    color: var(--muted);
    margin-bottom: 0.5rem;
    font-family: 'DM Mono', monospace;
    display: flex;
    justify-content: space-between;
  }

  .progress-bar-bg {
    height: 6px;
    background: var(--light);
    border-radius: 99px;
    overflow: hidden;
    border: 1px solid var(--border);
  }

  .progress-bar-fill {
    height: 100%;
    background: linear-gradient(90deg, var(--accent), #48cfc0);
    border-radius: 99px;
    width: 0%;
    transition: width 0.3s ease;
  }

  .result {
    display: none;
    margin-top: 1.5rem;
    border-radius: 10px;
    overflow: hidden;
    border: 1px solid var(--border);
  }

  .result.visible { display: block; }

  .result-header {
    background: var(--navy);
    padding: 0.9rem 1.2rem;
    display: flex;
    align-items: center;
    gap: 10px;
  }

  .result-header svg { width: 18px; height: 18px; fill: var(--accent); flex-shrink: 0; }

  .result-title {
    font-size: 0.9rem;
    font-weight: 600;
    color: var(--white);
    flex: 1;
  }

  .stats {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 0;
    border-top: 1px solid var(--border);
  }

  .stat {
    padding: 1rem 1.2rem;
    border-right: 1px solid var(--border);
  }
  .stat:last-child { border-right: none; }

  .stat-label {
    font-size: 0.72rem;
    color: var(--muted);
    font-family: 'DM Mono', monospace;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    margin-bottom: 4px;
  }

  .stat-value {
    font-size: 1.1rem;
    font-weight: 600;
    color: var(--navy);
    font-family: 'DM Mono', monospace;
  }

  .stat-value.green { color: var(--green-text); }
  .stat-value.red { color: var(--red-text); }

  .download-btn {
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 8px;
    width: calc(100% - 2.4rem);
    margin: 1rem 1.2rem 1.2rem;
    padding: 0.75rem;
    background: var(--accent);
    color: var(--white);
    border: none;
    border-radius: 8px;
    font-family: 'Sora', sans-serif;
    font-size: 0.9rem;
    font-weight: 600;
    cursor: pointer;
    text-decoration: none;
    transition: background 0.15s ease;
  }
  .download-btn:hover { background: #238f82; }
  .download-btn svg { width: 16px; height: 16px; fill: none; stroke: currentColor; stroke-width: 2; stroke-linecap: round; stroke-linejoin: round; }

  .error {
    display: none;
    margin-top: 1.2rem;
    padding: 0.85rem 1rem;
    background: var(--red);
    border: 1px solid #f5c6c6;
    border-radius: 8px;
    font-size: 0.85rem;
    color: var(--red-text);
    font-family: 'DM Mono', monospace;
  }
  .error.visible { display: block; }

  .note {
    max-width: 680px;
    margin-top: 1.5rem;
    font-size: 0.78rem;
    color: var(--muted);
    text-align: center;
    line-height: 1.6;
    font-family: 'DM Mono', monospace;
  }
</style>
</head>
<body>

<header>
  <div class="logo">
    <div class="logo-icon">
      <svg viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/><polyline points="10 9 9 9 8 9"/></svg>
    </div>
    <div>
      <h1>BICA PDF → Excel</h1>
      <div class="subtitle">Extracto bancario · Banco BICA SA</div>
    </div>
  </div>
</header>

<div class="card">
  <div class="drop-zone" id="dropZone">
    <input type="file" id="fileInput" accept=".pdf" />
    <div class="drop-icon">
      <svg viewBox="0 0 24 24"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>
    </div>
    <div class="drop-label">Arrastrá el PDF acá</div>
    <div class="drop-hint">o hacé click para seleccionar · solo .pdf</div>
  </div>

  <div class="file-info" id="fileInfo">
    <svg style="width:16px;height:16px;flex-shrink:0" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
    <span class="file-name" id="fileName"></span>
    <span class="file-size" id="fileSize"></span>
  </div>

  <div class="error" id="errorBox"></div>

  <button class="btn" id="convertBtn" disabled>Convertir a Excel</button>

  <div class="progress-wrap" id="progressWrap">
    <div class="progress-label">
      <span id="progressText">Procesando PDF...</span>
      <span id="progressPct">0%</span>
    </div>
    <div class="progress-bar-bg">
      <div class="progress-bar-fill" id="progressFill"></div>
    </div>
  </div>

  <div class="result" id="result">
    <div class="result-header">
      <svg viewBox="0 0 24 24"><path d="M9 11l3 3L22 4"/><path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11"/></svg>
      <span class="result-title">Excel generado correctamente</span>
    </div>
    <div class="stats">
      <div class="stat">
        <div class="stat-label">Movimientos</div>
        <div class="stat-value" id="statTotal">—</div>
      </div>
      <div class="stat">
        <div class="stat-label">Total créditos</div>
        <div class="stat-value green" id="statCred">—</div>
      </div>
      <div class="stat">
        <div class="stat-label">Total débitos</div>
        <div class="stat-value red" id="statDeb">—</div>
      </div>
    </div>
    <a class="download-btn" id="downloadBtn" href="#" download>
      <svg viewBox="0 0 24 24"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
      Descargar Excel
    </a>
  </div>
</div>

<div class="note">
  Procesamiento local · el PDF no sale de tu computadora<br>
  Compatible con extractos del Banco BICA SA (formato Reporting Services)
</div>

<script>
const dropZone   = document.getElementById('dropZone');
const fileInput  = document.getElementById('fileInput');
const fileInfo   = document.getElementById('fileInfo');
const fileName   = document.getElementById('fileName');
const fileSize   = document.getElementById('fileSize');
const convertBtn = document.getElementById('convertBtn');
const progressWrap = document.getElementById('progressWrap');
const progressFill = document.getElementById('progressFill');
const progressText = document.getElementById('progressText');
const progressPct  = document.getElementById('progressPct');
const resultEl   = document.getElementById('result');
const errorBox   = document.getElementById('errorBox');
const downloadBtn= document.getElementById('downloadBtn');

let selectedFile = null;

function fmtSize(b) {
  if (b < 1024) return b + ' B';
  if (b < 1024*1024) return (b/1024).toFixed(0) + ' KB';
  return (b/1024/1024).toFixed(1) + ' MB';
}

function fmtNum(n) {
  return new Intl.NumberFormat('es-AR', {minimumFractionDigits:2, maximumFractionDigits:2}).format(n);
}

function setFile(file) {
  if (!file || !file.name.endsWith('.pdf')) {
    showError('Solo se aceptan archivos .pdf');
    return;
  }
  selectedFile = file;
  fileName.textContent = file.name;
  fileSize.textContent = fmtSize(file.size);
  fileInfo.classList.add('visible');
  convertBtn.disabled = false;
  resultEl.classList.remove('visible');
  errorBox.classList.remove('visible');
}

function showError(msg) {
  errorBox.textContent = msg;
  errorBox.classList.add('visible');
}

function animateProgress(targetPct, duration, label) {
  progressText.textContent = label;
  const start = parseFloat(progressFill.style.width) || 0;
  const diff = targetPct - start;
  const startTime = performance.now();
  function step(now) {
    const elapsed = now - startTime;
    const t = Math.min(elapsed / duration, 1);
    const ease = t < 0.5 ? 2*t*t : -1+(4-2*t)*t;
    const current = start + diff * ease;
    progressFill.style.width = current + '%';
    progressPct.textContent = Math.round(current) + '%';
    if (t < 1) requestAnimationFrame(step);
  }
  requestAnimationFrame(step);
}

fileInput.addEventListener('change', e => {
  if (e.target.files[0]) setFile(e.target.files[0]);
});

dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('drag-over');
  if (e.dataTransfer.files[0]) setFile(e.dataTransfer.files[0]);
});

convertBtn.addEventListener('click', async () => {
  if (!selectedFile) return;
  errorBox.classList.remove('visible');
  resultEl.classList.remove('visible');
  convertBtn.disabled = true;
  progressWrap.classList.add('visible');
  progressFill.style.width = '0%';

  animateProgress(30, 400, 'Leyendo PDF...');

  const formData = new FormData();
  formData.append('file', selectedFile);

  setTimeout(() => animateProgress(65, 1200, 'Extrayendo movimientos...'), 500);

  try {
    const resp = await fetch('/convert', { method: 'POST', body: formData });

    if (!resp.ok) {
      const err = await resp.json().catch(() => ({error: 'Error desconocido'}));
      throw new Error(err.error || 'Error en el servidor');
    }

    animateProgress(90, 400, 'Generando Excel...');

    const stats = JSON.parse(resp.headers.get('X-Stats') || '{}');
    const blob  = await resp.blob();

    setTimeout(() => {
      animateProgress(100, 300, 'Listo!');
      setTimeout(() => {
        progressWrap.classList.remove('visible');
        convertBtn.disabled = false;

        const url = URL.createObjectURL(blob);
        downloadBtn.href = url;
        const baseName = selectedFile.name.replace(/\.pdf$/i, '');
        downloadBtn.download = baseName + '_Movimientos.xlsx';

        document.getElementById('statTotal').textContent = stats.total || '—';
        document.getElementById('statCred').textContent  = stats.creditos ? fmtNum(stats.creditos) : '—';
        document.getElementById('statDeb').textContent   = stats.debitos  ? fmtNum(stats.debitos)  : '—';

        resultEl.classList.add('visible');
        downloadBtn.click();
      }, 400);
    }, 600);

  } catch(err) {
    progressWrap.classList.remove('visible');
    convertBtn.disabled = false;
    showError('Error: ' + err.message);
  }
});
</script>
</body>
</html>
"""

@app.route('/')
def index():
    return HTML

@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return jsonify({'error': 'No se recibió archivo'}), 400

    f = request.files['file']
    if not f.filename.lower().endswith('.pdf'):
        return jsonify({'error': 'El archivo debe ser un PDF'}), 400

    try:
        file_bytes = f.read()
        rows = parse_pdf(file_bytes)

        if not rows:
            return jsonify({'error': 'No se encontraron movimientos en el PDF. ¿Es un extracto del Banco BICA?'}), 400

        excel_buf = build_excel(rows)

        total_deb  = sum(r['debito']  for r in rows if r['debito']  is not None)
        total_cred = sum(r['credito'] for r in rows if r['credito'] is not None)

        import json
        stats = json.dumps({'total': len(rows), 'debitos': total_deb, 'creditos': total_cred})

        response = send_file(
            excel_buf,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='movimientos.xlsx'
        )
        response.headers['X-Stats'] = stats
        return response

    except Exception as e:
        return jsonify({'error': str(e)}), 500

def open_browser():
    import time
    time.sleep(1)
    webbrowser.open('http://localhost:5050')

if __name__ == '__main__':
    print("=" * 55)
    print("  BICA PDF → Excel  |  http://localhost:5050")
    print("  Cerrá esta ventana para detener el servidor")
    print("=" * 55)
    threading.Thread(target=open_browser, daemon=True).start()
    app.run(host='0.0.0.0', port=5050, debug=False)
