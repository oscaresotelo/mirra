"""
MIRRA — Generador Libro IVA Digital (ARCA/AFIP)
Acepta: Plantilla XLS/XLSX de Contabilium  |  CSV descargado del portal AFIP
Genera: 4 archivos TXT con el formato exacto de posiciones fijas requerido por ARCA
"""

import streamlit as st
import pandas as pd
import csv
import os, io, tempfile, zipfile
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation

# ─── CONFIG ──────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="MIRRA · Libro IVA Digital",
    page_icon="📒",
    layout="centered",
)

st.markdown("""
<style>
[data-testid="stAppViewContainer"] { background: #f4f6f8; }
[data-testid="stHeader"]           { background: transparent; }

[data-testid="stSidebar"]          { background: #003d1f; }
[data-testid="stSidebar"] *        { color: #c8e6c9 !important; }
[data-testid="stSidebar"] hr       { border-color: #1b5e20; }
[data-testid="stSidebar"] a        { color: #a5d6a7 !important; }

div[data-testid="stButton"] > button {
    background: #2e7d32 !important; color: white !important;
    border: none !important; border-radius: 7px !important;
    font-weight: 600 !important; width: 100%;
    padding: 0.65rem 1.5rem !important; letter-spacing: 0.02em;
    transition: background 0.2s;
}
div[data-testid="stButton"] > button:hover { background: #1b5e20 !important; }
div[data-testid="stButton"] > button:disabled { opacity: 0.4; }

div[data-testid="stDownloadButton"] > button {
    background: white !important; color: #2e7d32 !important;
    border: 1.5px solid #2e7d32 !important; border-radius: 7px !important;
    font-weight: 500 !important; width: 100%;
    transition: all 0.2s;
}
div[data-testid="stDownloadButton"] > button:hover {
    background: #e8f5e9 !important;
}

[data-testid="stFileUploader"] {
    border: 2px dashed #81c784; border-radius: 10px;
    padding: 0.5rem; background: #f9fff9;
}
[data-testid="stMetric"] {
    background: white; border: 1px solid #e0e0e0;
    border-radius: 10px; padding: 0.8rem 1rem;
}
details { border: 1px solid #dce8dc !important; border-radius: 9px !important; }
</style>
""", unsafe_allow_html=True)


# ─── TABLAS ARCA ─────────────────────────────────────────────────────────────

TIPOS_CBTE = {
    '1':'001','2':'002','3':'003','4':'004','5':'005',
    '6':'006','7':'007','8':'008','9':'009','10':'010',
    '11':'011','12':'012','13':'013','15':'015','19':'019',
    '20':'020','21':'021','51':'051','52':'052','53':'053',
    '54':'054','60':'060','61':'061','63':'063','64':'064',
}

# Código alícuota IVA → 4 dígitos ARCA
ALICUOTAS = {
    '0':'0003', '2.5':'0009', '5':'0008',
    '10.5':'0004', '21':'0005', '27':'0006',
}

# Columnas del CSV de AFIP con IVA discriminado por alícuota
ALICUOTA_COLS = [
    ('IVA 2,5%',  'Imp. Neto Gravado IVA 2,5%',  '2.5'),
    ('IVA 5%',    'Imp. Neto Gravado IVA 5%',    '5'),
    ('IVA 10,5%', 'Imp. Neto Gravado IVA 10,5%', '10.5'),
    ('IVA 21%',   'Imp. Neto Gravado IVA 21%',   '21'),
    ('IVA 27%',   'Imp. Neto Gravado IVA 27%',   '27'),
]

MONEDAS    = {'$':'PES', 'U$S':'DOL', 'PES':'PES', 'DOL':'DOL'}
CODIGOS_NC = {'3','8','13','21','53'}   # Notas de crédito → cod_op = 'R'


# ─── HELPERS ─────────────────────────────────────────────────────────────────

def to_decimal(s):
    if not s or str(s).strip() in ('','nan','NaN','None'): return Decimal('0')
    s = str(s).strip().replace('.','').replace(',','.')
    try:    return Decimal(s)
    except: return Decimal('0')

def fmt_importe(valor):
    """13 enteros + 2 decimales implícitos, sin punto, relleno a izquierda con 0."""
    d = Decimal(str(valor)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
    return str(abs(d)).replace('.','').zfill(15)

def fmt_num(valor, largo):
    s = ''.join(c for c in str(valor or '') if c.isdigit())
    return (s or '0').zfill(largo)

def fmt_alfa(valor, largo):
    s = str(valor) if valor and str(valor) not in ('nan','NaN','None') else ''
    return s[:largo].ljust(largo)

def parse_fecha(s):
    s = str(s).strip()[:10]
    for fmt in ('%d/%m/%Y','%Y-%m-%d','%d-%m-%Y'):
        try:    return datetime.strptime(s, fmt).strftime('%Y%m%d')
        except: pass
    return '00000000'

def tipo_cbte_txt(t):
    return TIPOS_CBTE.get(str(t).strip().split('.')[0], str(t).zfill(3))

def cod_operacion(t):
    return 'R' if str(t).strip().split('.')[0] in CODIGOS_NC else ' '

def moneda_txt(m):
    return MONEDAS.get(str(m).strip(), 'PES')

def tipo_cambio_txt(tc):
    d = to_decimal(tc) or Decimal('1')
    e   = int(d)
    dec = str((d - e).quantize(Decimal('0.000001')))[2:]
    return str(e).zfill(4) + dec.ljust(6,'0')[:6]

def nro_doc(n, largo=20):
    return ''.join(c for c in str(n or '') if c.isdigit()).zfill(largo)

def detectar_periodo(rows):
    for row in rows:
        f = parse_fecha(row.get('Fecha',''))
        if f != '00000000': return f[:6]
    return datetime.now().strftime('%Y%m')


# ─── GENERADORES ─────────────────────────────────────────────────────────────

def _alics(row):
    """Extrae lista de (cod_alicuota, neto, iva) de una fila."""
    alics = []
    for col_iva, col_neto, alic in ALICUOTA_COLS:
        iva_v  = to_decimal(row.get(col_iva,  0))
        neto_v = to_decimal(row.get(col_neto, 0))
        if neto_v != 0 or iva_v != 0:
            alics.append((alic, neto_v, iva_v))
    if not alics:
        neto_t = to_decimal(row.get('Imp. Neto Gravado Total', 0))
        iva_t  = to_decimal(row.get('Total IVA', 0))
        if neto_t != 0:
            alics = [('21', neto_t, iva_t)]
    return alics


def generar_ventas(rows):
    """VENTAS_CBTE (266 chars) + VENTAS_ALICUOTAS (62 chars)."""
    cbte_lines, alic_lines, errores = [], [], []

    for i, row in enumerate(rows):
        try:
            fecha     = parse_fecha(row.get('Fecha',''))
            tipo      = str(row.get('Tipo de Comprobante','')).strip().split('.')[0]
            pto_vta   = fmt_num(row.get('Punto de Venta',  0), 5)
            nro_desde = fmt_num(row.get('Número Desde',    0), 20)
            nro_hasta = fmt_num(row.get('Número Hasta', nro_desde), 20)
            cod_doc   = fmt_num(row.get('Tipo Doc. Receptor', 99), 2)
            nro_id    = nro_doc(row.get('Nro. Doc. Receptor', 0))
            nombre    = fmt_alfa(row.get('Denominación Receptor',''), 30)
            moneda    = moneda_txt(row.get('Moneda','$'))
            tc        = tipo_cambio_txt(row.get('Tipo Cambio','1'))

            total      = fmt_importe(to_decimal(row.get('Imp. Total',           0)))
            no_gravado = fmt_importe(to_decimal(row.get('Imp. Neto No Gravado', 0)))
            exentas    = fmt_importe(to_decimal(row.get('Imp. Op. Exentas',     0)))
            otros_trib = fmt_importe(to_decimal(row.get('Otros Tributos',       0)))
            cero15     = fmt_importe(0)
            cod_op     = cod_operacion(tipo)
            alics      = _alics(row)
            cant_alic  = str(len(alics)) if alics else '1'

            # ── CBTE: 266 chars ──────────────────────────────────────────
            # C1 fecha(8) C2 tipo(3) C3 pto(5) C4 nro_desde(20) C5 nro_hasta(20)
            # C6 cod_doc(2) C7 nro_id(20) C8 nombre(30)
            # C9 total(15) C10 no_gravado(15) C11 perc_no_cat(15) C12 exentas(15)
            # C13 perc_nac(15) C14 perc_iibb(15) C15 perc_mun(15) C16 imp_int(15)
            # C17 moneda(3) C18 tc(10) C19 cant_alic(1) C20 cod_op(1)
            # C21 otros_trib(15) C22 fecha_vto(8)
            linea = (
                fecha + tipo_cbte_txt(tipo) + pto_vta + nro_desde + nro_hasta +
                cod_doc + nro_id + nombre +
                total + no_gravado + cero15 + exentas +
                cero15 + cero15 + cero15 + cero15 +
                moneda + tc + cant_alic + cod_op +
                otros_trib + '00000000'
            )
            assert len(linea) == 266, f"CBTE len={len(linea)}"
            cbte_lines.append(linea)

            # ── ALÍCUOTAS: 62 chars ──────────────────────────────────────
            # tipo(3) pto(5) nro_desde(20) neto_grav(15) cod_alic(4) imp_liq(15)
            for alic, neto_v, iva_v in alics:
                la = (
                    tipo_cbte_txt(tipo) + pto_vta + nro_desde +
                    fmt_importe(neto_v) + ALICUOTAS.get(alic,'0005') + fmt_importe(iva_v)
                )
                assert len(la) == 62, f"ALIC len={len(la)}"
                alic_lines.append(la)

        except Exception as e:
            errores.append(f"Ventas fila {i+1}: {e}")

    return cbte_lines, alic_lines, errores


def generar_compras(rows):
    """COMPRAS_CBTE (325 chars) + COMPRAS_ALICUOTAS (84 chars)."""
    cbte_lines, alic_lines, errores = [], [], []

    for i, row in enumerate(rows):
        try:
            fecha    = parse_fecha(row.get('Fecha',''))
            tipo     = str(row.get('Tipo de Comprobante','')).strip().split('.')[0]
            pto_vta  = fmt_num(row.get('Punto de Venta', 0), 5)
            nro_cbte = fmt_num(row.get('Número Desde',   0), 20)
            despacho = fmt_alfa('', 16)   # vacío para compras locales
            cod_doc  = fmt_num(row.get('Tipo Doc. Vendedor',
                               row.get('Tipo Doc. Receptor', 80)), 2)
            nro_id   = nro_doc(row.get('Nro. Doc. Vendedor',
                               row.get('Nro. Doc. Receptor', 0)))
            nombre   = fmt_alfa(row.get('Denominación Vendedor',
                                row.get('Denominación Receptor','')), 30)
            moneda   = moneda_txt(row.get('Moneda','$'))
            tc       = tipo_cambio_txt(row.get('Tipo Cambio','1'))

            total      = fmt_importe(to_decimal(row.get('Imp. Total',           0)))
            no_gravado = fmt_importe(to_decimal(row.get('Imp. Neto No Gravado', 0)))
            exentas    = fmt_importe(to_decimal(row.get('Imp. Op. Exentas',     0)))
            cred_fisc  = fmt_importe(to_decimal(row.get('Total IVA',            0)))
            otros_trib = fmt_importe(to_decimal(row.get('Otros Tributos',       0)))
            cero15     = fmt_importe(0)
            cod_op     = cod_operacion(tipo)
            alics      = _alics(row)
            cant_alic  = str(len(alics)) if alics else '1'

            # ── CBTE: 325 chars ──────────────────────────────────────────
            # C1 fecha(8) C2 tipo(3) C3 pto(5) C4 nro(20) C5 despacho(16)
            # C6 cod_doc(2) C7 nro_id(20) C8 nombre(30)
            # C9 total(15) C10 no_gravado(15) C11 exentas(15)
            # C12 perc_iva(15) C13 perc_nac(15) C14 perc_iibb(15)
            # C15 perc_mun(15) C16 imp_int(15)
            # C17 moneda(3) C18 tc(10) C19 cant_alic(1) C20 cod_op(1)
            # C21 cred_fisc(15) C22 otros_trib(15)
            # C23 cuit_emisor(11) C24 denom_emisor(30) C25 iva_comision(15)
            linea = (
                fecha + tipo_cbte_txt(tipo) + pto_vta + nro_cbte + despacho +
                cod_doc + nro_id + nombre +
                total + no_gravado + exentas +
                cero15 + cero15 + cero15 + cero15 + cero15 +
                moneda + tc + cant_alic + cod_op +
                cred_fisc + otros_trib +
                fmt_num(0, 11) + fmt_alfa('', 30) + cero15
            )
            assert len(linea) == 325, f"CBTE len={len(linea)}"
            cbte_lines.append(linea)

            # ── ALÍCUOTAS: 84 chars ──────────────────────────────────────
            # tipo(3) pto(5) nro(20) cod_doc(2) nro_id(20) neto(15) alic(4) iva(15)
            for alic, neto_v, iva_v in alics:
                la = (
                    tipo_cbte_txt(tipo) + pto_vta + nro_cbte +
                    cod_doc + nro_id +
                    fmt_importe(neto_v) + ALICUOTAS.get(alic,'0005') + fmt_importe(iva_v)
                )
                assert len(la) == 84, f"ALIC len={len(la)}"
                alic_lines.append(la)

        except Exception as e:
            errores.append(f"Compras fila {i+1}: {e}")

    return cbte_lines, alic_lines, errores


# ─── LECTORES ─────────────────────────────────────────────────────────────────

def leer_xls_contabilium(uploaded_file):
    """Plantilla XLS/XLSX exportada de Contabilium (campo Número = PPPPP-NNNNNNN)."""
    ext = os.path.splitext(uploaded_file.name)[1].lower()
    with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as tmp:
        tmp.write(uploaded_file.getvalue())
        tmp_path = tmp.name

    if ext == '.xls':
        out_dir = tempfile.mkdtemp()
        os.system(f'libreoffice --headless --convert-to xlsx "{tmp_path}" --outdir "{out_dir}" 2>/dev/null')
        base = os.path.splitext(os.path.basename(tmp_path))[0]
        tmp_path = os.path.join(out_dir, base + '.xlsx')

    df = pd.read_excel(tmp_path, sheet_name='Comprobantes', dtype=str).fillna('')

    rows = []
    for _, r in df.iterrows():
        numero = str(r.get('Número','')).strip()
        pto, nro = numero.split('-',1) if '-' in numero else ('00001', numero)

        cuit_raw = str(r.get('Nro. Documento / Cuit','')).strip()
        try:    cuit_raw = str(int(float(cuit_raw))) if cuit_raw else '0'
        except: cuit_raw = '0'

        rows.append({
            'Fecha':                     str(r.get('Fecha','')).strip()[:10],
            'Tipo de Comprobante':       str(r.get('Tipo comprobante AFIP','')).strip().split('.')[0],
            'Punto de Venta':            pto.lstrip('0') or '0',
            'Número Desde':              nro.lstrip('0') or '0',
            'Número Hasta':              nro.lstrip('0') or '0',
            'Tipo Doc. Receptor':        str(r.get('Tipo documento','99')).strip().split('.')[0],
            'Nro. Doc. Receptor':        cuit_raw,
            'Denominación Receptor':     str(r.get('Razón social','')).strip(),
            'Tipo Cambio':               str(r.get('Cotización','1')).strip() or '1',
            'Moneda':                    '$',
            'IVA 21%':                   str(r.get('IVA 21','')).strip(),
            'Imp. Neto Gravado IVA 21%': str(r.get('Neto Gravado 21','')).strip(),
            'Imp. Neto Gravado Total':   str(r.get('Neto Gravado 21','')).strip(),
            'Imp. Neto No Gravado':      '0',
            'Imp. Op. Exentas':          '0',
            'Otros Tributos':            '0',
            'Total IVA':                 str(r.get('IVA 21','')).strip(),
            'Imp. Total':                str(r.get('Total','')).strip(),
        })
    return rows


def leer_csv_afip(uploaded_file):
    """CSV de comprobantes descargado del portal AFIP (CUIT en fila 0 col 0)."""
    content = uploaded_file.read().decode('utf-8-sig')
    sep = ';' if content.count(';') > content.count(',') else ','
    all_lines = [l for l in content.splitlines() if l.strip()]

    first_col = all_lines[0].split(sep)[0].strip()
    if first_col.isdigit() and len(first_col) == 11:
        reader = csv.DictReader(all_lines, delimiter=sep)
        rows = []
        for row in reader:
            fecha_key = list(row.keys())[0]
            row['Fecha'] = row[fecha_key]
            rows.append(dict(row))
        return rows
    else:
        reader = csv.DictReader(all_lines, delimiter=sep)
        return [dict(r) for r in reader if any(v.strip() for v in r.values())]


def leer_archivo(uploaded_file):
    ext = os.path.splitext(uploaded_file.name)[1].lower()
    if ext in ('.xls','.xlsx'):
        return leer_xls_contabilium(uploaded_file), 'Excel (Contabilium)'
    return leer_csv_afip(uploaded_file), 'CSV AFIP'


def lineas_a_bytes(lines):
    return ('\r\n'.join(lines) + '\r\n').encode('utf-8')

def hacer_zip(archivos):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        for nombre, contenido in archivos.items():
            zf.writestr(nombre, contenido)
    return buf.getvalue()


# ─── SIDEBAR ─────────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown("## 📒 MIRRA")
    st.markdown("**Gestión IVA Digital**")
    st.markdown("---")
    st.markdown("""
**Archivos generados:**
- `VENTAS_CBTE` — 266 chars/línea
- `VENTAS_ALICUOTAS` — 62 chars/línea
- `COMPRAS_CBTE` — 325 chars/línea
- `COMPRAS_ALICUOTAS` — 84 chars/línea
""")
    st.markdown("---")
    st.markdown("""
**Formatos aceptados:**
- Plantilla XLS/XLSX (Contabilium)
- CSV del portal AFIP
""")
    st.markdown("---")
    st.caption(f"Estudio Contable MIRRA · {datetime.now().year}")


# ─── HEADER ──────────────────────────────────────────────────────────────────

st.markdown("""
<div style="background:linear-gradient(135deg,#1b5e20 0%,#2e7d32 60%,#388e3c 100%);
            padding:2rem 2rem 1.75rem;border-radius:14px;margin-bottom:2rem;">
    <div style="font-size:0.75rem;font-weight:600;letter-spacing:0.12em;
                color:#a5d6a7;text-transform:uppercase;margin-bottom:0.4rem;">
        Estudio Contable MIRRA
    </div>
    <h1 style="color:white;margin:0;font-size:1.75rem;font-weight:700;letter-spacing:-0.5px;">
        📒 Libro IVA Digital
    </h1>
    <p style="color:#c8e6c9;margin:0.5rem 0 0;font-size:0.9rem;">
        Generador de archivos TXT · Formato ARCA/AFIP · Posiciones fijas
    </p>
</div>
""", unsafe_allow_html=True)


# ─── CARGA DE ARCHIVOS ───────────────────────────────────────────────────────

col1, col2 = st.columns(2)

with col1:
    st.markdown("#### 📤 Ventas")
    st.caption("XLS/XLSX de Contabilium **o** CSV del portal AFIP")
    file_v = st.file_uploader(
        "ventas", type=['csv','xls','xlsx'],
        key='ventas', label_visibility='collapsed'
    )

with col2:
    st.markdown("#### 📥 Compras *(opcional)*")
    st.caption("CSV del portal AFIP de comprobantes recibidos")
    file_c = st.file_uploader(
        "compras", type=['csv','xls','xlsx'],
        key='compras', label_visibility='collapsed'
    )

with st.expander("⚙️ Configuración avanzada"):
    periodo_manual = st.text_input(
        "Forzar período (YYYYMM)",
        placeholder="Ej: 202601 — si lo dejás vacío se detecta automáticamente",
        max_chars=6,
    )

st.divider()

# ─── BOTÓN PROCESAR ──────────────────────────────────────────────────────────

if not file_v:
    st.info("👆 Cargá el archivo de ventas para comenzar.")

if st.button("🚀 GENERAR ARCHIVOS TXT", disabled=(file_v is None)):

    archivos_out, todos_errores = {}, []

    with st.spinner("Procesando..."):
        try:
            rows_v, fmt_v = leer_archivo(file_v)
            periodo = periodo_manual.strip() if periodo_manual.strip() else detectar_periodo(rows_v)

            cbte_v, alic_v, err_v = generar_ventas(rows_v)
            todos_errores.extend(err_v)

            archivos_out[f"LIBRO_IVA_DIGITAL_VENTAS_CBTE_{periodo}.txt"]       = lineas_a_bytes(cbte_v)
            archivos_out[f"LIBRO_IVA_DIGITAL_VENTAS_ALICUOTAS_{periodo}.txt"]  = lineas_a_bytes(alic_v)

            fmt_c, n_compras = '', 0
            if file_c:
                rows_c, fmt_c = leer_archivo(file_c)
                n_compras = len(rows_c)
                cbte_c, alic_c, err_c = generar_compras(rows_c)
                todos_errores.extend(err_c)
                archivos_out[f"LIBRO_IVA_DIGITAL_COMPRAS_CBTE_{periodo}.txt"]      = lineas_a_bytes(cbte_c)
                archivos_out[f"LIBRO_IVA_DIGITAL_COMPRAS_ALICUOTAS_{periodo}.txt"] = lineas_a_bytes(alic_c)

            st.session_state.resultado = {
                "archivos":   archivos_out,
                "periodo":    periodo,
                "fmt_v":      fmt_v,
                "fmt_c":      fmt_c,
                "n_ventas":   len(rows_v),
                "n_compras":  n_compras,
                "errores":    todos_errores,
            }

        except Exception as e:
            st.error(f"❌ Error crítico al procesar: {e}")
            st.stop()


# ─── RESULTADOS ──────────────────────────────────────────────────────────────

if 'resultado' in st.session_state:
    res      = st.session_state.resultado
    periodo  = res["periodo"]
    archivos = res["archivos"]

    # ── Estado ──
    if res["errores"]:
        st.warning(f"⚠️ Completado con {len(res['errores'])} advertencia(s).")
        with st.expander("Ver advertencias"):
            for e in res["errores"]:
                st.write(f"- {e}")
    else:
        st.success("✅ Archivos generados correctamente, sin errores.")

    # ── Info período ──
    info = f"**Período {periodo[:4]}/{periodo[4:]}** · Ventas: {res['fmt_v']}"
    if res["n_compras"]:
        info += f" · Compras: {res['fmt_c']}"
    st.markdown(info)

    # ── Métricas ──
    cols_m = st.columns(len(archivos))
    for col, (nombre, contenido) in zip(cols_m, archivos.items()):
        n_reg = contenido.count(b'\r\n')
        label = (nombre
                 .replace(f'_{periodo}','')
                 .replace('.txt','')
                 .replace('LIBRO_IVA_DIGITAL_',''))
        col.metric(label, f"{n_reg} reg.")

    st.divider()

    # ── Descargas ──
    st.markdown("#### 📦 Descargar")

    zip_bytes = hacer_zip(archivos)
    st.download_button(
        label=f"📦 Descargar todo en ZIP  —  LID_{periodo}.zip",
        data=zip_bytes,
        file_name=f"LID_{periodo}.zip",
        mime="application/zip",
    )

    with st.expander("Descargar archivos individuales"):
        cols_d = st.columns(2)
        for i, (nombre, contenido) in enumerate(archivos.items()):
            cols_d[i % 2].download_button(
                label=f"⬇ {nombre}",
                data=contenido,
                file_name=nombre,
                mime="text/plain",
                key=f"dl_{nombre}",
            )