import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

# Configuración de la página
st.set_page_config(
    page_title="Extractor Factura Movistar",
    page_icon="📄",
    layout="centered"
)

st.title("📄 Extractor Factura Movistar")
st.markdown("Sube el PDF de la factura y descarga los datos en Excel automáticamente.")

# Rangos X de cada columna según el layout del PDF Movistar
COLUMNAS = [
    ("Línea",                             0,    53),
    ("Plan",                             53,    85),
    ("Cargo Fijo Mensual",               85,   130),
    ("Consumo Adicional Voz",           130,   165),
    ("Mensajes",                         165,   205),
    ("Larga Distancia Internacional",   205,   245),
    ("Roaming",                          245,   285),
    ("Datos",                            285,   325),
    ("Servicios movistar",              325,   365),
    ("Servicios Especiales",            365,   405),
    ("Otros Cargos Facturados",         405,   445),
    ("Descuentos",                       445,   485),
    ("Iva 19%",                          485,   525),
    ("Impto. al Consumo y Otros Grav.", 525,   565),
    ("SUBTOTAL",                         565,   700),
]

def es_linea_valida(v):
    return bool(re.match(r'^\d{10}$', str(v).strip()))

def es_fila_total(v):
    return str(v).strip().lower() == 'total'

def limpiar_valor(texto):
    if not texto:
        return 0
    t = texto.strip()
    negativo = '-' in t
    t = re.sub(r'[-\$\s,]', '', t)
    if not t or not t.isdigit():
        return 0
    return -int(t) if negativo else int(t)

def extraer_pagina(page):
    words = page.extract_words(x_tolerance=3, y_tolerance=3)
    lineas = {}
    for w in words:
        top = round(w['top'])
        lineas.setdefault(top, []).append(w)

    filas = []
    for top in sorted(lineas.keys()):
        palabras = sorted(lineas[top], key=lambda x: x['x0'])
        celda = {col: [] for col, _, _ in COLUMNAS}
        for w in palabras:
            cx = (w['x0'] + w['x1']) / 2
            for col, xmin, xmax in COLUMNAS:
                if xmin <= cx < xmax:
                    celda[col].append(w['text'])
                    break

        linea_val = ' '.join(celda['Línea']).strip()

        if es_fila_total(linea_val):
            return filas, True

        if not es_linea_valida(linea_val):
            continue

        fila = {
            'Línea': linea_val,
            'Plan':  ' '.join(celda['Plan']).strip()
        }
        for col, _, _ in COLUMNAS:
            if col in ('Línea', 'Plan'):
                continue
            fila[col] = limpiar_valor(' '.join(celda[col]).strip())
        filas.append(fila)

    return filas, False

def procesar_pdf(archivo):
    todas_las_filas = []
    with pdfplumber.open(archivo) as pdf:
        total_paginas = len(pdf.pages)
        progress = st.progress(0, text="Procesando páginas...")
        for i in range(2, total_paginas):
            filas, encontro_total = extraer_pagina(pdf.pages[i])
            todas_las_filas.extend(filas)
            progress.progress(
                int((i - 1) / (total_paginas - 2) * 100),
                text=f"Procesando página {i+1} de {total_paginas}..."
            )
            if encontro_total:
                break
        progress.empty()
    return pd.DataFrame(todas_las_filas)

def convertir_a_excel(df):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Abonados')
        ws = writer.sheets['Abonados']
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 30)
    return buffer.getvalue()

# Interfaz principal
archivo = st.file_uploader("Selecciona el PDF de la factura", type="pdf")

if archivo:
    with st.spinner("Extrayendo datos..."):
        df = procesar_pdf(archivo)

    if df.empty:
        st.error("No se encontraron datos de abonados en el PDF.")
    else:
        st.success(f"✅ {len(df)} abonados extraídos correctamente")

        # Resumen
        col1, col2, col3 = st.columns(3)
        col1.metric("Total abonados", len(df))
        col2.metric("Planes distintos", df['Plan'].nunique())
        col3.metric("SUBTOTAL", f"${df['SUBTOTAL'].sum():,.0f}")

        # Resumen por plan
        st.subheader("Resumen por Plan")
        resumen = df.groupby('Plan').agg(
            Abonados=('Línea', 'count'),
            SUBTOTAL=('SUBTOTAL', 'sum')
        ).reset_index()
        st.dataframe(resumen, use_container_width=True)

        # Vista previa
        st.subheader("Vista previa de datos")
        st.dataframe(df.head(10), use_container_width=True)

        # Botón de descarga
        nombre_excel = archivo.name.replace('.pdf', '_datos.xlsx').replace('.PDF', '_datos.xlsx')
        st.download_button(
            label="⬇️ Descargar Excel",
            data=convertir_a_excel(df),
            file_name=nombre_excel,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
