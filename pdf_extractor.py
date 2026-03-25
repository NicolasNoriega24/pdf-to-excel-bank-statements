"""
Proyecto: PDF to Excel - Bank Statements
Autor: Nicolás Noriega

Descripción:
Script secundario del Main.py que extrae datos desde PDFs de extractos bancarios
y los convierte en tablas estructuradas en Excel.
"""


import pdfplumber
import re
import pandas as pd
from collections import defaultdict
import os
from openpyxl import load_workbook
from openpyxl.styles import numbers

# =========================
# CONFIGURACIÓN
# =========================
CARPETA_PDFS = r'\\PC Lenovo\Dropbox\CLIENTES\Cliente\Auditoria\2026\Bancos\Macro\Nuevo'
OUTPUT_EXCEL = os.path.join(CARPETA_PDFS, "Resumen_PDF_CBU.xlsx")

regex_fecha = re.compile(r"^\d{1,2}/\d{1,2}/\d{2,4}$")
ROUND_TOP_DECIMALS = 2
TOL_X1_DEBITOS = 30
TOL_X1_CREDITOS = 30
TOL_X1_SALDO = 30
TOL_TOP_LINEA = 1

def convertir_numero_pdf(valor):
    if not valor:
        return None
    valor = valor.replace(".", "").replace(",", ".")
    try:
        return float(valor)
    except:
        return None

# =========================
# DETECTAR ENCABEZADOS Y CBUs
# =========================
def detectar_encabezados_y_cbu(pdf_path):
    cbu_por_encabezado = []
    with pdfplumber.open(pdf_path) as pdf:
        for num_pagina, page in enumerate(pdf.pages, start=1):
            words = page.extract_words()
            if not words:
                continue

            grupos = defaultdict(list)
            for w in words:
                grupos[round(w['top'], ROUND_TOP_DECIMALS)].append(w)

            for top, grupo in grupos.items():
                encabezados = {}
                for w in grupo:
                    txt = w['text'].upper()
                    if txt in ['FECHA','DESCRIPCION','REFERENCIA','DEBITOS','CREDITOS','SALDO']:
                        encabezados[txt] = w

                if len(encabezados) >= 3:
                    tops_anteriores = [round(w['top'], ROUND_TOP_DECIMALS) for w in words if round(w['top'], ROUND_TOP_DECIMALS) < top]
                    cbu_texto = None
                    for t in sorted(tops_anteriores, reverse=True):
                        linea = ' '.join([w['text'] for w in grupos[t]])
                        match_cbu = re.search(r"(Clave Bancaria Uniforme para Debito Directo:\s*\d+-\d+-\d+-\d+)", linea, re.IGNORECASE)
                        if match_cbu:
                            cbu_texto = match_cbu.group(1)
                            break
                    if cbu_texto:
                        top_global = (num_pagina - 1) * 10000 + t
                        cbu_por_encabezado.append({
                            'pagina': num_pagina,
                            'top_encabezado': top,
                            'top_cbu': t,
                            'top_global': top_global,
                            'cbu': cbu_texto,
                            'encabezados': encabezados
                        })
    return sorted(cbu_por_encabezado, key=lambda x: x['top_global'])

# =========================
# EXTRAER REGISTROS POR CBU
# =========================
def extraer_registros(pdf_path, cbu_por_encabezado):
    cbu_dfs = defaultdict(pd.DataFrame)
    with pdfplumber.open(pdf_path) as pdf:
        idx = 0
        total = len(cbu_por_encabezado)
        while idx < total:
            item = cbu_por_encabezado[idx]
            cbu_texto = item['cbu']
            registros = []

            while idx < total and cbu_por_encabezado[idx]['cbu'] == cbu_texto:
                pagina = cbu_por_encabezado[idx]['pagina']
                encabezados = cbu_por_encabezado[idx]['encabezados']
                top_actual = cbu_por_encabezado[idx]['top_cbu']

                top_siguiente = float('inf')
                for j in range(idx + 1, total):
                    if cbu_por_encabezado[j]['cbu'] != cbu_texto and cbu_por_encabezado[j]['pagina'] == pagina:
                        top_siguiente = cbu_por_encabezado[j]['top_cbu']
                        break

                page = pdf.pages[pagina - 1]
                palabras = page.extract_words()
                lineas = defaultdict(list)
                for p in palabras:
                    lineas[round(p['top'], ROUND_TOP_DECIMALS)].append(p)

                x0_fecha = encabezados['FECHA']['x0']
                x0_ref = encabezados['REFERENCIA']['x0']
                x1_debitos_enc = encabezados['DEBITOS']['x1']
                x1_creditos_enc = encabezados['CREDITOS']['x1']
                x1_saldo_enc = encabezados['SALDO']['x1']

                for top_linea, grupo in sorted(lineas.items()):
                    if top_linea <= encabezados['FECHA']['top']:
                        continue
                    if not (top_actual - TOL_TOP_LINEA <= top_linea < top_siguiente):
                        continue

                    fecha = None
                    for p in grupo:
                        texto = p['text'].strip()
                        if regex_fecha.match(texto) and p['x0'] <= x0_fecha:
                            fecha = texto
                            break
                    if not fecha:
                        continue

                    referencia = ''
                    descripcion = ''
                    debitos = None
                    creditos = None
                    saldo = None

                    x0_descripcion = None
                    for p in grupo:
                        if p['x0'] > x0_fecha:
                            x0_descripcion = p['x0']
                            break
                    for p in grupo:
                        texto = p['text'].strip()
                        x0 = p['x0']
                        if x0_descripcion and x0 >= x0_descripcion and x0 < x0_ref:
                            descripcion += (' ' + texto if descripcion else texto)
                        if abs(x0 - x0_ref) < 1e-6:
                            referencia = texto

                    candidatos_debitos = [p for p in grupo if abs(p['x1'] - x1_debitos_enc) <= TOL_X1_DEBITOS]
                    if candidatos_debitos:
                        token_debito = min(candidatos_debitos, key=lambda p: abs(p['x1'] - x1_debitos_enc))
                        debitos = convertir_numero_pdf(token_debito['text'])

                    candidatos_creditos = [p for p in grupo if abs(p['x1'] - x1_creditos_enc) <= TOL_X1_CREDITOS]
                    if candidatos_creditos:
                        token_credito = min(candidatos_creditos, key=lambda p: abs(p['x1'] - x1_creditos_enc))
                        creditos = convertir_numero_pdf(token_credito['text'])

                    candidatos_saldo = [p for p in grupo if abs(p['x1'] - x1_saldo_enc) <= TOL_X1_SALDO]
                    if candidatos_saldo:
                        token_saldo = min(candidatos_saldo, key=lambda p: abs(p['x1'] - x1_saldo_enc))
                        saldo = convertir_numero_pdf(token_saldo['text'])

                    registros.append({
                        'FECHA': fecha,
                        'DESCRIPCION': descripcion,
                        'REFERENCIA': referencia,
                        'DEBITOS': debitos if debitos is not None else 0,
                        'CREDITOS': creditos if creditos is not None else 0,
                        'SALDO': saldo if saldo is not None else 0
                    })

                idx += 1

            if not registros:
                registros = [{'FECHA': '', 'DESCRIPCION': 'SIN MOVIMIENTOS', 'REFERENCIA': '', 'DEBITOS': 0, 'CREDITOS': 0, 'SALDO': 0}]

            df_nuevo = pd.DataFrame(registros)
            if cbu_texto in cbu_dfs:
                cbu_dfs[cbu_texto] = pd.concat([cbu_dfs[cbu_texto], df_nuevo], ignore_index=True)
            else:
                cbu_dfs[cbu_texto] = df_nuevo

    return cbu_dfs

# =========================
# EXPORTAR A EXCEL
# =========================
def exportar_excel(cbu_dict, archivo_salida):
    with pd.ExcelWriter(archivo_salida, engine='openpyxl') as writer:
        fila_actual = 0
        for cbu, df in cbu_dict.items():
            ws_cbu = pd.DataFrame([[f"{cbu}"] + [""]*(df.shape[1]-1)], columns=df.columns)
            ws_cbu.to_excel(writer, index=False, header=False, startrow=fila_actual)
            fila_actual += 1

            df.to_excel(writer, index=False, header=True, startrow=fila_actual)
            fila_actual += len(df) + 2

    print(f"\nArchivo Excel creado en: {archivo_salida}")

# =========================
# MAIN
# =========================
pdf_files = [f for f in os.listdir(CARPETA_PDFS) if f.lower().endswith('.pdf')]
pdf_files.sort(key=lambda x: int(re.search(r'\d+', x).group()) if re.search(r'\d+', x) else float('inf'))

cbu_dict_total = {}

for pdf_nombre in pdf_files:
    pdf_path = os.path.join(CARPETA_PDFS, pdf_nombre)
    print(f"\nProcesando PDF: {pdf_nombre}")
    cbu_por_encabezado = detectar_encabezados_y_cbu(pdf_path)
    if cbu_por_encabezado:
        dfs_pdf = extraer_registros(pdf_path, cbu_por_encabezado)
        # Concatenar al diccionario total
        for cbu, df in dfs_pdf.items():
            if cbu in cbu_dict_total:
                cbu_dict_total[cbu] = pd.concat([cbu_dict_total[cbu], df], ignore_index=True)
            else:
                cbu_dict_total[cbu] = df
    else:
        print(f"No encontré encabezados y CBUs en el PDF: {pdf_nombre}")

if cbu_dict_total:
    exportar_excel(cbu_dict_total, OUTPUT_EXCEL)
else:
    print("No se encontraron datos en ningún PDF de la carpeta.")
