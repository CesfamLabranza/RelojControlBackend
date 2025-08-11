# procesador.py
# -*- coding: utf-8 -*-

import re
from io import BytesIO
from datetime import datetime, time

import pandas as pd
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# ==========================
# Utilidades de fecha/tiempo
# ==========================

def normalizar_fecha(fecha):
    """Convierte a date soportando dd-mm-YYYY, dd/mm/YYYY o datetime."""
    if isinstance(fecha, datetime):
        return fecha.date()
    try:
        s = str(fecha).strip()
        if not s or s.lower() == "none":
            return None
        if "-" in s:
            # Formato dd-mm-YYYY
            return datetime.strptime(s, "%d-%m-%Y").date()
        if "/" in s:
            # Formato dd/mm/YYYY
            return datetime.strptime(s, "%d/%m/%Y").date()
    except Exception:
        return None
    return None


def convertir_a_hhmm(horas):
    """Convierte horas (float) a HH:MM, redondeando a minutos."""
    minutos = int(round(horas * 60))
    return f"{minutos // 60:02d}:{minutos % 60:02d}"


def minutos_a_hhmm(minutos):
    """Convierte minutos (int) a HH:MM."""
    return f"{minutos // 60:02d}:{minutos % 60:02d}"


def obtener_horario_turno(turno, dia_semana):
    """
    Extrae las horas desde el texto de turno.
    Espera patrón tipo 'HH:MM' repetido. Para Vi (4) permite otros dos.
    """
    horas = re.findall(r"\d{1,2}:\d{2}", str(turno))
    if len(horas) < 2:
        return None, None

    if dia_semana in (0, 1, 2, 3):  # Lu-Ju
        return horas[0], horas[1]
    if dia_semana == 4 and len(horas) >= 4:  # Vi
        return horas[2], horas[3]
    return None, None


def calcular_atraso(entrada, fecha, turno):
    """Minutos de atraso respecto al turno del día."""
    if not entrada or str(entrada).strip() in ("-", "", "None"):
        return 0

    fecha_dt = normalizar_fecha(fecha)
    if not fecha_dt:
        return 0

    try:
        entrada_dt = datetime.strptime(str(entrada), "%H:%M:%S").time()
    except Exception:
        # A veces viene 'HH:MM'
        try:
            entrada_dt = datetime.strptime(str(entrada), "%H:%M").time()
        except Exception:
            return 0

    dia_semana = fecha_dt.weekday()
    inicio_str, _ = obtener_horario_turno(turno, dia_semana)
    if not inicio_str:
        return 0

    hora_inicio = datetime.strptime(inicio_str, "%H:%M").time()
    if entrada_dt > hora_inicio:
        atraso_min = (
            datetime.combine(fecha_dt, entrada_dt)
            - datetime.combine(fecha_dt, hora_inicio)
        ).total_seconds() / 60.0
        return int(atraso_min)

    return 0


def calcular_horas_extras(entrada, salida, fecha, turno, descripcion):
    """Retorna (horas_50, horas_25). Reglas básicas + límites 21:00 y 07:00."""
    if (
        not entrada
        or not salida
        or str(entrada).strip() in ("-", "", "None")
        or str(salida).strip() in ("-", "", "None")
    ):
        return 0, 0

    desc = str(descripcion or "").lower()
    if "ausente" in desc or "libre" in desc:
        return 0, 0

    fecha_dt = normalizar_fecha(fecha)
    if not fecha_dt:
        return 0, 0

    # Parseo de horas
    fmt_ok = None
    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            entrada_dt = datetime.strptime(str(entrada), fmt)
            salida_dt = datetime.strptime(str(salida), fmt)
            fmt_ok = True
            break
        except Exception:
            pass
    if not fmt_ok:
        return 0, 0

    # Salidas después de medianoche
    if salida_dt < entrada_dt:
        salida_dt += pd.Timedelta(days=1)

    dia_semana = fecha_dt.weekday()
    inicio_str, fin_str = obtener_horario_turno(turno, dia_semana)

    # Sin turno definido → si el bloque es > 30 min, todo como 25%
    if not inicio_str or not fin_str:
        total_min = (salida_dt - entrada_dt).total_seconds() / 60.0
        return (total_min / 60.0, 0) if total_min > 30 else (0, 0)

    hora_inicio = datetime.combine(fecha_dt, datetime.strptime(inicio_str, "%H:%M").time())
    hora_fin = datetime.combine(fecha_dt, datetime.strptime(fin_str, "%H:%M").time())

    minutos_50 = 0.0
    minutos_25 = 0.0

    # Antes del inicio
    if entrada_dt < hora_inicio:
        # Antes de 07:00 es 50%, de 07:00 a inicio es 25%
        limite_7 = datetime.combine(fecha_dt, time(7, 0))
        if entrada_dt < limite_7:
            minutos_50 += (min(salida_dt, hora_inicio) - entrada_dt).total_seconds() / 60.0
        else:
            minutos_25 += (min(salida_dt, hora_inicio) - entrada_dt).total_seconds() / 60.0

    # Después del fin
    if salida_dt > hora_fin:
        limite_21 = datetime.combine(fecha_dt, time(21, 0))
        if salida_dt > limite_21:
            minutos_25 += max(0.0, (limite_21 - hora_fin).total_seconds() / 60.0)
            minutos_50 += (salida_dt - limite_21).total_seconds() / 60.0
        else:
            minutos_25 += (salida_dt - hora_fin).total_seconds() / 60.0

    horas_50 = round(minutos_50 / 60.0, 2) if minutos_50 > 30 else 0
    horas_25 = round(minutos_25 / 60.0, 2) if minutos_25 > 30 else 0
    return horas_50, horas_25


# ====================================
# Lectura y normalización desde HTML
# ====================================

def _decode_bytes(b):
    """Intenta decodificar HTML con varios encodings comunes."""
    for enc in ("utf-8", "latin-1", "cp1252"):
        try:
            return b.decode(enc)
        except Exception:
            pass
    # Último recurso: errores 'replace'
    return b.decode("utf-8", errors="replace")


def _es_html(bytes_head: bytes) -> bool:
    """Heurística rápida: ¿empieza como HTML?"""
    h = bytes_head.lstrip().lower()
    return h.startswith(b"<!doctype html") or h.startswith(b"<html") or b"<table" in h[:400]


def _extraer_meta_html(soup: BeautifulSoup):
    """
    Busca cabeceras tipo:
      Funcionario : X
      N° Rut : Y
      Unidad/Organigrama : Z
      Tipo de Turno : T
      Periodo : P
    en celdas/filas iniciales.
    """
    meta = {"Funcionario": "", "Rut": "", "Organigrama": "", "Turno": "", "Periodo": ""}

    # Texto global rápido
    texto = soup.get_text(" ").strip()
    patrones = {
        "Funcionario": r"Funcionario\s*:?\s*(.+?)\s{2,}",
        "Rut": r"Rut\s*:?\s*([^\s]+)",
        "Organigrama": r"(Organigrama|Unidad|Unidad/Organigrama)\s*:?\s*(.+?)\s{2,}",
        "Turno": r"(Tipo\s*de\s*Turno|Turno)\s*:?\s*(.+?)\s{2,}",
        "Periodo": r"Periodo\s*:?\s*(.+?)\s{2,}",
    }
    for k, pat in patrones.items():
        m = re.search(pat, texto, flags=re.IGNORECASE)
        if m:
            meta[k] = m.group(m.lastindex or 1).strip(" :")

    # Si todavía falta algo, recorremos filas (tr/td) arriba
    if not meta["Funcionario"] or not meta["Periodo"]:
        for tr in soup.find_all("tr")[:30]:
            celdas = [c.get_text(strip=True) for c in tr.find_all(["td", "th"])]
            if len(celdas) >= 2:
                key = celdas[0].lower()
                val = celdas[1]
                if "funcionario" in key and not meta["Funcionario"]:
                    meta["Funcionario"] = val
                elif "rut" in key and not meta["Rut"]:
                    meta["Rut"] = val
                elif ("organigrama" in key or "unidad" in key) and not meta["Organigrama"]:
                    meta["Organigrama"] = val
                elif "turno" in key and not meta["Turno"]:
                    meta["Turno"] = val
                elif "periodo" in key and not meta["Periodo"]:
                    meta["Periodo"] = val

    return meta


def _leer_tablas_html(bhtml: bytes):
    """Devuelve lista de DataFrames de tablas HTML (usa pandas.read_html)."""
    # pandas necesita string, probamos varios encodings
    html_txt = _decode_bytes(bhtml)
    # parse con bs4/lxml para limpiar
    soup = BeautifulSoup(html_txt, "lxml")
    # pandas read_html sobre el string limpio
    df_list = pd.read_html(str(soup), flavor="lxml")
    return soup, df_list


def _procesar_detalle_desde_tabla(df, meta, data_detalle, data_resumen):
    """
    Recibe un DF tipo detalle y acumula en data_detalle/data_resumen.
    Se buscan columnas por nombres aproximados.
    """
    # Normalizar nombres
    cols = {c: str(c).strip().lower() for c in df.columns}
    inv = {v: k for k, v in cols.items()}

    # Mapeos heurísticos
    col_dia = inv.get("dia")
    col_fecha = inv.get("fecha")
    col_entrada = inv.get("entrada")
    col_salida = inv.get("salida")
    col_desc = None
    for label in ("descripción", "descripcion", "obs", "observación", "observacion"):
        if inv.get(label):
            col_desc = inv[label]
            break

    if not (col_fecha and col_entrada and col_salida):
        # No parece detalle
        return

    for _, row in df.iterrows():
        dia_text = str(row[col_dia]) if col_dia else ""
        if str(dia_text).strip().lower() in ("totales", "none"):
            continue

        fecha = row[col_fecha]
        entrada = row[col_entrada]
        salida  = row[col_salida]
        desc    = row[col_desc] if col_desc else ""

        atraso_min = calcular_atraso(entrada, fecha, meta.get("Turno", ""))
        h50, h25   = calcular_horas_extras(entrada, salida, fecha, meta.get("Turno", ""), desc)

        data_detalle.append([
            meta.get("Funcionario",""),
            meta.get("Rut",""),
            meta.get("Organigrama",""),
            meta.get("Turno",""),
            meta.get("Periodo",""),
            fecha, entrada, salida,
            minutos_a_hhmm(atraso_min),
            convertir_a_hhmm(h50),
            convertir_a_hhmm(h25),
            desc,
        ])

        f = meta.get("Funcionario","")
        if f not in data_resumen:
            data_resumen[f] = {
                "Rut": meta.get("Rut",""),
                "Organigrama": meta.get("Organigrama",""),
                "Turno": meta.get("Turno",""),
                "Periodo": meta.get("Periodo",""),
                "Total 50%": 0.0,
                "Total 25%": 0.0,
                "Total Atraso": 0,
            }
        data_resumen[f]["Total 50%"] += h50
        data_resumen[f]["Total 25%"] += h25
        data_resumen[f]["Total Atraso"] += atraso_min


def _procesar_html(bytes_file: bytes) -> BytesIO:
    """
    Procesa el HTML exportado como .xls del Reloj Control.
    Retorna BytesIO con un .xlsx que contiene:
      - 'Detalle Diario'
      - 'Resumen'
    """
    soup, tablas = _leer_tablas_html(bytes_file)
    meta = _extraer_meta_html(soup)

    data_detalle = []
    data_resumen = {}

    # Heurística: la tabla con columnas 'Fecha / Entrada / Salida' es el detalle.
    for df in tablas:
        _procesar_detalle_desde_tabla(df, meta, data_detalle, data_resumen)

    # Construcción de DataFrames finales
    df_detalle = pd.DataFrame(data_detalle, columns=[
        "Funcionario", "Rut", "Organigrama", "Turno", "Periodo",
        "Fecha", "Entrada", "Salida", "Atraso (hh:mm)", "50%", "25%", "Descripción"
    ])

    df_resumen = pd.DataFrame([
        [
            f,
            data_resumen[f]["Rut"],
            data_resumen[f]["Organigrama"],
            data_resumen[f]["Turno"],
            data_resumen[f]["Periodo"],
            convertir_a_hhmm(data_resumen[f]["Total 50%"]),
            convertir_a_hhmm(data_resumen[f]["Total 25%"]),
            minutos_a_hhmm(data_resumen[f]["Total Atraso"]),
            convertir_a_hhmm(data_resumen[f]["Total 50%"] + data_resumen[f]["Total 25%"]),
        ]
        for f in data_resumen
    ], columns=[
        "Funcionario", "Rut", "Organigrama", "Turno", "Periodo",
        "Total 50%", "Total 25%", "Total Atraso", "Total Horas"
    ])

    # Salida a Excel con colores
    return _emitir_excel_coloreado(df_detalle, df_resumen)


# =============================
# Lectura “real” con openpyxl
# =============================

def _leer_xlsx_con_openpyxl(file_stream) -> BytesIO:
    """Parser para verdaderos XLSX (no HTML). Mantiene compatibilidad con la lógica previa."""
    wb = load_workbook(filename=file_stream)
    sh = wb.active

    data_detalle = []
    data_resumen = {}

    fila = 1
    max_row = sh.max_row
    meta = {"Funcionario": "", "Rut": "", "Organigrama": "", "Turno": "", "Periodo": ""}

    while fila <= max_row:
        celda = str(sh.cell(row=fila, column=1).value).strip().lower()

        if celda.startswith("funcionario"):
            meta["Funcionario"]  = str(sh.cell(row=fila + 0, column=2).value or "").strip(": ")
            meta["Rut"]          = str(sh.cell(row=fila + 1, column=2).value or "").strip(": ")
            meta["Organigrama"]  = str(sh.cell(row=fila + 2, column=2).value or "").strip(": ")
            meta["Turno"]        = str(sh.cell(row=fila + 3, column=2).value or "").strip(": ")
            # >>>> LÍNEA CORREGIDA (antes tenía un tab entre row y fila) <<<<
            meta["Periodo"]      = str(sh.cell(row=fila + 4, column=2).value or "").strip(": ")
            fila += 6
            continue

        if celda == "dia":
            fila += 1
            while fila <= max_row and sh.cell(row=fila, column=1).value:
                dia_text = str(sh.cell(row=fila, column=1).value).strip().lower()
                if dia_text == "totales":
                    fila += 1
                    continue
                if dia_text.startswith("funcionario") or dia_text == "none":
                    break

                fecha      = sh.cell(row=fila, column=2).value
                entrada    = sh.cell(row=fila, column=3).value
                salida     = sh.cell(row=fila, column=4).value
                descripcion= str(sh.cell(row=fila, column=6).value or "").strip()

                atraso_min = calcular_atraso(entrada, fecha, meta["Turno"])
                h50, h25   = calcular_horas_extras(entrada, salida, fecha, meta["Turno"], descripcion)

                data_detalle.append([
                    meta["Funcionario"], meta["Rut"], meta["Organigrama"], meta["Turno"], meta["Periodo"],
                    fecha, entrada, salida, minutos_a_hhmm(atraso_min),
                    convertir_a_hhmm(h50), convertir_a_hhmm(h25), descripcion
                ])

                f = meta["Funcionario"]
                if f not in data_resumen:
                    data_resumen[f] = {
                        "Rut": meta["Rut"], "Organigrama": meta["Organigrama"], "Turno": meta["Turno"], "Periodo": meta["Periodo"],
                        "Total 50%": 0.0, "Total 25%": 0.0, "Total Atraso": 0
                    }
                data_resumen[f]["Total 50%"] += h50
                data_resumen[f]["Total 25%"] += h25
                data_resumen[f]["Total Atraso"] += atraso_min
                fila += 1
        else:
            fila += 1

    df_detalle = pd.DataFrame(data_detalle, columns=[
        "Funcionario", "Rut", "Organigrama", "Turno", "Periodo",
        "Fecha", "Entrada", "Salida", "Atraso (hh:mm)", "50%", "25%", "Descripción"
    ])

    df_resumen = pd.DataFrame([
        [
            f,
            data_resumen[f]["Rut"],
            data_resumen[f]["Organigrama"],
            data_resumen[f]["Turno"],
            data_resumen[f]["Periodo"],
            convertir_a_hhmm(data_resumen[f]["Total 50%"]),
            convertir_a_hhmm(data_resumen[f]["Total 25%"]),
            minutos_a_hhmm(data_resumen[f]["Total Atraso"]),
            convertir_a_hhmm(data_resumen[f]["Total 50%"] + data_resumen[f]["Total 25%"]),
        ]
        for f in data_resumen
    ], columns=[
        "Funcionario", "Rut", "Organigrama", "Turno", "Periodo",
        "Total 50%", "Total 25%", "Total Atraso", "Total Horas"
    ])

    return _emitir_excel_coloreado(df_detalle, df_resumen)


# ===========================
# Emisión Excel + coloreado
# ===========================

def _emitir_excel_coloreado(df_detalle: pd.DataFrame, df_resumen: pd.DataFrame) -> BytesIO:
    """Escribe ambos DF a un .xlsx y colorea celdas según reglas."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_detalle.to_excel(writer, sheet_name="Detalle Diario", index=False)
        df_resumen.to_excel(writer, sheet_name="Resumen", index=False)

    # Coloreado
    output.seek(0)
    wb_final = load_workbook(output)
    ws = wb_final["Detalle Diario"]

    fill_rojo     = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")  # Ausente
    fill_amarillo = PatternFill(start_color="FFFACD", end_color="FFFACD", fill_type="solid")  # Faltas

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=12):
        descripcion = str(row[11].value or "")
        entrada     = str(row[6].value or "").strip()
        salida      = str(row[7].value or "").strip()

        if "Ausente" in descripcion:
            for cell in row:
                cell.fill = fill_rojo
        elif "Falta Entrada" in descripcion or "Falta Salida" in descripcion or entrada in ("", "-") or salida in ("", "-"):
            for cell in row:
                cell.fill = fill_amarillo

    final_output = BytesIO()
    wb_final.save(final_output)
    final_output.seek(0)
    return final_output


# ===========================
# APIs que usa app.py
# ===========================

def detectar_html_y_procesar(file_stream) -> BytesIO:
    """
    Para compatibilidad: detecta si es HTML (.xls falso) y procesa;
    si no, intenta como XLSX real.
    """
    # Leemos bytes sin consumir el stream original (Render/Flask guarda en memoria)
    data = file_stream.read()
    # Volvemos a posicionar por si alguien reusa el stream
    try:
        file_stream.seek(0)
    except Exception:
        pass

    head = data[:500].lstrip().lower()
    if _es_html(head):
        # HTML exportado con extensión .xls
        return _procesar_html(data)

    # No parece HTML → intentamos XLSX
    return _leer_xlsx_con_openpyxl(BytesIO(data))


def procesar_excel(file_stream) -> BytesIO:
    """Alias (app.py puede llamar a cualquiera)."""
    return detectar_html_y_procesar(file_stream)
