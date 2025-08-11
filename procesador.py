# procesador.py
import re
from io import BytesIO, StringIO
from datetime import datetime, time

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# ============================================================
# Utilidades de compatibilidad de entrada (bytes vs file-like)
# ============================================================

def _ensure_buffer(stream_or_bytes):
    """
    Acepta bytes/bytearray o un objeto con .read() y devuelve
    SIEMPRE algo con .read() (BytesIO si era bytes).
    """
    if hasattr(stream_or_bytes, "read"):
        return stream_or_bytes
    return BytesIO(stream_or_bytes)

# ============================================================
# Funciones auxiliares de negocio (idénticas/compatibles)
# ============================================================

def normalizar_fecha(fecha):
    if isinstance(fecha, datetime):
        return fecha.date()
    try:
        fecha_str = str(fecha).strip()
        if "-" in fecha_str:
            # dd-mm-YYYY
            return datetime.strptime(fecha_str, "%d-%m-%Y").date()
        elif "/" in fecha_str:
            # dd/mm/YYYY
            return datetime.strptime(fecha_str, "%d/%m/%Y").date()
    except Exception:
        return None
    return None

def convertir_a_hhmm(horas):
    try:
        minutos = int(round(float(horas) * 60))
        return f"{minutos // 60:02}:{minutos % 60:02}"
    except Exception:
        return "00:00"

def minutos_a_hhmm(minutos):
    try:
        m = int(round(float(minutos)))
        return f"{m // 60:02}:{m % 60:02}"
    except Exception:
        return "00:00"

def obtener_horario_turno(turno, dia_semana):
    """
    Busca horas en el string de turno. Formato típico:
    'Lu-Ju 08:00-17:00 / Vi 08:00-16:00' (ejemplo).
    Extrae pares de HH:MM.
    """
    horas = re.findall(r"\d{1,2}:\d{2}", str(turno or ""))
    if len(horas) < 2:
        return None, None

    # Simplificación:
    #   Lu-Ju -> horas[0]-horas[1]
    #   Vi    -> horas[2]-horas[3] si existen
    if dia_semana in [0, 1, 2, 3]:  # L a J
        return horas[0], horas[1]
    elif dia_semana == 4 and len(horas) >= 4:  # V
        return horas[2], horas[3]
    return None, None

def _parse_time_hms(value):
    """
    Convierte 'HH:MM:SS' a datetime.time. Devuelve None si no se puede.
    """
    if not value or value == "-":
        return None
    try:
        return datetime.strptime(str(value), "%H:%M:%S").time()
    except Exception:
        # aceptemos HH:MM también si viene así
        try:
            return datetime.strptime(str(value), "%H:%M").time()
        except Exception:
            return None

def calcular_atraso(entrada, fecha, turno):
    """
    Minutos de atraso: diferencia entre entrada y hora de inicio de turno.
    """
    ent = _parse_time_hms(entrada)
    if ent is None:
        return 0

    fecha_dt = normalizar_fecha(fecha)
    if not fecha_dt:
        return 0

    dia_semana = fecha_dt.weekday()
    inicio_str, _ = obtener_horario_turno(turno, dia_semana)
    if not inicio_str:
        return 0

    try:
        hora_inicio = datetime.strptime(inicio_str, "%H:%M").time()
    except Exception:
        return 0

    if ent > hora_inicio:
        atraso_min = (
            datetime.combine(fecha_dt, ent)
            - datetime.combine(fecha_dt, hora_inicio)
        ).total_seconds() / 60
        return int(atraso_min)
    return 0

def calcular_horas_extras(entrada, salida, fecha, turno, descripcion):
    """
    Cálculo simple de horas extra 50% y 25%:
      - Antes del inicio: <07:00 -> 50%, de 07:00 a inicio -> 25%
      - Después del fin:  fin a 21:00 -> 25%, luego de 21:00 -> 50%
    No contabiliza si 'ausente' o 'libre'.
    """
    if (
        not entrada or not salida or entrada == "-" or salida == "-"
        or "ausente" in str(descripcion).lower() or "libre" in str(descripcion).lower()
    ):
        return 0, 0

    fecha_dt = normalizar_fecha(fecha)
    if not fecha_dt:
        return 0, 0

    ent_t = _parse_time_hms(entrada)
    sal_t = _parse_time_hms(salida)
    if ent_t is None or sal_t is None:
        return 0, 0

    ent_dt = datetime.combine(fecha_dt, ent_t)
    sal_dt = datetime.combine(fecha_dt, sal_t)
    if sal_dt < ent_dt:
        # cruce de medianoche
        sal_dt = sal_dt + pd.Timedelta(days=1)

    dia_semana = fecha_dt.weekday()
    inicio_str, fin_str = obtener_horario_turno(turno, dia_semana)

    # si no hay referencia de turno, contamos todo como tiempo trabajado
    if not inicio_str or not fin_str:
        total_min = (sal_dt - ent_dt).total_seconds() / 60
        return ((total_min / 60.0), 0) if total_min > 30 else (0, 0)

    hora_inicio = datetime.combine(fecha_dt, datetime.strptime(inicio_str, "%H:%M").time())
    hora_fin = datetime.combine(fecha_dt, datetime.strptime(fin_str, "%H:%M").time())

    minutos_50 = 0
    minutos_25 = 0

    # Antes de la hora de inicio
    if ent_dt < hora_inicio:
        if ent_dt.time() < time(7, 0):
            minutos_50 += (min(sal_dt, hora_inicio) - ent_dt).total_seconds() / 60
        else:
            minutos_25 += (min(sal_dt, hora_inicio) - ent_dt).total_seconds() / 60

    # Después de la hora de fin
    if sal_dt > hora_fin:
        if sal_dt > datetime.combine(fecha_dt, time(21, 0)):
            minutos_25 += max(0, (datetime.combine(fecha_dt, time(21, 0)) - hora_fin).total_seconds() / 60)
            minutos_50 += (sal_dt - datetime.combine(fecha_dt, time(21, 0))).total_seconds() / 60
        else:
            minutos_25 += (sal_dt - hora_fin).total_seconds() / 60

    horas_50 = round(minutos_50 / 60.0, 2) if minutos_50 > 30 else 0
    horas_25 = round(minutos_25 / 60.0, 2) if minutos_25 > 30 else 0
    return horas_50, horas_25

# ============================================================
# Procesamiento de EXCEL real (openpyxl) — Detalle + Resumen
# ============================================================

def procesar_excel(file_stream):
    """
    Lee planilla 'real' con openpyxl. Busca bloques:
      Funcionario / Rut / Organigrama / Turno / Periodo
      ...luego la tabla desde "Dia" con columnas fecha, entrada, salida, desc.
    Devuelve BytesIO con .xlsx (Detalle Diario + Resumen) y celdas coloreadas.
    """
    file_stream = _ensure_buffer(file_stream)
    wb = load_workbook(filename=file_stream, data_only=True)
    sheet = wb.active

    data_detalle = []
    data_resumen = {}

    fila = 1
    max_row = sheet.max_row
    meta = {"Funcionario": "", "Rut": "", "Organigrama": "", "Turno": "", "Periodo": ""}

    while fila <= max_row:
        celda = sheet.cell(row=fila, column=1).value
        celda_str = str(celda).strip().lower() if celda is not None else ""

        # encabezado de metadatos
        if celda_str.startswith("funcionario"):
            meta["Funcionario"] = str(sheet.cell(row=fila,     column=2).value or "").strip(": ")
            meta["Rut"]         = str(sheet.cell(row=fila + 1, column=2).value or "").strip(": ")
            meta["Organigrama"] = str(sheet.cell(row=fila + 2, column=2).value or "").strip(": ")
            meta["Turno"]       = str(sheet.cell(row=fila + 3, column=2).value or "").strip(": ")
            meta["Periodo"]     = str(sheet.cell(row=fila + 4, column=2).value or "").strip(": ")
            fila += 6
            continue

        # cabecera de tabla
        if celda_str == "dia":
            fila += 1
            # filas hasta toparse con otra sección
            while fila <= max_row and sheet.cell(row=fila, column=1).value:
                dia_text = str(sheet.cell(row=fila, column=1).value).strip().lower()
                if dia_text == "totales":
                    fila += 1
                    continue
                if dia_text.startswith("funcionario") or dia_text == "none":
                    break

                fecha       = sheet.cell(row=fila, column=2).value
                entrada     = sheet.cell(row=fila, column=3).value
                salida      = sheet.cell(row=fila, column=4).value
                descripcion = str(sheet.cell(row=fila, column=6).value or "").strip()

                atraso_min = calcular_atraso(entrada, fecha, meta["Turno"])
                h50, h25   = calcular_horas_extras(entrada, salida, fecha, meta["Turno"], descripcion)

                data_detalle.append([
                    meta["Funcionario"], meta["Rut"], meta["Organigrama"], meta["Turno"], meta["Periodo"],
                    fecha, entrada, salida,
                    minutos_a_hhmm(atraso_min), convertir_a_hhmm(h50), convertir_a_hhmm(h25), descripcion
                ])

                fkey = meta["Funcionario"] or "SIN NOMBRE"
                if fkey not in data_resumen:
                    data_resumen[fkey] = {
                        "Rut": meta["Rut"], "Organigrama": meta["Organigrama"],
                        "Turno": meta["Turno"], "Periodo": meta["Periodo"],
                        "Total 50%": 0.0, "Total 25%": 0.0, "Total Atraso": 0
                    }
                data_resumen[fkey]["Total 50%"] += h50
                data_resumen[fkey]["Total 25%"] += h25
                data_resumen[fkey]["Total Atraso"] += atraso_min

                fila += 1
        else:
            fila += 1

    # DataFrames
    df_detalle = pd.DataFrame(
        data_detalle,
        columns=["Funcionario", "Rut", "Organigrama", "Turno", "Periodo",
                 "Fecha", "Entrada", "Salida", "Atraso (hh:mm)", "50%", "25%", "Descripción"]
    )

    df_resumen = pd.DataFrame([
        [
            f,
            data_resumen[f]["Rut"], data_resumen[f]["Organigrama"],
            data_resumen[f]["Turno"], data_resumen[f]["Periodo"],
            convertir_a_hhmm(data_resumen[f]["Total 50%"]),
            convertir_a_hhmm(data_resumen[f]["Total 25%"]),
            minutos_a_hhmm(data_resumen[f]["Total Atraso"]),
            convertir_a_hhmm(data_resumen[f]["Total 50%"] + data_resumen[f]["Total 25%"])
        ]
        for f in data_resumen
    ], columns=["Funcionario", "Rut", "Organigrama", "Turno", "Periodo",
                "Total 50%", "Total 25%", "Total Atraso", "Total Horas"])

    # Escribir a XLSX en memoria
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_detalle.to_excel(writer, sheet_name="Detalle Diario", index=False)
        df_resumen.to_excel(writer, sheet_name="Resumen", index=False)

    # Colorear filas en Detalle
    output.seek(0)
    wb_final = load_workbook(output)
    ws = wb_final["Detalle Diario"]

    fill_rojo     = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    fill_amarillo = PatternFill(start_color="FFFACD", end_color="FFFACD", fill_type="solid")

    # Columnas (base 1): Fecha(6) Entrada(7) Salida(8) Descripción(12)
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=12):
        descripcion = str(row[11].value or "")
        entrada_str = str(row[6].value or "").strip()
        salida_str  = str(row[7].value or "").strip()

        if "ausente" in descripcion.lower():
            for c in row:
                c.fill = fill_rojo
        elif ("falta entrada" in descripcion.lower()
              or "falta salida" in descripcion.lower()
              or entrada_str == "-" or salida_str == "-"):
            for c in row:
                c.fill = fill_amarillo

    final_output = BytesIO()
    wb_final.save(final_output)
    final_output.seek(0)
    return final_output

# ============================================================
# Procesamiento de HTML exportado como .xls (fallback)
# ============================================================

def _leer_html_con_encoding(buf):
    """
    Intenta decodificar los bytes del HTML con varios encodings comunes.
    """
    raw = buf.read()
    for enc in ("utf-8", "latin-1", "cp1252"):
        try:
            s = raw.decode(enc)
            return s
        except Exception:
            continue
    # último recurso: reemplazar errores
    return raw.decode("latin-1", errors="replace")

def _procesar_html_xls(buf):
    """
    Lee el HTML exportado por la plataforma (con extensión .xls),
    extrae la tabla más grande y genera:
      - Hoja 'Detalle Diario' con la tabla más grande.
      - Hoja 'Resumen' agregando por una columna clave si existe.
    """
    # reset posición
    buf.seek(0)
    html_text = _leer_html_con_encoding(buf)

    # Buscar tablas
    try:
        tablas = pd.read_html(StringIO(html_text), flavor="lxml")
    except Exception:
        # intentar con html5lib
        tablas = pd.read_html(StringIO(html_text), flavor="bs4")

    if not tablas:
        raise ValueError("No se pudieron leer tablas HTML desde el archivo.")

    # Tomemos la tabla de mayor tamaño como 'detalle'
    tabla = max(tablas, key=lambda t: t.shape[0] * t.shape[1])

    # Normalización suave de nombres de columnas
    tabla.columns = [str(c).strip() for c in tabla.columns]

    # Intentar fechas: si hay columna con 'fecha'
    for col in tabla.columns:
        if "fecha" in col.lower():
            try:
                tabla[col] = pd.to_datetime(tabla[col], dayfirst=True, errors="coerce")
            except Exception:
                pass
            break

    # Construir un resumen simple:
    posibles_persona = [c for c in tabla.columns if any(k in c.lower() for k in ["funcionario", "nombre", "trabajador", "empleado"])]
    resumen = None
    if posibles_persona:
        key = posibles_persona[0]
        resumen = (
            tabla.groupby(key)
                 .size()
                 .reset_index(name="Registros")
                 .sort_values("Registros", ascending=False)
        )
    else:
        resumen = pd.DataFrame({"Resumen": ["Total filas"], "Valor": [len(tabla)]})

    # Salida XLSX
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        tabla.to_excel(writer, sheet_name="Detalle Diario", index=False)
        resumen.to_excel(writer, sheet_name="Resumen", index=False)
    out.seek(0)
    return out

# ============================================================
# Selector de formato: HTML vs EXCEL real
# ============================================================

def detectar_html_y_procesar(stream_or_bytes):
    """
    Si el archivo 'xls' en realidad es HTML, procesamos con pandas.read_html.
    Si no, usamos openpyxl para la planilla 'real'.
    Devuelve BytesIO con el XLSX final.
    """
    buf = _ensure_buffer(stream_or_bytes)

    # inspección de cabecera
    head = buf.read(512)
    buf.seek(0)

    # ¿es HTML?
    es_html = b"<html" in head.lower() or b"<!doctype html" in head.lower()
    if es_html:
        return _procesar_html_xls(buf)

    # si no parece HTML, tratamos como excel real
    return procesar_excel(buf)
