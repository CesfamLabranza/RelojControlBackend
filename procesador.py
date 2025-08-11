import re
from io import BytesIO
from datetime import datetime, time

import pandas as pd
import xlrd  # ✅ para leer .xls
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

FERIADOS_2025 = [
    "2025-01-01", "2025-04-18", "2025-04-19", "2025-05-01",
    "2025-05-21", "2025-06-29", "2025-07-16", "2025-08-15",
    "2025-09-18", "2025-09-19", "2025-10-12", "2025-10-31",
    "2025-11-01", "2025-12-08", "2025-12-25"
]
FERIADOS_2025 = [datetime.strptime(d, "%Y-%m-%d").date() for d in FERIADOS_2025]

# ----------------- helpers -----------------

def normalizar_fecha_xls(valor, datemode):
    """
    Convierte un valor de fecha proveniente de .xls a date.
    Acepta datetime, string 'dd-mm-aaaa' / 'dd/mm/aaaa' o numérico Excel.
    """
    if isinstance(valor, datetime):
        return valor.date()

    # numérico Excel
    if isinstance(valor, (int, float)):
        try:
            return xlrd.xldate_as_datetime(valor, datemode).date()
        except Exception:
            pass

    # string
    try:
        s = str(valor).strip()
        if "-" in s:
            return datetime.strptime(s, "%d-%m-%Y").date()
        if "/" in s:
            return datetime.strptime(s, "%d/%m/%Y").date()
    except Exception:
        return None

    return None

def convertir_a_hhmm(horas):
    minutos = int(round(horas * 60))
    return f"{minutos // 60:02}:{minutos % 60:02}"

def minutos_a_hhmm(minutos):
    return f"{minutos // 60:02}:{minutos % 60:02}"

def obtener_horario_turno(turno, dia_semana):
    horas = re.findall(r"\d{1,2}:\d{2}", turno or "")
    if len(horas) < 2:
        return None, None
    if dia_semana in [0, 1, 2, 3]:  # Lu-Ju
        return horas[0], horas[1]
    if dia_semana == 4 and len(horas) >= 4:  # Vi
        return horas[2], horas[3]
    return None, None

def parse_hora(h):
    """Devuelve time o None a partir de 'HH:MM:SS' / 'HH:MM' o '-'."""
    if not h or str(h).strip() == "-" or str(h).strip().lower() == "none":
        return None
    s = str(h).strip()
    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            return datetime.strptime(s, fmt).time()
        except Exception:
            pass
    return None

def calcular_atraso(entrada, fecha_date, turno):
    if not entrada:
        return 0
    dia_semana = fecha_date.weekday()
    inicio_str, _ = obtener_horario_turno(turno, dia_semana)
    if not inicio_str:
        return 0
    hora_inicio = datetime.strptime(inicio_str, "%H:%M").time()
    if entrada > hora_inicio:
        return int(
            (datetime.combine(fecha_date, entrada) - datetime.combine(fecha_date, hora_inicio)).total_seconds() / 60
        )
    return 0

def calcular_horas_extras(entrada, salida, fecha_date, turno, descripcion):
    if not entrada or not salida:
        return 0, 0
    desc = (descripcion or "").lower()
    if "ausente" in desc or "libre" in desc:
        return 0, 0

    ent_dt = datetime.combine(fecha_date, entrada)
    sal_dt = datetime.combine(fecha_date, salida)
    if sal_dt < ent_dt:
        # pasó medianoche
        sal_dt = sal_dt.replace(day=sal_dt.day + 1)

    dia_semana = fecha_date.weekday()
    inicio_str, fin_str = obtener_horario_turno(turno, dia_semana)

    # sin horario: todo lo trabajado cuenta como 25% si supera 30 min
    if not inicio_str or not fin_str:
        total_min = (sal_dt - ent_dt).total_seconds() / 60
        return (total_min / 60, 0) if total_min > 30 else (0, 0)

    hora_inicio = datetime.combine(fecha_date, datetime.strptime(inicio_str, "%H:%M").time())
    hora_fin = datetime.combine(fecha_date, datetime.strptime(fin_str, "%H:%M").time())

    minutos_50 = 0
    minutos_25 = 0

    if ent_dt < hora_inicio:
        # antes de hora de inicio
        if ent_dt.time() < time(7, 0):
            minutos_50 += (min(sal_dt, hora_inicio) - ent_dt).total_seconds() / 60
        else:
            minutos_25 += (min(sal_dt, hora_inicio) - ent_dt).total_seconds() / 60

    if sal_dt > hora_fin:
        if sal_dt.time() > time(21, 0):
            minutos_25 += max(0, (datetime.combine(fecha_date, time(21, 0)) - hora_fin).total_seconds() / 60)
            minutos_50 += (sal_dt - datetime.combine(fecha_date, time(21, 0))).total_seconds() / 60
        else:
            minutos_25 += (sal_dt - hora_fin).total_seconds() / 60

    horas_50 = round(minutos_50 / 60, 2) if minutos_50 > 30 else 0
    horas_25 = round(minutos_25 / 60, 2) if minutos_25 > 30 else 0
    return horas_50, horas_25

# ----------------- procesador principal (.xls -> .xlsx) -----------------

def procesar_excel(file_stream: BytesIO) -> BytesIO:
    # Abrimos .xls con xlrd
    book = xlrd.open_workbook(file_contents=file_stream.read())
    sheet = book.sheet_by_index(0)

    def get(r, c):
        """1-based como antes: get(fila, columna)."""
        try:
            return sheet.cell_value(r - 1, c - 1)
        except Exception:
            return None

    data_detalle = []
    data_resumen = {}
    fila = 1
    max_row = sheet.nrows

    funcionario = rut = organigrama = turno = periodo = ""

    while fila <= max_row:
        celda = str(get(fila, 1)).strip().lower()
        if celda.startswith("funcionario"):
            funcionario = str(get(fila, 2)).strip(": ")
            rut = str(get(fila + 1, 2)).strip(": ")
            organigrama = str(get(fila + 2, 2)).strip(": ")
            turno = str(get(fila + 3, 2)).strip(": ")
            periodo = str(get(fila + 4, 2)).strip(": ")
            fila += 6
            continue

        if celda == "dia":
            fila += 1
            while fila <= max_row and str(get(fila, 1)).strip():
                dia_text = str(get(fila, 1)).strip().lower()
                if dia_text == "totales":
                    fila += 1
                    continue
                if dia_text.startswith("funcionario") or dia_text == "none":
                    break

                fecha_raw = get(fila, 2)
                entrada_raw = get(fila, 3)
                salida_raw = get(fila, 4)
                descripcion = str(get(fila, 6) or "").strip()

                # fecha como date
                fecha_date = normalizar_fecha_xls(fecha_raw, book.datemode)
                # horas como time
                entrada_t = parse_hora(entrada_raw)
                salida_t  = parse_hora(salida_raw)

                atraso_min = 0
                horas_50 = horas_25 = 0

                if fecha_date:
                    atraso_min = calcular_atraso(entrada_t, fecha_date, turno)
                    h50, h25 = calcular_horas_extras(entrada_t, salida_t, fecha_date, turno, descripcion)
                    horas_50, horas_25 = h50, h25

                data_detalle.append([
                    funcionario, rut, organigrama, turno, periodo,
                    fecha_date.strftime("%d-%m-%Y") if fecha_date else "",
                    entrada_raw if entrada_t else "-",
                    salida_raw if salida_t else "-",
                    minutos_a_hhmm(atraso_min),
                    convertir_a_hhmm(horas_50),
                    convertir_a_hhmm(horas_25),
                    descripcion
                ])

                if funcionario not in data_resumen:
                    data_resumen[funcionario] = {
                        "Rut": rut, "Organigrama": organigrama, "Turno": turno, "Periodo": periodo,
                        "Total 50%": 0, "Total 25%": 0, "Total Atraso": 0
                    }
                data_resumen[funcionario]["Total 50%"] += horas_50
                data_resumen[funcionario]["Total 25%"] += horas_25
                data_resumen[funcionario]["Total Atraso"] += atraso_min

                fila += 1
        else:
            fila += 1

    df_detalle = pd.DataFrame(data_detalle, columns=[
        "Funcionario", "Rut", "Organigrama", "Turno", "Periodo",
        "Fecha", "Entrada", "Salida", "Atraso (hh:mm)", "50%", "25%", "Descripción"
    ])

    df_resumen = pd.DataFrame([
        [
            f, data_resumen[f]["Rut"], data_resumen[f]["Organigrama"], data_resumen[f]["Turno"],
            data_resumen[f]["Periodo"],
            convertir_a_hhmm(data_resumen[f]["Total 50%"]),
            convertir_a_hhmm(data_resumen[f]["Total 25%"]),
            minutos_a_hhmm(data_resumen[f]["Total Atraso"]),
            convertir_a_hhmm(data_resumen[f]["Total 50%"] + data_resumen[f]["Total 25%"])
        ]
        for f in data_resumen
    ], columns=["Funcionario", "Rut", "Organigrama", "Turno", "Periodo",
                "Total 50%", "Total 25%", "Total Atraso", "Total Horas"])

    # --- Generar XLSX de salida con formato ---
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_detalle.to_excel(writer, sheet_name="Detalle Diario", index=False)
        df_resumen.to_excel(writer, sheet_name="Resumen", index=False)

    output.seek(0)
    wb_final = load_workbook(output)
    ws = wb_final["Detalle Diario"]

    fill_rojo = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    fill_amarillo = PatternFill(start_color="FFFACD", end_color="FFFACD", fill_type="solid")

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=12):
        descripcion = str(row[11].value or "")
        entrada = str(row[6].value or "").strip()
        salida = str(row[7].value or "").strip()
        if "Ausente" in descripcion:
            for cell in row:
                cell.fill = fill_rojo
        elif "Falta Entrada" in descripcion or "Falta Salida" in descripcion or entrada == "-" or salida == "-":
            for cell in row:
                cell.fill = fill_amarillo

    final_output = BytesIO()
    wb_final.save(final_output)
    final_output.seek(0)
    return final_output
