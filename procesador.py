import pandas as pd
import re
from datetime import datetime, time
from openpyxl.styles import PatternFill
from openpyxl import load_workbook

FERIADOS_2025 = [
    "2025-01-01", "2025-04-18", "2025-04-19", "2025-05-01",
    "2025-05-21", "2025-06-29", "2025-07-16", "2025-08-15",
    "2025-09-18", "2025-09-19", "2025-10-12", "2025-10-31",
    "2025-11-01", "2025-12-08", "2025-12-25"
]
FERIADOS_2025 = [datetime.strptime(d, "%Y-%m-%d").date() for d in FERIADOS_2025]

def normalizar_fecha(fecha):
    if isinstance(fecha, datetime):
        return fecha.date()
    try:
        return datetime.strptime(str(fecha).strip(), "%d-%m-%Y").date()
    except:
        try:
            return datetime.strptime(str(fecha).strip(), "%d/%m/%Y").date()
        except:
            return None

def convertir_a_hhmm(horas):
    minutos = int(round(horas * 60))
    return f"{minutos // 60:02}:{minutos % 60:02}"

def minutos_a_hhmm(minutos):
    return f"{minutos // 60:02}:{minutos % 60:02}"

def obtener_horario_turno(turno, dia_semana):
    horas = re.findall(r"\d{1,2}:\d{2}", turno)
    if len(horas) < 2:
        return None, None
    if dia_semana in [0, 1, 2, 3]:  # Lu-Ju
        return horas[0], horas[1]
    elif dia_semana == 4 and len(horas) >= 4:  # Vi
        return horas[2], horas[3]
    return None, None

def calcular_atraso(entrada, fecha, turno):
    if not entrada or entrada == "-":
        return 0
    fecha_dt = normalizar_fecha(fecha)
    if not fecha_dt:
        return 0
    try:
        entrada_dt = datetime.strptime(str(entrada), "%H:%M:%S").time()
    except:
        return 0

    dia_semana = fecha_dt.weekday()
    inicio_str, _ = obtener_horario_turno(turno, dia_semana)
    if not inicio_str:
        return 0

    hora_inicio = datetime.strptime(inicio_str, "%H:%M").time()
    if entrada_dt > hora_inicio:
        atraso_min = (datetime.combine(fecha_dt, entrada_dt) - datetime.combine(fecha_dt, hora_inicio)).total_seconds() / 60
        return int(atraso_min)
    return 0

def calcular_horas_extras(entrada, salida, fecha, turno, descripcion):
    if not entrada or not salida or entrada == "-" or salida == "-" or "ausente" in descripcion.lower() or "libre" in descripcion.lower():
        return 0, 0
    fecha_dt = normalizar_fecha(fecha)
    if not fecha_dt:
        return 0, 0
    try:
        entrada_dt = datetime.strptime(str(entrada), "%H:%M:%S")
        salida_dt = datetime.strptime(str(salida), "%H:%M:%S")
    except:
        return 0, 0
    if salida_dt < entrada_dt:
        salida_dt += pd.Timedelta(days=1)

    dia_semana = fecha_dt.weekday()
    inicio_str, fin_str = obtener_horario_turno(turno, dia_semana)
    if not inicio_str or not fin_str:
        total_min = (salida_dt - entrada_dt).total_seconds() / 60
        return (total_min / 60, 0) if total_min > 30 else (0, 0)

    hora_inicio = datetime.combine(fecha_dt, datetime.strptime(inicio_str, "%H:%M").time())
    hora_fin = datetime.combine(fecha_dt, datetime.strptime(fin_str, "%H:%M").time())

    minutos_50 = 0
    minutos_25 = 0

    if entrada_dt < hora_inicio:
        if entrada_dt.time() < time(7, 0):
            minutos_50 += (min(salida_dt, hora_inicio) - entrada_dt).total_seconds() / 60
        else:
            minutos_25 += (min(salida_dt, hora_inicio) - entrada_dt).total_seconds() / 60

    if salida_dt > hora_fin:
        if salida_dt > datetime.combine(fecha_dt, time(21, 0)):
            minutos_25 += max(0, (datetime.combine(fecha_dt, time(21, 0)) - hora_fin).total_seconds() / 60)
            minutos_50 += (salida_dt - datetime.combine(fecha_dt, time(21, 0))).total_seconds() / 60
        else:
            minutos_25 += (salida_dt - hora_fin).total_seconds() / 60

    horas_50 = round(minutos_50 / 60, 2) if minutos_50 > 30 else 0
    horas_25 = round(minutos_25 / 60, 2) if minutos_25 > 30 else 0
    return horas_50, horas_25

def procesar_excel(entrada, salida):
    df_raw = pd.read_excel(entrada, engine="xlrd")  # Para archivos .xls

    data_detalle = []
    data_resumen = {}
    funcionario = rut = organigrama = turno = periodo = ""

    i = 0
    while i < len(df_raw):
        celda = str(df_raw.iloc[i, 0]).strip().lower()
        if celda.startswith("funcionario"):
            funcionario = str(df_raw.iloc[i, 1]).strip(": ")
            rut = str(df_raw.iloc[i + 1, 1]).strip(": ")
            organigrama = str(df_raw.iloc[i + 2, 1]).strip(": ")
            turno = str(df_raw.iloc[i + 3, 1]).strip(": ")
            periodo = str(df_raw.iloc[i + 4, 1]).strip(": ")
            i += 6
            continue

        if celda == "dia":
            i += 1
            while i < len(df_raw) and str(df_raw.iloc[i, 0]).strip():
                dia_text = str(df_raw.iloc[i, 0]).strip().lower()
                if dia_text == "totales" or dia_text.startswith("funcionario"):
                    break

                fecha = df_raw.iloc[i, 1]
                entrada_val = df_raw.iloc[i, 2]
                salida_val = df_raw.iloc[i, 3]
                descripcion = str(df_raw.iloc[i, 5])

                atraso = calcular_atraso(entrada_val, fecha, turno)
                h50, h25 = calcular_horas_extras(entrada_val, salida_val, fecha, turno, descripcion)

                data_detalle.append([
                    funcionario, rut, organigrama, turno, periodo,
                    fecha, entrada_val, salida_val,
                    minutos_a_hhmm(atraso), convertir_a_hhmm(h50), convertir_a_hhmm(h25), descripcion
                ])

                if funcionario not in data_resumen:
                    data_resumen[funcionario] = {"Rut": rut, "Organigrama": organigrama, "Turno": turno, "Periodo": periodo,
                                                 "Total 50%": 0, "Total 25%": 0, "Total Atraso": 0}
                data_resumen[funcionario]["Total 50%"] += h50
                data_resumen[funcionario]["Total 25%"] += h25
                data_resumen[funcionario]["Total Atraso"] += atraso
                i += 1
        else:
            i += 1

    df_detalle = pd.DataFrame(data_detalle, columns=[
        "Funcionario", "Rut", "Organigrama", "Turno", "Periodo",
        "Fecha", "Entrada", "Salida", "Atraso (hh:mm)", "50%", "25%", "Descripción"
    ])

    df_resumen = pd.DataFrame([
        [
            f, data_resumen[f]["Rut"], data_resumen[f]["Organigrama"], data_resumen[f]["Turno"], data_resumen[f]["Periodo"],
            convertir_a_hhmm(data_resumen[f]["Total 50%"]),
            convertir_a_hhmm(data_resumen[f]["Total 25%"]),
            minutos_a_hhmm(data_resumen[f]["Total Atraso"]),
            convertir_a_hhmm(data_resumen[f]["Total 50%"] + data_resumen[f]["Total 25%"])
        ]
        for f in data_resumen
    ], columns=["Funcionario", "Rut", "Organigrama", "Turno", "Periodo", "Total 50%", "Total 25%", "Total Atraso", "Total Horas"])

    with pd.ExcelWriter(salida, engine="openpyxl") as writer:
        df_detalle.to_excel(writer, sheet_name="Detalle Diario", index=False)
        df_resumen.to_excel(writer, sheet_name="Resumen", index=False)

    # Colorear celas
    wb = load_workbook(salida)
    ws = wb["Detalle Diario"]
    fill_rojo = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    fill_amarillo = PatternFill(start_color="FFFACD", end_color="FFFACD", fill_type="solid")

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=12):
        descripcion = str(row[11].value).strip().lower()
        entrada = str(row[6].value).strip()
        salida = str(row[7].value).strip()

        if "ausente" in descripcion:
            for cell in row:
                cell.fill = fill_rojo
        elif "falta entrada" in descripcion or "falta salida" in descripcion or entrada == "-" or salida == "-":
            for cell in row:
                cell.fill = fill_amarillo

    wb.save(salida)
