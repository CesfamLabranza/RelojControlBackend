import pandas as pd
from datetime import datetime, time
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill
from io import BytesIO
import re

# ----------------- Utilidades -----------------

def normalizar_fecha(fecha):
    """
    Convierte valores de fecha de .xls a date.
    Acepta datetime, 'dd-mm-aaaa', 'dd/mm/aaaa'
    """
    if isinstance(fecha, datetime):
        return fecha.date()

    if fecha is None:
        return None

    s = str(fecha).strip()
    if not s or s.lower() == "none":
        return None

    for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%d-%m-%y", "%d/%m/%y"):
        try:
            return datetime.strptime(s, fmt).date()
        except Exception:
            pass

    # último intento: pandas puede traer '2025-08-04 00:00:00'
    try:
        return pd.to_datetime(s).date()
    except Exception:
        return None


def minutos_a_hhmm(minutos):
    return f"{int(minutos) // 60:02}:{int(minutos) % 60:02}"


def horas_a_hhmm(horas):
    mins = int(round(horas * 60))
    return minutos_a_hhmm(mins)


def _parse_hora(h):
    """Acepta time, 'HH:MM:SS', 'HH:MM'."""
    if h is None or str(h).strip() == "":
        return None
    if isinstance(h, datetime):
        return h.time()
    if isinstance(h, time):
        return h
    s = str(h).strip()
    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            return datetime.strptime(s, fmt).time()
        except Exception:
            pass
    return None


def obtener_horario_turno(turno, dia_semana):
    """
    Extrae horas del texto de turno. Ej: '08:00-17:00 / 08:00-16:00'
    Para Lu-Ju (0..3) usa el primer par, para Viernes (4) el segundo si existe.
    """
    horas = re.findall(r"\d{1,2}:\d{2}", turno or "")
    # horas en orden: [ini1, fin1, ini2, fin2]
    if len(horas) < 2:
        return None, None
    if dia_semana in [0, 1, 2, 3]:
        return horas[0], horas[1]
    if dia_semana == 4 and len(horas) >= 4:
        return horas[2], horas[3]
    return None, None


def calcular_atraso(entrada, fecha, turno):
    ent = _parse_hora(entrada)
    fecha_dt = normalizar_fecha(fecha)
    if not ent or not fecha_dt:
        return 0

    ini_str, _ = obtener_horario_turno(turno, fecha_dt.weekday())
    if not ini_str:
        return 0

    hora_ini = datetime.strptime(ini_str, "%H:%M").time()
    if ent > hora_ini:
        return int((datetime.combine(fecha_dt, ent) - datetime.combine(fecha_dt, hora_ini)).total_seconds() / 60)
    return 0


def calcular_horas_extras(entrada, salida, fecha, turno, descripcion):
    """
    Devuelve (horas_50, horas_25) como floats.
    """
    if not entrada or not salida:
        return 0, 0
    if str(descripcion or "").lower() in ("ausente",) or "libre" in str(descripcion or "").lower():
        return 0, 0

    ent = _parse_hora(entrada)
    sal = _parse_hora(salida)
    fecha_dt = normalizar_fecha(fecha)
    if not ent or not sal or not fecha_dt:
        return 0, 0

    ent_dt = datetime.combine(fecha_dt, ent)
    sal_dt = datetime.combine(fecha_dt, sal)
    if sal_dt < ent_dt:
        sal_dt += pd.Timedelta(days=1)

    ini_str, fin_str = obtener_horario_turno(turno, fecha_dt.weekday())
    if not ini_str or not fin_str:
        total_min = (sal_dt - ent_dt).total_seconds() / 60
        return ((total_min / 60), 0) if total_min > 30 else (0, 0)

    hora_ini = datetime.combine(fecha_dt, datetime.strptime(ini_str, "%H:%M").time())
    hora_fin = datetime.combine(fecha_dt, datetime.strptime(fin_str, "%H:%M").time())

    min_50 = 0
    min_25 = 0

    # antes de la jornada
    if ent_dt < hora_ini:
        if ent_dt.time() < time(7, 0):
            min_50 += (min(sal_dt, hora_ini) - ent_dt).total_seconds() / 60
        else:
            min_25 += (min(sal_dt, hora_ini) - ent_dt).total_seconds() / 60

    # después de la jornada
    if sal_dt > hora_fin:
        if sal_dt > datetime.combine(fecha_dt, time(21, 0)):
            min_25 += max(0, (datetime.combine(fecha_dt, time(21, 0)) - hora_fin).total_seconds() / 60)
            min_50 += (sal_dt - datetime.combine(fecha_dt, time(21, 0))).total_seconds() / 60
        else:
            min_25 += (sal_dt - hora_fin).total_seconds() / 60

    h50 = round(min_50 / 60, 2) if min_50 > 30 else 0
    h25 = round(min_25 / 60, 2) if min_25 > 30 else 0
    return h50, h25


# ----------------- PROCESADOR PRINCIPAL (.xls) -----------------

def procesar_excel(file_stream):
    """
    Lee el .xls (engine=xlrd) que subes, genera un .xlsx con:
    - 'Detalle Diario'
    - 'Resumen'
    y colorea filas según descripción.
    """
    # Leemos TODO como texto para facilitar búsquedas (formato fijo)
    df = pd.read_excel(file_stream, header=None, dtype=str, engine="xlrd")
    df = df.fillna("")

    data_detalle = []
    resumen = {}

    i = 0
    n = len(df)

    funcionario = rut = organigrama = turno = periodo = ""

    while i < n:
        col0 = str(df.iat[i, 0]).strip().lower()

        # Bloque de encabezado de funcionario
        if col0.startswith("funcionario"):
            funcionario = str(df.iat[i, 1]).strip(": ").strip()
            rut         = str(df.iat[i + 1, 1]).strip(": ").strip() if i + 1 < n else ""
            organigrama = str(df.iat[i + 2, 1]).strip(": ").strip() if i + 2 < n else ""
            turno       = str(df.iat[i + 3, 1]).strip(": ").strip() if i + 3 < n else ""
            periodo     = str(df.iat[i + 4, 1]).strip(": ").strip() if i + 4 < n else ""
            i += 6
            continue

        # Encabezado de tabla diaria
        if col0 == "dia":
            i += 1
            # Recorrer filas de detalle hasta que aparezca otra cabecera o vacío
            while i < n:
                dia_txt = str(df.iat[i, 0]).strip().lower()
                if dia_txt in ("", "none", "totales"):
                    i += 1
                    continue
                if dia_txt.startswith("funcionario"):
                    # termina este bloque
                    break

                fecha = df.iat[i, 1]
                entrada = df.iat[i, 2]
                salida = df.iat[i, 3]
                # columna 5 (index 5) suele ser 'Descripción'
                descripcion = str(df.iat[i, 5]).strip()

                # convertir algunos posibles formatos de hora
                atras_min = calcular_atraso(entrada, fecha, turno)
                h50, h25 = calcular_horas_extras(entrada, salida, fecha, turno, descripcion)

                data_detalle.append([
                    funcionario, rut, organigrama, turno, periodo,
                    fecha, entrada, salida,
                    minutos_a_hhmm(atras_min),
                    horas_a_hhmm(h50),
                    horas_a_hhmm(h25),
                    descripcion
                ])

                if funcionario not in resumen:
                    resumen[funcionario] = {
                        "Rut": rut, "Organigrama": organigrama, "Turno": turno, "Periodo": periodo,
                        "Total 50%": 0.0, "Total 25%": 0.0, "Total Atraso": 0
                    }
                resumen[funcionario]["Total 50%"] += h50
                resumen[funcionario]["Total 25%"] += h25
                resumen[funcionario]["Total Atraso"] += atras_min

                i += 1
            continue

        i += 1

    # DataFrames
    df_detalle = pd.DataFrame(data_detalle, columns=[
        "Funcionario", "Rut", "Organigrama", "Turno", "Periodo",
        "Fecha", "Entrada", "Salida", "Atraso (hh:mm)", "50%", "25%", "Descripción"
    ])

    df_resumen = pd.DataFrame([
        [
            f,
            datos["Rut"], datos["Organigrama"], datos["Turno"], datos["Periodo"],
            horas_a_hhmm(datos["Total 50%"]),
            horas_a_hhmm(datos["Total 25%"]),
            minutos_a_hhmm(datos["Total Atraso"]),
            horas_a_hhmm(datos["Total 50%"] + datos["Total 25%"]),
        ]
        for f, datos in resumen.items()
    ], columns=["Funcionario", "Rut", "Organigrama", "Turno", "Periodo", "Total 50%", "Total 25%", "Total Atraso", "Total Horas"])

    # Escribir a xlsx en memoria
    mem = BytesIO()
    with pd.ExcelWriter(mem, engine="openpyxl") as writer:
        df_detalle.to_excel(writer, sheet_name="Detalle Diario", index=False)
        df_resumen.to_excel(writer, sheet_name="Resumen", index=False)

    mem.seek(0)
    wb = load_workbook(mem)
    ws = wb["Detalle Diario"]

    # Colores
    fill_rojo = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")     # Ausente
    fill_amar = PatternFill(start_color="FFFACD", end_color="FFFACD", fill_type="solid")     # Falta entrada/salida o '-'

    # Aplicar color por fila
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=12):
        descripcion = str(row[11].value or "")
        entrada = str(row[6].value or "").strip()
        salida  = str(row[7].value or "").strip()

        if "ausente" in descripcion.lower():
            for c in row:
                c.fill = fill_rojo
        elif "falta entrada" in descripcion.lower() or "falta salida" in descripcion.lower() or entrada == "-" or salida == "-":
            for c in row:
                c.fill = fill_amar

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out
