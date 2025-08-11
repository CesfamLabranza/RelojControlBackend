import pandas as pd
from datetime import datetime, time
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from io import BytesIO
import re

from bs4 import BeautifulSoup

# ========== Utilidades de fecha/tiempo ==========

def normalizar_fecha(fecha):
    """Devuelve date() si es datetime o intenta parsear dd-mm-YYYY / dd/mm/YYYY."""
    if isinstance(fecha, datetime):
        return fecha.date()
    try:
        s = str(fecha).strip()
        if "-" in s:
            return datetime.strptime(s, "%d-%m-%Y").date()
        if "/" in s:
            return datetime.strptime(s, "%d/%m/%Y").date()
    except Exception:
        return None
    return None


def convertir_a_hhmm(horas):
    minutos = int(round(float(horas) * 60))
    return f"{minutos // 60:02}:{minutos % 60:02}"


def minutos_a_hhmm(minutos):
    m = int(round(float(minutos)))
    return f"{m // 60:02}:{m % 60:02}"


def obtener_horario_turno(turno, dia_semana):
    """
    Espera cadenas tipo '08:00-17:00 / 08:00-16:00 (Vi)' etc.
    Simplificado: extrae HH:MM en orden y toma [0,1] para Lu-Ju, [2,3] para Vi si existen.
    """
    horas = re.findall(r"\d{1,2}:\d{2}", str(turno))
    if len(horas) < 2:
        return None, None
    if dia_semana in [0, 1, 2, 3]:
        return horas[0], horas[1]
    if dia_semana == 4 and len(horas) >= 4:
        return horas[2], horas[3]
    return horas[0], horas[1]  # fallback


def calcular_atraso(entrada, fecha, turno):
    if not entrada or entrada == "-":
        return 0
    f = normalizar_fecha(fecha)
    if not f:
        return 0
    try:
        ent_t = datetime.strptime(str(entrada), "%H:%M:%S").time()
    except Exception:
        try:
            ent_t = datetime.strptime(str(entrada), "%H:%M").time()
        except Exception:
            return 0

    dia_semana = f.weekday()
    inicio_str, _ = obtener_horario_turno(turno, dia_semana)
    if not inicio_str:
        return 0

    hora_inicio = datetime.strptime(inicio_str, "%H:%M").time()
    if ent_t > hora_inicio:
        atraso_min = (datetime.combine(f, ent_t) - datetime.combine(f, hora_inicio)).total_seconds() / 60
        return int(atraso_min)
    return 0


def calcular_horas_extras(entrada, salida, fecha, turno, descripcion):
    desc = str(descripcion or "").lower()
    if (not entrada or not salida or entrada == "-" or salida == "-" or
            "ausente" in desc or "libre" in desc):
        return 0, 0

    f = normalizar_fecha(fecha)
    if not f:
        return 0, 0

    # Parse HH:MM[:SS]
    def _parse_hhmm(x):
        for fmt in ("%H:%M:%S", "%H:%M"):
            try:
                return datetime.strptime(str(x), fmt)
            except Exception:
                pass
        return None

    ent_dt = _parse_hhmm(entrada)
    sal_dt = _parse_hhmm(salida)
    if not ent_dt or not sal_dt:
        return 0, 0

    if sal_dt < ent_dt:
        sal_dt += pd.Timedelta(days=1)

    dia_semana = f.weekday()
    inicio_str, fin_str = obtener_horario_turno(turno, dia_semana)
    if not inicio_str or not fin_str:
        total_min = (sal_dt - ent_dt).total_seconds() / 60
        return (total_min / 60, 0) if total_min > 30 else (0, 0)

    hora_inicio = datetime.combine(f, datetime.strptime(inicio_str, "%H:%M").time())
    hora_fin = datetime.combine(f, datetime.strptime(fin_str, "%H:%M").time())

    minutos_50 = 0
    minutos_25 = 0

    # Antes del inicio
    if ent_dt < hora_inicio:
        if ent_dt.time() < time(7, 0):
            minutos_50 += (min(sal_dt, hora_inicio) - ent_dt).total_seconds() / 60
        else:
            minutos_25 += (min(sal_dt, hora_inicio) - ent_dt).total_seconds() / 60

    # Después del fin
    if sal_dt > hora_fin:
        if sal_dt > datetime.combine(f, time(21, 0)):
            minutos_25 += max(0, (datetime.combine(f, time(21, 0)) - hora_fin).total_seconds() / 60)
            minutos_50 += (sal_dt - datetime.combine(f, time(21, 0))).total_seconds() / 60
        else:
            minutos_25 += (sal_dt - hora_fin).total_seconds() / 60

    horas_50 = round(minutos_50 / 60, 2) if minutos_50 > 30 else 0
    horas_25 = round(minutos_25 / 60, 2) if minutos_25 > 30 else 0
    return horas_50, horas_25


# ========== Excel real (.xlsx) ==========

def procesar_excel(file_stream):
    """
    Lee un Excel real con openpyxl (xlsx) que tenga una hoja activa con la misma
    estructura que el HTML exportado. Si tu Excel real es otra estructura,
    adapta este lector.
    """
    wb = load_workbook(filename=file_stream, data_only=True)
    sheet = wb.active

    data_detalle = []
    data_resumen = {}
    fila = 1
    max_row = sheet.max_row
    funcionario = rut = organigrama = turno = periodo = ""

    while fila <= max_row:
        celda = str(sheet.cell(row=fila, column=1).value or "").strip().lower()
        if celda.startswith("funcionario"):
            funcionario = str(sheet.cell(row=fila, column=2).value or "").strip(": ")
            rut = str(sheet.cell(row=fila + 1, column=2).value or "").strip(": ")
            organigrama = str(sheet.cell(row=fila + 2, column=2).value or "").strip(": ")
            turno = str(sheet.cell(row=fila + 3, column=2).value or "").strip(": ")
            periodo = str(sheet.cell(row=fila + 4, column=2).value or "").strip(": ")
            fila += 6
            continue

        if celda == "dia":
            fila += 1
            while fila <= max_row and sheet.cell(row=fila, column=1).value:
                dia_text = str(sheet.cell(row=fila, column=1).value or "").strip().lower()
                if dia_text == "totales":
                    fila += 1
                    continue
                if dia_text.startswith("funcionario") or dia_text == "none":
                    break

                fecha = sheet.cell(row=fila, column=2).value
                entrada = sheet.cell(row=fila, column=3).value
                salida = sheet.cell(row=fila, column=4).value
                descripcion = str(sheet.cell(row=fila, column=6).value or "").strip()

                atraso_min = calcular_atraso(entrada, fecha, turno)
                horas_50, horas_25 = calcular_horas_extras(entrada, salida, fecha, turno, descripcion)

                data_detalle.append([
                    funcionario, rut, organigrama, turno, periodo, fecha, entrada, salida,
                    minutos_a_hhmm(atraso_min), convertir_a_hhmm(horas_50), convertir_a_hhmm(horas_25), descripcion
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

    return _exportar_xlsx(data_detalle, data_resumen)


# ========== HTML “.xls” exportado ==========

def detectar_html_y_procesar(contenido_bytes: bytes):
    """
    Procesa el archivo HTML (típico 'exportar a .xls' de la plataforma).
    Se asume estructura:
      - Bloque de metadatos con filas 'Funcionario:', 'RUT:', 'Organigrama:', 'Turno:', 'Periodo:'
      - Luego tabla con cabeceras: Dia | Fecha | Hora Entrada | Hora Salida | ... | Descripción
    Si el HTML cambia, ajusta los selectores.
    """
    soup = BeautifulSoup(contenido_bytes, "lxml")  # lxml es rápido; html5lib también funciona

    # 1) Extraer todas las tablas
    tablas_bs = soup.find_all("table")
    if not tablas_bs:
        raise ValueError("No se encontraron tablas HTML en el archivo.")

    # Heurística: tabla de metadatos = tabla pequeña antes de la tabla de día/fecha
    # y la tabla grande contiene cabeceras 'Dia' / 'Fecha'
    df_list = []
    for t in tablas_bs:
        try:
            df = pd.read_html(str(t), flavor="bs4")[0]
            df_list.append(df)
        except Exception:
            continue

    if not df_list:
        raise ValueError("No fue posible leer tablas con pandas.read_html.")

    # Encuentra tabla de detalle (cabeceras esperadas)
    idx_detalle = None
    for i, d in enumerate(df_list):
        cols = [str(c).strip().lower() for c in d.columns]
        if any("dia" == c for c in cols) and any("fecha" == c for c in cols):
            idx_detalle = i
            break
    if idx_detalle is None:
        # Si no encontramos por cabecera, toma la tabla más grande
        idx_detalle = max(range(len(df_list)), key=lambda i: df_list[i].shape[0])

    detalle_raw = df_list[idx_detalle].copy()

    # Intenta leer metadatos de la primera(s) tabla(s)
    meta = {"Funcionario": "", "Rut": "", "Organigrama": "", "Turno": "", "Periodo": ""}
    for d in df_list[: idx_detalle + 1]:
        # Busca pares tipo 'Funcionario' -> valor en las primeras dos columnas
        for r in d.itertuples(index=False):
            c0 = str(r[0]).strip().lower()
            val = (str(r[1]).strip() if len(r) > 1 else "")
            if c0.startswith("funcionario"):
                meta["Funcionario"] = val.strip(": ")
            elif c0.startswith("rut"):
                meta["Rut"] = val.strip(": ")
            elif "organigrama" in c0:
                meta["Organigrama"] = val.strip(": ")
            elif c0.startswith("turno"):
                meta["Turno"] = val.strip(": ")
            elif c0.startswith("periodo"):
                meta["Periodo"] = val.strip(": ")

    # Normaliza cabeceras del detalle
    detalle_raw.columns = [str(c).strip().lower() for c in detalle_raw.columns]

    # Columnas candidatas
    col_dia = next((c for c in detalle_raw.columns if c.startswith("dia")), None)
    col_fecha = next((c for c in detalle_raw.columns if c.startswith("fecha")), None)
    col_ent = next((c for c in detalle_raw.columns if "entrada" in c), None)
    col_sal = next((c for c in detalle_raw.columns if "salida" in c), None)
    col_desc = next((c for c in detalle_raw.columns if "descr" in c or "observ" in c), None)

    if not (col_dia and col_fecha and col_ent and col_sal):
        raise ValueError("La tabla de detalle no contiene columnas esperadas (Dia/Fecha/Entrada/Salida).")

    # Filtra filas válidas (descarta totales/NaN masivos)
    detalle = detalle_raw.copy()
    detalle = detalle[detalle[col_fecha].notna()]

    data_detalle = []
    data_resumen = {}

    for _, row in detalle.iterrows():
        fecha = row[col_fecha]
        entrada = row[col_ent]
        salida  = row[col_sal]
        desc    = row[col_desc] if col_desc in row else ""

        atraso_min = calcular_atraso(entrada, fecha, meta["Turno"])
        h50, h25   = calcular_horas_extras(entrada, salida, fecha, meta["Turno"], desc)

        data_detalle.append([
            meta["Funcionario"], meta["Rut"], meta["Organigrama"], meta["Turno"], meta["Periodo"],
            fecha, entrada, salida, minutos_a_hhmm(atraso_min), convertir_a_hhmm(h50), convertir_a_hhmm(h25), str(desc or "")
        ])

        f = meta["Funcionario"]
        if f not in data_resumen:
            data_resumen[f] = {
                "Rut": meta["Rut"], "Organigrama": meta["Organigrama"], "Turno": meta["Turno"],
                "Periodo": meta["Periodo"], "Total 50%": 0, "Total 25%": 0, "Total Atraso": 0
            }
        data_resumen[f]["Total 50%"] += h50
        data_resumen[f]["Total 25%"] += h25
        data_resumen[f]["Total Atraso"] += atraso_min

    return _exportar_xlsx(data_detalle, data_resumen)


# ========== Exportador común ==========

def _exportar_xlsx(data_detalle, data_resumen):
    df_detalle = pd.DataFrame(data_detalle, columns=[
        "Funcionario", "Rut", "Organigrama", "Turno", "Periodo",
        "Fecha", "Entrada", "Salida", "Atraso (hh:mm)", "50%", "25%", "Descripción"
    ])

    df_resumen = pd.DataFrame([
        [
            f,
            v["Rut"], v["Organigrama"], v["Turno"], v["Periodo"],
            convertir_a_hhmm(v["Total 50%"]),
            convertir_a_hhmm(v["Total 25%"]),
            minutos_a_hhmm(v["Total Atraso"]),
            convertir_a_hhmm(v["Total 50%"] + v["Total 25%"])
        ]
        for f, v in data_resumen.items()
    ], columns=["Funcionario", "Rut", "Organigrama", "Turno", "Periodo",
                "Total 50%", "Total 25%", "Total Atraso", "Total Horas"])

    # Escribir a memoria
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_detalle.to_excel(writer, sheet_name="Detalle Diario", index=False)
        df_resumen.to_excel(writer, sheet_name="Resumen", index=False)

    # Pintar colores en Detalle (Ausente / Falta Entrada/Salida / '-')
    output.seek(0)
    wb_final = load_workbook(output)
    ws = wb_final["Detalle Diario"]
    fill_rojo = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    fill_amarillo = PatternFill(start_color="FFFACD", end_color="FFFACD", fill_type="solid")

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=12):
        descripcion = str(row[11].value or "")
        entrada = str(row[6].value or "").strip()
        salida  = str(row[7].value or "").strip()
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
