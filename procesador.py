import pandas as pd
from datetime import datetime, time
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from io import BytesIO
import re

# =================== FERIADOS (si aplica) ===================
FERIADOS_2025 = [
    "2025-01-01", "2025-04-18", "2025-04-19", "2025-05-01",
    "2025-05-21", "2025-06-29", "2025-07-16", "2025-08-15",
    "2025-09-18", "2025-09-19", "2025-10-12", "2025-10-31",
    "2025-11-01", "2025-12-08", "2025-12-25"
]
FERIADOS_2025 = [datetime.strptime(d, "%Y-%m-%d").date() for d in FERIADOS_2025]

# =================== AUXILIARES ===================

def normalizar_fecha(fecha):
    """Convierte 'dd-mm-aaaa' o 'dd/mm/aaaa' o datetime a date."""
    if isinstance(fecha, datetime):
        return fecha.date()
    try:
        s = str(fecha).strip()
        if "-" in s:
            return datetime.strptime(s, "%d-%m-%Y").date()
        if "/" in s:
            return datetime.strptime(s, "%d/%m/%Y").date()
    except:
        return None
    return None

def minutos_a_hhmm(minutos):
    return f"{int(minutos) // 60:02}:{int(minutos) % 60:02}"

def convertir_a_hhmm(horas):
    minutos = int(round(float(horas) * 60))
    return minutos_a_hhmm(minutos)

def obtener_horario_turno(turno, dia_semana):
    """
    Extrae horas del string de turno. Formato esperado:
    'Lu-Ju 08:00-17:00; Vi 08:00-16:00'  o similar
    Busca hh:mm en orden.
    """
    horas = re.findall(r"\d{1,2}:\d{2}", turno or "")
    if len(horas) < 2:
        return None, None
    # Lu-Ju
    if dia_semana in [0, 1, 2, 3]:
        return horas[0], horas[1]
    # Viernes
    if dia_semana == 4 and len(horas) >= 4:
        return horas[2], horas[3]
    return None, None

def calcular_atraso(entrada, fecha, turno):
    if not entrada or entrada == "-":
        return 0
    f = normalizar_fecha(fecha)
    if not f:
        return 0
    try:
        ent_t = datetime.strptime(str(entrada), "%H:%M:%S").time()
    except:
        return 0
    dia = f.weekday()
    ini_str, _ = obtener_horario_turno(turno, dia)
    if not ini_str:
        return 0
    hora_inicio = datetime.strptime(ini_str, "%H:%M").time()
    if ent_t > hora_inicio:
        return int((datetime.combine(f, ent_t) - datetime.combine(f, hora_inicio)).total_seconds() / 60)
    return 0

def calcular_horas_extras(entrada, salida, fecha, turno, descripcion):
    if not entrada or not salida or entrada == "-" or salida == "-":
        return 0, 0
    if str(descripcion).lower() in ["ausente", "libre"]:
        return 0, 0

    f = normalizar_fecha(fecha)
    if not f:
        return 0, 0
    try:
        e_dt = datetime.strptime(str(entrada), "%H:%M:%S")
        s_dt = datetime.strptime(str(salida), "%H:%M:%S")
    except:
        return 0, 0
    if s_dt < e_dt:
        s_dt += pd.Timedelta(days=1)

    dia = f.weekday()
    ini_str, fin_str = obtener_horario_turno(turno, dia)
    if not ini_str or not fin_str:
        total_min = (s_dt - e_dt).total_seconds() / 60
        return (total_min / 60, 0) if total_min > 30 else (0, 0)

    h_ini = datetime.combine(f, datetime.strptime(ini_str, "%H:%M").time())
    h_fin = datetime.combine(f, datetime.strptime(fin_str, "%H:%M").time())

    min50 = 0
    min25 = 0

    if e_dt < h_ini:
        if e_dt.time() < time(7, 0):
            min50 += (min(s_dt, h_ini) - e_dt).total_seconds() / 60
        else:
            min25 += (min(s_dt, h_ini) - e_dt).total_seconds() / 60

    if s_dt > h_fin:
        if s_dt > datetime.combine(f, time(21, 0)):
            min25 += max(0, (datetime.combine(f, time(21, 0)) - h_fin).total_seconds() / 60)
            min50 += (s_dt - datetime.combine(f, time(21, 0))).total_seconds() / 60
        else:
            min25 += (s_dt - h_fin).total_seconds() / 60

    h50 = round(min50 / 60, 2) if min50 > 30 else 0
    h25 = round(min25 / 60, 2) if min25 > 30 else 0
    return h50, h25

# =================== PIPELINE GENERICO ===================

def armar_resultados(registros, meta):
    """
    registros: lista de dicts con llaves:
      Fecha (date/str), Entrada "HH:MM:SS", Salida "HH:MM:SS", Descripcion (str)
    meta: dict con Funcionario, Rut, Organigrama, Turno, Periodo
    """
    detalle = []
    resumen = {}

    for r in registros:
        fecha = r.get("Fecha")
        entrada = r.get("Entrada")
        salida = r.get("Salida")
        descr = r.get("Descripcion", "")

        atraso_min = calcular_atraso(entrada, fecha, meta.get("Turno"))
        h50, h25 = calcular_horas_extras(entrada, salida, fecha, meta.get("Turno"), descr)

        detalle.append([
            meta.get("Funcionario", ""),
            meta.get("Rut", ""),
            meta.get("Organigrama", ""),
            meta.get("Turno", ""),
            meta.get("Periodo", ""),
            fecha, entrada, salida,
            minutos_a_hhmm(atraso_min),
            convertir_a_hhmm(h50),
            convertir_a_hhmm(h25),
            descr
        ])

        f = meta.get("Funcionario", "")
        if f not in resumen:
            resumen[f] = {
                "Rut": meta.get("Rut", ""),
                "Organigrama": meta.get("Organigrama", ""),
                "Turno": meta.get("Turno", ""),
                "Periodo": meta.get("Periodo", ""),
                "Total 50%": 0, "Total 25%": 0, "Total Atraso": 0
            }
        resumen[f]["Total 50%"] += h50
        resumen[f]["Total 25%"] += h25
        resumen[f]["Total Atraso"] += atraso_min

    df_detalle = pd.DataFrame(detalle, columns=[
        "Funcionario", "Rut", "Organigrama", "Turno", "Periodo",
        "Fecha", "Entrada", "Salida", "Atraso (hh:mm)", "50%", "25%", "Descripción"
    ])

    df_resumen = pd.DataFrame([
        [
            f, r["Rut"], r["Organigrama"], r["Turno"], r["Periodo"],
            convertir_a_hhmm(r["Total 50%"]),
            convertir_a_hhmm(r["Total 25%"]),
            minutos_a_hhmm(r["Total Atraso"]),
            convertir_a_hhmm(r["Total 50%"] + r["Total 25%"])
        ]
        for f, r in resumen.items()
    ], columns=["Funcionario", "Rut", "Organigrama", "Turno", "Periodo",
                "Total 50%", "Total 25%", "Total Atraso", "Total Horas"])

    # Exportar + colores
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df_detalle.to_excel(writer, sheet_name="Detalle Diario", index=False)
        df_resumen.to_excel(writer, sheet_name="Resumen", index=False)
    out.seek(0)

    wb = load_workbook(out)
    ws = wb["Detalle Diario"]
    fill_rojo = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    fill_amarillo = PatternFill(start_color="FFFACD", end_color="FFFACD", fill_type="solid")

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=12):
        descripcion = str(row[11].value or "")
        entrada = str(row[6].value or "").strip()
        salida = str(row[7].value or "").strip()
        if "Ausente" in descripcion:
            for c in row:
                c.fill = fill_rojo
        elif "Falta Entrada" in descripcion or "Falta Salida" in descripcion or entrada == "-" or salida == "-":
            for c in row:
                c.fill = fill_amarillo

    final = BytesIO()
    wb.save(final)
    final.seek(0)
    return final

# =================== .XLS/.XLSX (OpenXML) ===================

def procesar_excel(file_stream):
    """
    Intenta leer como Excel real (xlsx/xls binario) usando openpyxl.
    Pero tu archivo original (plataforma) no es un xls binario, por eso
    el app.py detecta si es HTML y nos manda por la otra ruta.
    """
    wb = load_workbook(filename=file_stream, data_only=True)
    sh = wb.active

    # El parser “antiguo” leía un layout específico con cabeceras sueltas.
    # Lo mantenemos por compatibilidad si alguien sube un Excel “limpio”.
    registros = []
    meta = {"Funcionario": "", "Rut": "", "Organigrama": "", "Turno": "", "Periodo": ""}

    fila = 1
    max_row = sh.max_row
    while fila <= max_row:
        celda = str(sh.cell(row=fila, column=1).value or "").strip().lower()
        if celda.startswith("funcionario"):
            meta["Funcionario"] = str(sh.cell(row=fila, column=2).value or "").strip(": ")
            meta["Rut"] = str(sh.cell(row=fila + 1, column=2).value or "").strip(": ")
            meta["Organigrama"] = str(sh.cell(row=fila + 2, column=2).value or "").strip(": ")
            meta["Turno"] = str(sh.cell(row=fila + 3, column=2).value or "").strip(": ")
            meta["Periodo"] = str(sh.cell(row=fila + 4, column=2).value or "").strip(": ")
            fila += 6
            continue

        if celda == "dia":
            fila += 1
            while fila <= max_row and sh.cell(row=fila, column=1).value:
                dtext = str(sh.cell(row=fila, column=1).value or "").strip().lower()
                if dtext == "totales":
                    fila += 1
                    continue
                if dtext.startswith("funcionario") or dtext == "none":
                    break

                fecha = sh.cell(row=fila, column=2).value
                entrada = sh.cell(row=fila, column=3).value
                salida = sh.cell(row=fila, column=4).value
                descr = str(sh.cell(row=fila, column=6).value or "").strip()

                registros.append({
                    "Fecha": fecha,
                    "Entrada": entrada,
                    "Salida": salida,
                    "Descripcion": descr
                })
                fila += 1
        else:
            fila += 1

    if not registros:
        # No encontramos layout; devolvemos los datos crudos para no romper.
        raise Exception("No se detectó el layout esperado en Excel.")

    return armar_resultados(registros, meta)

# =================== .XLS “HTML” ===================

def detectar_html_y_procesar(html_bytes):
    """
    Lee las tablas del HTML (archivo con extensión .xls generado por la plataforma),
    intenta mapear columnas y metadatos esperados. Si no puede, devuelve
    un XLSX con las tablas crudas.
    """
    # Intentar leer tablas
    try:
        dfs = pd.read_html(BytesIO(html_bytes), flavor="bs4")  # requiere beautifulsoup4 y lxml
    except Exception as e:
        raise Exception(f"No se pudieron leer tablas HTML: {e}")

    if not dfs:
        raise Exception("No se encontraron tablas en el archivo HTML.")

    # 1) Buscar metadatos en tablas “cabecera”
    meta = {"Funcionario": "", "Rut": "", "Organigrama": "", "Turno": "", "Periodo": ""}
    for df in dfs:
        # Examina pares clave/valor típicos: 'Funcionario', 'Rut', 'Organigrama', 'Turno', 'Periodo'
        for i in range(min(len(df), 10)):
            fila = [str(x).strip() for x in list(df.iloc[i].values)]
            if len(fila) >= 2:
                k = fila[0].lower()
                v = fila[1]
                if "funcionario" in k: meta["Funcionario"] = v
                elif "rut" in k: meta["Rut"] = v
                elif "organigrama" in k: meta["Organigrama"] = v
                elif "turno" in k: meta["Turno"] = v
                elif "periodo" in k or "período" in k: meta["Periodo"] = v

    # 2) Buscar la tabla de detalle (con Día/Fecha/Entrada/Salida/Descripción)
    candidatos = []
    for df in dfs:
        cols = [str(c).strip().lower() for c in df.columns]
        tiene_fecha = any(c in cols for c in ["fecha"])
        tiene_entrada = any("entrada" in c for c in cols)
        tiene_salida = any("salida" in c for c in cols)
        if tiene_fecha and (tiene_entrada or tiene_salida):
            candidatos.append(df)

    if not candidatos:
        # devolvemos las tablas “tal cual”, al menos unificadas en un solo xlsx
        return exportar_tablas_crudas(dfs)

    # Tomamos la “mejor” candidata (más columnas)
    detalle_df = max(candidatos, key=lambda d: d.shape[1])
    detalle_df = detalle_df.fillna("")

    # Mapeo flexible de columnas
    def pick(colnames, *opciones):
        low = [c.lower().strip() for c in colnames]
        for op in opciones:
            if op in low:
                return colnames[low.index(op)]
        return None

    col_fecha = pick(detalle_df.columns, "fecha")
    col_ent = pick(detalle_df.columns, "entrada", "hora entrada", "hora de entrada")
    col_sal = pick(detalle_df.columns, "salida", "hora salida", "hora de salida")
    col_desc = pick(detalle_df.columns, "descripcion", "descripción", "observacion", "observación", "detalle")

    registros = []
    for _, row in detalle_df.iterrows():
        fecha = row.get(col_fecha, "")
        ent = row.get(col_ent, "")
        sal = row.get(col_sal, "")
        des = row.get(col_desc, "")

        # Normalizar hora: si viene "08:00" -> "08:00:00"
        def to_hms(x):
            s = str(x).strip()
            if s in ["", "nan", "None"]:
                return "-"
            if re.match(r"^\d{1,2}:\d{2}:\d{2}$", s):
                return s
            if re.match(r"^\d{1,2}:\d{2}$", s):
                return s + ":00"
            return "-"

        registros.append({
            "Fecha": fecha,
            "Entrada": to_hms(ent),
            "Salida": to_hms(sal),
            "Descripcion": des
        })

    # Si no hay meta suficientes, al menos deja Turno vacío (no habrá cálculo fino)
    return armar_resultados(registros, meta)

def exportar_tablas_crudas(dfs):
    """
    Si no logramos mapear, devolvemos un XLSX con todas las tablas
    (para que el usuario igual reciba algo).
    """
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        for i, df in enumerate(dfs, 1):
            df.to_excel(writer, sheet_name=f"Tabla_{i}", index=False)
    out.seek(0)
    return out
