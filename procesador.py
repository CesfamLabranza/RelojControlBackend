import re
from io import BytesIO
from datetime import datetime, time

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# ──────────────────────────────────────────────────────────────────────────────
# Utilidades de fecha/hora
# ──────────────────────────────────────────────────────────────────────────────

def normalizar_fecha(fecha):
    """Acepta datetime, 'dd-mm-aaaa', 'dd/mm/aaaa'."""
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


def convertir_a_hhmm(horas_float):
    """0.75h -> '00:45'."""
    minutos = int(round(float(horas_float) * 60))
    return f"{minutos // 60:02}:{minutos % 60:02}"


def minutos_a_hhmm(minutos):
    return f"{int(minutos) // 60:02}:{int(minutos) % 60:02}"


def obtener_horario_turno(turno: str, dia_semana: int):
    """
    turno: string con horas (ej: '08:00-17:00 / 08:00-16:00 (vi)')
    Retorna (inicio, fin) para el día de la semana (0-Lun ... 4-Vie, etc.)
    """
    if not turno:
        return None, None
    horas = re.findall(r"\d{1,2}:\d{2}", turno)
    if len(horas) < 2:
        return None, None

    # L-J: primeras 2 horas
    if dia_semana in (0, 1, 2, 3):
        return horas[0], horas[1]
    # Viernes: si hay 4 horas, usar las 2 últimas
    if dia_semana == 4 and len(horas) >= 4:
        return horas[2], horas[3]
    # Otro día
    return None, None


def calcular_atraso(entrada, fecha, turno):
    """Retorna atraso en minutos (int)."""
    if not entrada or str(entrada).strip() in ("-", ""):
        return 0
    f = normalizar_fecha(fecha)
    if not f:
        return 0
    try:
        ent = datetime.strptime(str(entrada), "%H:%M:%S").time()
    except Exception:
        # Acepta HH:MM
        try:
            ent = datetime.strptime(str(entrada), "%H:%M").time()
        except Exception:
            return 0

    dia_semana = f.weekday()
    inicio_str, _ = obtener_horario_turno(turno, dia_semana)
    if not inicio_str:
        return 0

    hora_inicio = datetime.strptime(inicio_str, "%H:%M").time()
    if ent > hora_inicio:
        atraso = (
            datetime.combine(f, ent) - datetime.combine(f, hora_inicio)
        ).total_seconds() / 60
        return int(atraso)
    return 0


def calcular_horas_extras(entrada, salida, fecha, turno, descripcion):
    """
    Retorna (horas50, horas25) en horas decimales.
    Reglas: extras antes del inicio o después del fin. No cuenta ausente/libre.
    Tope: a partir de 21:00 las horas son 50%.
    """
    desc = (descripcion or "").lower()
    if any(p in desc for p in ("ausente", "libre")):
        return 0, 0
    if not entrada or not salida:
        return 0, 0
    if str(entrada).strip() in ("", "-") or str(salida).strip() in ("", "-"):
        return 0, 0

    f = normalizar_fecha(fecha)
    if not f:
        return 0, 0

    # Parse de horas
    def _to_dt(hhmmss):
        s = str(hhmmss)
        for fmt in ("%H:%M:%S", "%H:%M"):
            try:
                return datetime.strptime(s, fmt)
            except Exception:
                continue
        return None

    ent_dt = _to_dt(entrada)
    sal_dt = _to_dt(salida)
    if not ent_dt or not sal_dt:
        return 0, 0

    if sal_dt < ent_dt:
        sal_dt = sal_dt + pd.Timedelta(days=1)

    dia_semana = f.weekday()
    inicio_str, fin_str = obtener_horario_turno(turno, dia_semana)

    # Si no hay horario definido, considerar todo como 25% si supera 30 min
    if not inicio_str or not fin_str:
        total_min = (sal_dt - ent_dt).total_seconds() / 60
        return (total_min / 60, 0) if total_min > 30 else (0, 0)

    h_ini = datetime.combine(f, datetime.strptime(inicio_str, "%H:%M").time())
    h_fin = datetime.combine(f, datetime.strptime(fin_str, "%H:%M").time())
    limite_50 = datetime.combine(f, time(21, 0))

    min_25 = 0
    min_50 = 0

    # Antes del inicio
    if ent_dt < h_ini:
        if ent_dt.time() < time(7, 0):
            min_50 += (min(sal_dt, h_ini) - ent_dt).total_seconds() / 60
        else:
            min_25 += (min(sal_dt, h_ini) - ent_dt).total_seconds() / 60

    # Después del fin
    if sal_dt > h_fin:
        if sal_dt > limite_50:
            # 25% hasta 21:00, 50% después
            min_25 += max(0, (min(sal_dt, limite_50) - h_fin).total_seconds() / 60)
            if sal_dt > limite_50:
                min_50 += (sal_dt - max(h_fin, limite_50)).total_seconds() / 60
        else:
            min_25 += (sal_dt - h_fin).total_seconds() / 60

    h_25 = round(min_25 / 60, 2) if min_25 > 30 else 0
    h_50 = round(min_50 / 60, 2) if min_50 > 30 else 0
    return h_50, h_25


# ──────────────────────────────────────────────────────────────────────────────
# Helpers de salida (Excel con color)
# ──────────────────────────────────────────────────────────────────────────────

def _armar_excel_salida(detalle_rows, resumen_rows):
    df_detalle = pd.DataFrame(
        detalle_rows,
        columns=[
            "Funcionario", "Rut", "Organigrama", "Turno", "Periodo",
            "Fecha", "Entrada", "Salida",
            "Atraso (hh:mm)", "50%", "25%", "Descripción"
        ],
    )
    df_resumen = pd.DataFrame(
        resumen_rows,
        columns=[
            "Funcionario", "Rut", "Organigrama", "Turno", "Periodo",
            "Total 50%", "Total 25%", "Total Atraso", "Total Horas"
        ],
    )

    # 1) Escribimos con pandas a un buffer
    tmp = BytesIO()
    with pd.ExcelWriter(tmp, engine="openpyxl") as writer:
        df_detalle.to_excel(writer, sheet_name="Detalle Diario", index=False)
        df_resumen.to_excel(writer, sheet_name="Resumen", index=False)

    # 2) Cargamos con openpyxl para colorear
    tmp.seek(0)
    wb = load_workbook(tmp)
    ws = wb["Detalle Diario"]

    fill_rojo = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    fill_amarillo = PatternFill(start_color="FFFACD", end_color="FFFACD", fill_type="solid")

    # Recorremos filas (saltando encabezado)
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=12):
        descripcion = str(row[11].value or "")
        entrada = str(row[6].value or "").strip()
        salida = str(row[7].value or "").strip()

        if "Ausente" in descripcion or "AUSENTE" in descripcion:
            for c in row:
                c.fill = fill_rojo
        elif "Falta Entrada" in descripcion or "Falta Salida" in descripcion or entrada == "-" or salida == "-":
            for c in row:
                c.fill = fill_amarillo

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out


# ──────────────────────────────────────────────────────────────────────────────
# Procesamiento para XLS/XLSX reales
# ──────────────────────────────────────────────────────────────────────────────

def procesar_excel(stream: BytesIO) -> BytesIO:
    """
    Intenta:
      1) Abrir como XLSX con openpyxl y parsear por bloques (layout tipo reloj).
      2) Si falla, reintenta con pandas.read_excel(engine='xlrd') para .xls.
    """
    # Intento XLSX con openpyxl
    try:
        wb = load_workbook(filename=stream, data_only=True)
        sh = wb.active
        return _procesar_hoja_openpyxl(sh)
    except Exception:
        # Reintenta .xls binario usando pandas+xlrd
        try:
            stream.seek(0)
            df = pd.read_excel(stream, engine="xlrd", header=None)
            return _procesar_dataframe_generico(df)
        except Exception as e:
            raise RuntimeError(f"No se pudo leer como XLSX ni como XLS: {e}")


def _procesar_hoja_openpyxl(sh):
    """
    Parser estilo previo:
    - Busca bloque con 'Funcionario', siguiente líneas con 'Rut', 'Organigrama', 'Turno', 'Periodo'
    - Luego una tabla con encabezado 'Dia'/'Día'
    """
    detalle = []
    resumen = {}
    fila = 1
    max_row = sh.max_row

    meta = {"Funcionario": "", "Rut": "", "Organigrama": "", "Turno": "", "Periodo": ""}

    while fila <= max_row:
        v = sh.cell(row=fila, column=1).value
        celda = str(v).strip().lower() if v is not None else ""
        if celda.startswith("funcionario"):
            # Lee metadatos
            meta["Funcionario"] = str(sh.cell(row=fila, column=2).value or "").strip(": ")
            meta["Rut"] = str(sh.cell(row=fila + 1, column=2).value or "").strip(": ")
            meta["Organigrama"] = str(sh.cell(row=fila + 2, column=2).value or "").strip(": ")
            meta["Turno"] = str(sh.cell(row=fila + 3, column=2).value or "").strip(": ")
            meta["Periodo"] = str(sh.cell(row=fila + 4, column=2).value or "").strip(": ")
            fila += 6
            continue

        if celda in ("dia", "día"):
            fila += 1
            while fila <= max_row and sh.cell(row=fila, column=1).value:
                dia_text = str(sh.cell(row=fila, column=1).value).strip().lower()
                if dia_text in ("totales", "total"):
                    fila += 1
                    continue
                if dia_text.startswith("funcionario") or dia_text == "none":
                    break

                fecha = sh.cell(row=fila, column=2).value
                entrada = sh.cell(row=fila, column=3).value
                salida = sh.cell(row=fila, column=4).value
                descripcion = str(sh.cell(row=fila, column=6).value or "").strip()

                atraso_min = calcular_atraso(entrada, fecha, meta["Turno"])
                h50, h25 = calcular_horas_extras(entrada, salida, fecha, meta["Turno"], descripcion)

                detalle.append([
                    meta["Funcionario"], meta["Rut"], meta["Organigrama"], meta["Turno"], meta["Periodo"],
                    fecha, entrada, salida, minutos_a_hhmm(atraso_min),
                    convertir_a_hhmm(h50), convertir_a_hhmm(h25), descripcion
                ])

                k = meta["Funcionario"]
                if k not in resumen:
                    resumen[k] = {
                        "Rut": meta["Rut"], "Organigrama": meta["Organigrama"],
                        "Turno": meta["Turno"], "Periodo": meta["Periodo"],
                        "Total50": 0.0, "Total25": 0.0, "AtrasoMin": 0
                    }
                resumen[k]["Total50"] += h50
                resumen[k]["Total25"] += h25
                resumen[k]["AtrasoMin"] += atraso_min
                fila += 1
        else:
            fila += 1

    resumen_rows = []
    for f, r in resumen.items():
        total_horas = r["Total50"] + r["Total25"]
        resumen_rows.append([
            f, r["Rut"], r["Organigrama"], r["Turno"], r["Periodo"],
            convertir_a_hhmm(r["Total50"]), convertir_a_hhmm(r["Total25"]),
            minutos_a_hhmm(r["AtrasoMin"]), convertir_a_hhmm(total_horas),
        ])

    return _armar_excel_salida(detalle, resumen_rows)


def _procesar_dataframe_generico(df: pd.DataFrame) -> BytesIO:
    """
    Fallback genérico para .xls con pandas.
    Intenta encontrar un bloque de tabla donde existan columnas Fecha/Entrada/Salida/Descripción.
    """
    # Buscar encabezado por palabras clave (muy tolerante)
    headers_candidates = {}
    for i in range(min(40, len(df))):
        fila = df.iloc[i].astype(str).str.lower().tolist()
        if any("fecha" in c for c in fila) and any("entrada" in c for c in fila) and any("salida" in c for c in fila):
            headers_candidates[i] = fila

    if not headers_candidates:
        raise RuntimeError("No se encontraron encabezados compatibles en el XLS (pandas).")

    # Toma el primero que parece válido
    hdr_row = sorted(headers_candidates.keys())[0]
    df_tabla = df.iloc[hdr_row + 1 :].reset_index(drop=True)
    cols_map = {j: str(c).strip().lower() for j, c in enumerate(df.iloc[hdr_row].tolist())}

    def col_idx(nombre):
        for j, c in cols_map.items():
            if nombre in c:
                return j
        return None

    idx_fecha = col_idx("fecha")
    idx_entrada = col_idx("entrada")
    idx_salida = col_idx("salida")
    idx_desc = col_idx("descrip") or col_idx("observ") or col_idx("detalle")

    if idx_fecha is None or idx_entrada is None or idx_salida is None:
        raise RuntimeError("No se hallaron columnas obligatorias (fecha/entrada/salida) en el XLS.")

    # Metadatos básicos (si se ven arriba en las primeras filas)
    def buscar_meta(df_all, clave):
        # Busca "Funcionario", "Rut", etc. en las primeras 30 filas/2 columnas
        for i in range(min(30, len(df_all))):
            fila = df_all.iloc[i, :4].astype(str).tolist()
            fila_l = [s.lower() for s in fila]
            for j, s in enumerate(fila_l):
                if clave in s:
                    # valor en siguiente columna si existe
                    try:
                        return str(df_all.iloc[i, j + 1])
                    except Exception:
                        return ""
        return ""

    meta_func = buscar_meta(df, "funcionario")
    meta_rut = buscar_meta(df, "rut")
    meta_org = buscar_meta(df, "organigrama")
    meta_turno = buscar_meta(df, "turno")
    meta_periodo = buscar_meta(df, "periodo")

    detalle = []
    resumen = {}

    for i in range(len(df_tabla)):
        fila = df_tabla.iloc[i].tolist()
        fecha = fila[idx_fecha] if idx_fecha is not None else ""
        entrada = fila[idx_entrada] if idx_entrada is not None else ""
        salida = fila[idx_salida] if idx_salida is not None else ""
        descripcion = fila[idx_desc] if idx_desc is not None else ""

        # Filas vacías o separadores
        if pd.isna(fecha) and pd.isna(entrada) and pd.isna(salida):
            continue

        atraso_min = calcular_atraso(entrada, fecha, meta_turno)
        h50, h25 = calcular_horas_extras(entrada, salida, fecha, meta_turno, descripcion)

        detalle.append([
            meta_func, meta_rut, meta_org, meta_turno, meta_periodo,
            fecha, entrada, salida,
            minutos_a_hhmm(atraso_min), convertir_a_hhmm(h50), convertir_a_hhmm(h25), descripcion
        ])

        k = meta_func or "(Funcionario)"
        if k not in resumen:
            resumen[k] = {
                "Rut": meta_rut, "Organigrama": meta_org,
                "Turno": meta_turno, "Periodo": meta_periodo,
                "Total50": 0.0, "Total25": 0.0, "AtrasoMin": 0
            }
        resumen[k]["Total50"] += h50
        resumen[k]["Total25"] += h25
        resumen[k]["AtrasoMin"] += atraso_min

    resumen_rows = []
    for f, r in resumen.items():
        total_horas = r["Total50"] + r["Total25"]
        resumen_rows.append([
            f, r["Rut"], r["Organigrama"], r["Turno"], r["Periodo"],
            convertir_a_hhmm(r["Total50"]), convertir_a_hhmm(r["Total25"]),
            minutos_a_hhmm(r["AtrasoMin"]), convertir_a_hhmm(total_horas),
        ])

    return _armar_excel_salida(detalle, resumen_rows)


# ──────────────────────────────────────────────────────────────────────────────
# Procesamiento para XLS-HTML (archivo HTML con extensión .xls)
# ──────────────────────────────────────────────────────────────────────────────

def detectar_html_y_procesar(html_bytes: bytes) -> BytesIO:
    """
    Lee tablas desde HTML (exportado por el reloj) usando pandas.read_html.
    Intenta con codificación latin-1 y utf-8.
    Busca una tabla con columnas Fecha/Entrada/Salida (y opcional Descripción).
    """
    tablas = None
    error_capturado = None

    for enc in ("latin-1", "utf-8"):
        try:
            tablas = pd.read_html(BytesIO(html_bytes), flavor="bs4", encoding=enc)
            if tablas and len(tablas) > 0:
                break
        except Exception as e:
            error_capturado = e
            tablas = None

    if not tablas:
        raise RuntimeError(f"No se pudieron leer tablas HTML: {error_capturado or 'Desconocido'}")

    # Heurística: elegir la tabla que contenga columnas de interés
    idx_tabla = None
    for i, t in enumerate(tablas):
        cols = [str(c).lower() for c in t.columns]
        if any("fecha" in c for c in cols) and any("entrada" in c for c in cols) and any("salida" in c for c in cols):
            idx_tabla = i
            break
    if idx_tabla is None:
        # Si no hay columnas con nombre, intenta la primera y asumimos orden
        idx_tabla = 0

    df = tablas[idx_tabla].copy()
    # Limpia encabezados tipo multiindex o filas basura repetidas
    df.columns = [str(c).strip() for c in df.columns]
    df = df.rename(columns=lambda c: c.strip())

    # Intenta detectar columnas
    def col_idx(nombre):
        for j, c in enumerate(df.columns):
            if nombre in str(c).lower():
                return j
        return None

    i_fecha = col_idx("fecha")
    i_ent = col_idx("entrada")
    i_sal = col_idx("salida")
    i_desc = col_idx("descrip") or col_idx("observ") or col_idx("detalle")

    if i_fecha is None or i_ent is None or i_sal is None:
        # algunos exportan sin header claro: asume posiciones típicas
        # (Fecha, Entrada, Salida, ..., Descripción)
        if df.shape[1] >= 4:
            i_fecha, i_ent, i_sal = 0, 1, 2
            i_desc = 4 if df.shape[1] > 4 else None
        else:
            raise RuntimeError("La tabla HTML no tiene columnas reconocibles (fecha/entrada/salida).")

    # Metadatos (si existen en tablas vecinas, se podrían inferir aquí.
    # Para mantener robusto, lo dejamos en blanco o 'Desconocido')
    meta_func = "Funcionario"
    meta_rut = ""
    meta_org = ""
    meta_turno = ""
    meta_periodo = ""

    detalle = []
    resumen = {}

    for _, row in df.iterrows():
        fecha = row.iloc[i_fecha] if i_fecha is not None else ""
        entrada = row.iloc[i_ent] if i_ent is not None else ""
        salida = row.iloc[i_sal] if i_sal is not None else ""
        descripcion = row.iloc[i_desc] if i_desc is not None and i_desc < len(row) else ""

        # Filas muy vacías -> saltar
        if (pd.isna(fecha) or str(fecha).strip() == "") and (pd.isna(entrada) or pd.isna(salida)):
            continue

        atraso_min = calcular_atraso(entrada, fecha, meta_turno)
        h50, h25 = calcular_horas_extras(entrada, salida, fecha, meta_turno, descripcion)

        detalle.append([
            meta_func, meta_rut, meta_org, meta_turno, meta_periodo,
            fecha, entrada, salida,
            minutos_a_hhmm(atraso_min), convertir_a_hhmm(h50), convertir_a_hhmm(h25), descripcion
        ])

        k = meta_func
        if k not in resumen:
            resumen[k] = {
                "Rut": meta_rut, "Organigrama": meta_org,
                "Turno": meta_turno, "Periodo": meta_periodo,
                "Total50": 0.0, "Total25": 0.0, "AtrasoMin": 0
            }
        resumen[k]["Total50"] += h50
        resumen[k]["Total25"] += h25
        resumen[k]["AtrasoMin"] += atraso_min

    resumen_rows = []
    for f, r in resumen.items():
        total_horas = r["Total50"] + r["Total25"]
        resumen_rows.append([
            f, r["Rut"], r["Organigrama"], r["Turno"], r["Periodo"],
            convertir_a_hhmm(r["Total50"]), convertir_a_hhmm(r["Total25"]),
            minutos_a_hhmm(r["AtrasoMin"]), convertir_a_hhmm(total_horas),
        ])

    return _armar_excel_salida(detalle, resumen_rows)
