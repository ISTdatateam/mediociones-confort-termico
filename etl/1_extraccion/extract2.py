#!/usr/bin/env python3
from pathlib import Path
import openpyxl
import csv
import sys
import math, re
import datetime as dt

# =========================
#  PARÁMETROS GENERALES
# =========================

NUM_AREAS_BASICAS = 11    # Áreas con Area_i / N_personas_i / Largo_i / Ancho_i / Alto_i
NUM_AREAS_DETALLE = 11    # Áreas con Area_i_bis, iny/ext, velocidades, etc.

# =========================
#  GENERADORES DE SCHEMA
# =========================

def build_schema_areas_basicas(num_areas=11, start_row=31, row_step=4):
    """
    Genera bloques:
      Area_i (col 2), N_personas_i (col 5), Largo_i (7), Ancho_i (9), Alto_i (11)
    empezando en start_row y avanzando de row_step filas por área.
    """
    schema = []
    for i in range(1, num_areas + 1):
        base_row = start_row + (i - 1) * row_step
        schema.extend([
            {"row": base_row, "col": 2,  "label": f"Area_{i}",        "type": "text"},
            {"row": base_row, "col": 5,  "label": f"N_personas_{i}",  "type": "integer"},
            {"row": base_row, "col": 7,  "label": f"Largo_{i}",       "type": "number"},
            {"row": base_row, "col": 9,  "label": f"Ancho_{i}",       "type": "number"},
            {"row": base_row, "col": 11, "label": f"Alto_{i}",        "type": "number"},
        ])
    return schema


def build_schema_areas_detalle(
    num_areas=11,
    start_row_bis=79,      # fila donde parte Area_1_bis en el NUEVO formato
    block_height=12        # altura (en filas) ocupada por cada área detallada
):
    """
    Genera bloques de detalle por área:
      - Area_i_bis (fila base), Area_i_desc
      - Area_i_iny_1..5, Area_i_ext_1..5
      - Area_i_iny_largo/ ancho / diam_1..5, Area_i_ext_...
      - Matriz de velocidades 5x5 para iny y ext: Area_i_iny_vel_f_c, Area_i_ext_vel_f_c
    Siguiendo el patrón que mostraste, pero iniciando en start_row_bis y repitiendo
    cada block_height filas por área.
    """
    schema = []

    for i in range(1, num_areas + 1):
        r_bis = start_row_bis + (i - 1) * block_height
        r_iny1 = r_bis + 1   # primera fila de inyección
        r_ext1 = r_bis + 7   # primera fila de extracción

        # ---- Cabecera del área (bis / desc) ----
        schema.extend([
            {"row": r_bis, "col": 2, "label": f"Area_{i}_bis",  "type": "text"},
            {"row": r_bis, "col": 6, "label": f"Area_{i}_desc", "type": "text"},
        ])

        # ---- iny / ext: nombres (texto) ----
        for j in range(1, 6):  # 1..5 bocas
            # Inyección
            schema.append({
                "row": r_iny1 + (j - 1),
                "col": 15,
                "label": f"Area_{i}_iny_{j}",
                "type": "text",
            })
            # Extracción
            schema.append({
                "row": r_ext1 + (j - 1),
                "col": 15,
                "label": f"Area_{i}_ext_{j}",
                "type": "text",
            })

        # ---- iny / ext: largo, ancho, diámetro ----
        for j in range(1, 6):
            # Inyección
            row_iny = r_iny1 + (j - 1)
            # Largo
            schema.append({
                "row": row_iny, "col": 19,
                "label": f"Area_{i}_iny_largo_{j}", "type": "number"
            })
            # Ancho
            schema.append({
                "row": row_iny, "col": 20,
                "label": f"Area_{i}_iny_ancho_{j}", "type": "number"
            })
            # Diámetro
            schema.append({
                "row": row_iny, "col": 21,
                "label": f"Area_{i}_iny_diam_{j}", "type": "number"
            })

            # Extracción
            row_ext = r_ext1 + (j - 1)
            schema.append({
                "row": row_ext, "col": 19,
                "label": f"Area_{i}_ext_largo_{j}", "type": "number"
            })
            schema.append({
                "row": row_ext, "col": 20,
                "label": f"Area_{i}_ext_ancho_{j}", "type": "number"
            })
            schema.append({
                "row": row_ext, "col": 21,
                "label": f"Area_{i}_ext_diam_{j}", "type": "number"
            })

        # ---- Velocidades iny / ext (matriz 5x5) ----
        for fila in range(1, 6):   # 1..5 filas de medición
            for col_idx in range(1, 6):  # 1..5 puntos
                col_vel = 23 + (col_idx - 1)

                # Inyección
                schema.append({
                    "row": r_iny1 + (fila - 1),
                    "col": col_vel,
                    "label": f"Area_{i}_iny_vel_{fila}_{col_idx}",
                    "type": "number",
                })

                # Extracción
                schema.append({
                    "row": r_ext1 + (fila - 1),
                    "col": col_vel,
                    "label": f"Area_{i}_ext_vel_{fila}_{col_idx}",
                    "type": "number",
                })

    return schema

# =========================
#  SCHEMA BASE (cabecera)
# =========================

SCHEMA_BASE = [
    {"row": 4, "col": 7,  "label": "RAZON SOCIAL EMPRESA", "type": "text"},
    {"row": 4, "col": 19, "label": "RUT EMPRESA",            "type": "text"},
    {"row": 6, "col": 7,  "label": "NOMBRE DEL CENTRO DE TRABAJO (CT)", "type": "text"},
    {"row": 6, "col": 19, "label": "HORARIO DE MEDICIÓN",    "type": "time"},   # <-- antes era "text"
    {"row": 8, "col": 7,  "label": "DIRECCIÓN (CT)",         "type": "text"},
    {"row": 10,"col": 7,  "label": "FECHA DE VISITA",        "type": "text"},
    {"row": 14,"col": 7,  "label": "NOMBRE PERSONA EMPRESA", "type": "text"},
    {"row": 14,"col": 19, "label": "NOMBRE PROFESIONAL IST", "type": "text"},
    {"row": 16,"col": 7,  "label": "CARGO PERSONA EMPRESA",  "type": "text"},
    {"row": 16,"col": 19, "label": "CARGO PROFESIONAL IST",  "type": "text"},

    {"row": 21,"col": 2,  "label": "Instru_nombre_1",        "type": "text"},
    {"row": 21,"col": 5,  "label": "Instru_marca_1",         "type": "text"},
    {"row": 21,"col": 7,  "label": "Instru_modelo_1",        "type": "text"},
    {"row": 21,"col": 9,  "label": "Instru_Nserie_1",        "type": "text"},
    {"row": 21,"col": 12, "label": "Instru_Ncertificado_1",  "type": "text"},

    {"row": 22,"col": 2,  "label": "Instru_nombre_2",        "type": "text"},
    {"row": 22,"col": 5,  "label": "Instru_marca_2",         "type": "text"},
    {"row": 22,"col": 7,  "label": "Instru_modelo_2",        "type": "text"},
    {"row": 22,"col": 9,  "label": "Instru_Nserie_2",        "type": "text"},
    {"row": 22,"col": 12, "label": "Instru_Ncertificado_2",  "type": "text"},
]

# =========================
#  SCHEMA COMPLETO
# =========================

SCHEMA = []
SCHEMA += SCHEMA_BASE
SCHEMA += build_schema_areas_basicas(
    num_areas=NUM_AREAS_BASICAS,
    start_row=31,
    row_step=4
)
SCHEMA += build_schema_areas_detalle(
    num_areas=NUM_AREAS_DETALLE,
    start_row_bis=79,   # "Ahora partirían en la fila 77"
    block_height=12
)

# =========================
#  CAMPOS DERIVADOS
# =========================

DERIVED_FIELDS = [
    f"Volumen_{i}" for i in range(1, NUM_AREAS_BASICAS + 1)
] + [
    f"m3xpersona_{i}" for i in range(1, NUM_AREAS_BASICAS + 1)
] + [
    f"Resultado_{i}" for i in range(1, NUM_AREAS_BASICAS + 1)
] + [
    "Estado"
]

def _is_empty(v):
    if v is None:
        return True
    if isinstance(v, str) and v.strip() == "":
        return True
    return False

def _to_number(v):
    if v is None:
        return float("nan")
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    s = str(v).strip()
    if s == "":
        return float("nan")
    s1 = s.replace(" ", "")
    # 1.234,56 -> 1234.56
    if re.match(r"^-?\d{1,3}(\.\d{3})*,\d+$", s1):
        s1 = s1.replace(".", "").replace(",", ".")
        try: return float(s1)
        except: return float("nan")
    # 1,234.56 -> 1234.56
    if re.match(r"^-?\d{1,3}(,\d{3})*\.\d+$", s1):
        s1 = s1.replace(",", "")
        try: return float(s1)
        except: return float("nan")
    # 1.234.567 o 1,234,567 -> 1234567
    if re.match(r"^-?\d{1,3}([.,]\d{3})+$", s1):
        s2 = s1.replace(",", "").replace(".", "")
        try: return float(s2)
        except: return float("nan")
    # genérico: coma como decimal
    s2 = s1.replace(",", ".")
    try: return float(s2)
    except: return float("nan")

def _fmt_number(val_float):
    """Devuelve string sin ceros a la derecha ni punto sobrante."""
    if val_float is None or (isinstance(val_float, float) and math.isnan(val_float)):
        return ""
    s = f"{float(val_float):.10f}".rstrip("0").rstrip(".")
    return s if s != "-0" else "0"

def _to_number_str(v):
    """Parsea a float y devuelve string normalizado (mantiene 3.128, 2.98, etc.)."""
    f = _to_number(v)
    return _fmt_number(f)

def _to_int(v):
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return ""
    if isinstance(v, bool):
        return "1" if v else "0"
    if isinstance(v, (int,)):
        return str(int(v))
    if isinstance(v, float):
        return str(int(round(v)))
    s = str(v).strip()
    if s == "":
        return ""
    # quitar palabras/espacios/separadores y dejar números y signo
    s = s.replace(" ", "")
    s = s.replace(".", "").replace(",", "")
    m = re.match(r"^(-?\d+)", s)
    if not m:
        f = _to_number(v)
        return "" if (isinstance(f, float) and math.isnan(f)) else str(int(round(f)))
    return str(int(m.group(1)))

def _to_time(v):
    """Devuelve HH:MM:SS. Acepta datetime, '14:30', '14:30:00', '15:00 HORAS', etc."""
    if v is None:
        return ""
    if isinstance(v, (dt.datetime, dt.time)):
        t = v.time() if isinstance(v, dt.datetime) else v
        return t.strftime("%H:%M:%S")
    s = str(v).strip().upper()
    # limpiar palabras comunes (HRS, HORAS, HORA)
    s = re.sub(r"\b(HRS?|HORAS?)\b", "", s)
    s = s.replace(".", ":")
    s = re.sub(r"\s+", "", s)
    # patrones HH:MM o HH:MM:SS
    m = re.match(r"^(\d{1,2}):(\d{2})(?::(\d{2}))?$", s)
    if m:
        h = int(m.group(1)); mnt = int(m.group(2)); sec = int(m.group(3) or 0)
        h = max(0, min(23, h)); mnt = max(0, min(59, mnt)); sec = max(0, min(59, sec))
        return f"{h:02d}:{mnt:02d}:{sec:02d}"
    # si viene 4 dígitos (1430) -> 14:30:00
    m = re.match(r"^(\d{1,2})(\d{2})$", s)
    if m:
        h = int(m.group(1)); mnt = int(m.group(2))
        h = max(0, min(23, h)); mnt = max(0, min(59, mnt))
        return f"{h:02d}:{mnt:02d}:00"
    return ""  # no reconocible

def _cast(v, t):
    if t == "integer":
        return _to_int(v)                  # -> "2", "1", etc.
    if t == "time":
        return _to_time(v)                 # -> "HH:MM:SS"
    if t == "number":
        return _to_number_str(v)           # -> "3.128", "2.98", "5.67"
    # text (incluye fecha V1 como texto plano)
    if isinstance(v, (dt.date, dt.datetime)):
        return v.isoformat()
    return "" if v is None else str(v).strip()

def _calc_area_metrics(largo, ancho, alto, n_personas):
    """
    largo/ancho/alto/n_personas: ya parseados con _to_number (float) o NaN.
    Retorna (volumen_str, m3xpersona_str, resultado_str) con:
      - volumen = largo * ancho * alto
      - m3xpersona = volumen / n_personas  (si n_personas <= 0 o NaN -> "")
      - resultado = "Cumple" si m3xpersona > 10, si no "No cumple"; si no hay datos -> ""
    """
    # Validar dimensiones
    if any([(isinstance(x, float) and math.isnan(x)) for x in (largo, ancho, alto)]):
        return ("", "", "")
    vol = float(largo) * float(ancho) * float(alto)
    vol_str = _fmt_number(vol)

    # Validar personas
    if n_personas is None or (isinstance(n_personas, float) and math.isnan(n_personas)) or n_personas <= 0:
        return (vol_str, "", "")

    m3pp = vol / float(n_personas)
    m3pp_str = _fmt_number(m3pp)

    # Umbral estricto: > 10 => Cumple; en otro caso => No cumple
    resultado = "Cumple" if m3pp > 10 else "No cumple"
    return (vol_str, m3pp_str, resultado)



def pick_best_sheet(wb):
    best_sheet = None
    best_hits = -1
    for sname in wb.sheetnames:
        ws = wb[sname]
        hits = 0
        for f in SCHEMA:
            v = ws.cell(row=f["row"], column=f["col"]).value
            if not _is_empty(v):
                hits += 1
        if hits > best_hits:
            best_hits = hits
            best_sheet = sname
    return best_sheet, best_hits

def process_file(path: Path):
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=False, keep_links=False)
    except Exception as e:
        return {"__error__": f"No se pudo abrir: {e}", "__source_file__": str(path)}

    sheet, hits = pick_best_sheet(wb)
    if sheet is None:
        return {"__error__": "No se detectó una hoja válida", "__source_file__": str(path)}

    ws = wb[sheet]
    row_dict = {"__source_file__": str(path), "__sheet__": sheet, "__hits__": hits}

    # 1) Guardar valores base según SCHEMA (con casteo)
    for f in SCHEMA:
        raw = ws.cell(row=f["row"], column=f["col"]).value
        row_dict[f["label"]] = _cast(raw, f["type"])

    # 2) Calcular métricas por área (usando los RAW parseados a número)
    # Coordenadas por área (filas base) → 31, 35, 39, ..., hasta NUM_AREAS_BASICAS
    area_rows = {i: 31 + (i - 1) * 4 for i in range(1, NUM_AREAS_BASICAS + 1)}
    # Columnas (E=5, G=7, I=9, K=11)
    COL_E, COL_G, COL_I, COL_K = 5, 7, 9, 11

    resultados_por_area = []
    for i, r in area_rows.items():
        # Raw values (sin casteo a string) para cálculos
        raw_np = ws.cell(row=r, column=COL_E).value
        raw_l  = ws.cell(row=r, column=COL_G).value
        raw_an = ws.cell(row=r, column=COL_I).value
        raw_al = ws.cell(row=r, column=COL_K).value

        np_num = _to_number(raw_np)
        l_num  = _to_number(raw_l)
        an_num = _to_number(raw_an)
        al_num = _to_number(raw_al)

        vol_str, m3pp_str, res_str = _calc_area_metrics(l_num, an_num, al_num, np_num)

        row_dict[f"Volumen_{i}"]     = vol_str
        row_dict[f"m3xpersona_{i}"]  = m3pp_str
        row_dict[f"Resultado_{i}"]   = res_str
        resultados_por_area.append(res_str)

    # 3) Estado global
    # - Si alguna área es "No cumple" -> Pendiente
    # - Si al menos una es "Cumple" y ninguna "No cumple" -> Cerrado
    # - Si todas vacías -> ""
    if any(r == "No cumple" for r in resultados_por_area):
        row_dict["Estado"] = "Pendiente"
    else:
        hay_cumple = any(r == "Cumple" for r in resultados_por_area)
        row_dict["Estado"] = "Cerrado" if hay_cumple else ""

    return row_dict

def discover_files(input_dir: Path):
    return sorted([p for p in input_dir.rglob("*") if p.suffix.lower() in [".xlsx", ".xlsm"]])

def write_csv(rows, out_csv: Path):
    meta = ["__source_file__", "__sheet__", "__hits__"]
    fields = [f["label"] for f in SCHEMA]
    headers = meta + fields + DERIVED_FIELDS   # <-- agrega derivados
    with out_csv.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=headers)
        w.writeheader()
        for r in rows:
            if "__error__" in r:
                base = {"__source_file__": r.get("__source_file__", ""), "__sheet__": "", "__hits__": -1}
                for k in (fields + DERIVED_FIELDS): base[k] = ""
                w.writerow(base)
            else:
                w.writerow({h: r.get(h, "") for h in headers})

from openpyxl import Workbook
from openpyxl.utils import get_column_letter

def _to_excel_number(s):
    """Convierte strings numéricos a float; deja vacío como ''."""
    if s is None:
        return ""
    if isinstance(s, (int, float)):
        return s
    s = str(s).strip()
    if s == "":
        return ""
    try:
        return float(s)
    except:
        return s  # si no es número, dejar como texto

def write_xlsx(rows, out_xlsx: Path, schema, derived_fields):
    """
    rows: lista de dicts (cada archivo = una fila)
    out_xlsx: ruta .xlsx
    schema: SCHEMA original para conocer tipos
    derived_fields: lista DERIVED_FIELDS
    """
    # columnas
    meta = ["__source_file__", "__sheet__", "__hits__"]
    base_fields = [f["label"] for f in schema]
    headers = meta + base_fields + derived_fields

    # mapa de tipos de SCHEMA para casteo selectivo
    type_by_label = {f["label"]: f["type"] for f in schema}

    wb = Workbook()
    ws = wb.active
    ws.title = "Resultados"

    # escribe headers
    for j, h in enumerate(headers, start=1):
        ws.cell(row=1, column=j, value=h)

    # escribe filas
    r_out = 2
    for r in rows:
        # aseguremos todas las claves
        row_vals = {h: r.get(h, "") for h in headers}

        for j, h in enumerate(headers, start=1):
            v = row_vals[h]

            # meta: dejar como texto/num natural
            if h in ["__hits__"]:
                # hits siempre numérico si viene
                v2 = _to_excel_number(v)
                ws.cell(row=r_out, column=j, value=v2 if v2 != "" else None)
                continue
            if h in ["__source_file__", "__sheet__"]:
                ws.cell(row=r_out, column=j, value=None if v == "" else v)
                continue

            # campos del schema con tipo
            if h in type_by_label:
                t = type_by_label[h]
                if t == "integer":
                    # forzamos a int si se puede
                    try:
                        v2 = int(float(str(v).replace(",", ".")))
                        ws.cell(row=r_out, column=j, value=v2)
                    except:
                        ws.cell(row=r_out, column=j, value=None if v == "" else v)
                elif t == "number":
                    v2 = _to_excel_number(v)
                    ws.cell(row=r_out, column=j, value=v2 if v2 != "" else None)
                elif t == "time":
                    # dejamos como texto HH:MM:SS
                    ws.cell(row=r_out, column=j, value=None if v == "" else v)
                else:
                    # text / otros
                    ws.cell(row=r_out, column=j, value=None if v == "" else v)
            else:
                # derivados: Volumen_i y m3xpersona_i como número si se puede; Resultado_i/Estado texto
                if h.startswith("Volumen_") or h.startswith("m3xpersona_"):
                    v2 = _to_excel_number(v)
                    ws.cell(row=r_out, column=j, value=v2 if v2 != "" else None)
                else:
                    ws.cell(row=r_out, column=j, value=None if v == "" else v)

        r_out += 1

    # ancho cómodo de columnas
    for j, h in enumerate(headers, start=1):
        ws.column_dimensions[get_column_letter(j)].width = max(12, min(50, len(h) + 2))

    wb.save(out_xlsx)



def main():
    # === RUTAS EN DURO (EDITA AQUÍ) ===
    input_dir = Path(r"C:\Users\Quantum-Malloco\Desktop\Ventilacion\Casos de prueba 3\Normalizados\procesadas")
    out_xlsx  = Path(r"C:\Users\Quantum-Malloco\Desktop\Ventilacion\resultados_unificados.xlsx")
    # ===================================

    if not input_dir.exists() or not input_dir.is_dir():
        print(f"Carpeta no encontrada: {input_dir}")
        sys.exit(1)

    files = discover_files(input_dir)
    if not files:
        print("No se encontraron archivos .xlsx/.xlsm")
        sys.exit(1)

    rows = []
    total = len(files)
    for idx, p in enumerate(files, start=1):
        print(f"[{idx}/{total}] Procesando: {p}")
        res = process_file(p)
        if "__error__" in res:
            print(f"    → ERROR: {res['__error__']}")
        else:
            print(f"    → OK: hoja='{res['__sheet__']}', hits={res['__hits__']}")
        rows.append(res)


    # ⬇️ Exporta a Excel en vez de CSV
    write_xlsx(rows, out_xlsx, SCHEMA, DERIVED_FIELDS)

    print(f"Listo. Archivos procesados: {len(files)}")
    print(f"Excel unificado: {out_xlsx}")

if __name__ == "__main__":
    main()
