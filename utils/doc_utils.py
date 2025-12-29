#!/usr/bin/env python3
import logging
import os
import re
import unicodedata
from collections import OrderedDict
from datetime import date, datetime
from io import BytesIO
from pathlib import Path
from typing import Iterable, List, Optional

import pandas as pd
from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import  WD_ROW_HEIGHT_RULE

from docx.oxml.ns import qn
from docx.shared import Cm, Inches
from natsort import natsorted
from pythermalcomfort.models import pmv_ppd_iso

from domain.confort.recommendations import generar_recomendaciones
from utils.doc_styles import (
    agregar_titulo_y_codigo,
    configurar_encabezado,
    look_informe,
    set_vertical_alignment,
)
from utils.media import generate_qr_code
from utils.report_data import (
    formatear_fecha,
    ftemp,
    format_columns,
    format_decimal,
    interpret_pmv,
    procesar_areas,
    todas_cumplen_m3,
)
from utils.word_tables import (
    add_row,
    add_summary_row,
    add_table_with_rows,
    agregar_contenido,
    agregar_medidas_por_accion,
    crear_tabla_recomendaciones,
    format_row,
    set_bold_cell_text,
    set_column_width,
)

# Configuración básica del logging
# logging.basicConfig(level=logging.INFO)
logging.basicConfig(level=logging.DEBUG)

# -----------------------------------------------
# FUNCIONES AUXILIARES PARA EL DOCUMENTO WORD
# -----------------------------------------------

AREA_IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp"}


def center_cell(cell):
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    # Centrar todos los párrafos dentro de la celda
    for p in cell.paragraphs:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER


def join_with_and(items):
    """
    Une los elementos de la lista usando la conjunción "y" para el último elemento.
    Ejemplos:
      - ['Oficina']             => "Oficina"
      - ['Oficina', 'Sala']     => "Oficina y Sala"
      - ['Oficina', 'Sala', 'Bodega'] => "Oficina, Sala y Bodega"
    """
    items = list(items)  # Convertir a lista para evitar ambigüedades
    if not items:
        return ""
    elif len(items) == 1:
        return items[0]
    elif len(items) == 2:
        return f"{items[0]} y {items[1]}"
    else:
        return ", ".join(items[:-1]) + " y " + items[-1]


def _slugify(value: object) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    normalized = unicodedata.normalize("NFD", text)
    ascii_text = normalized.encode("ascii", "ignore").decode("ascii")
    slug = re.sub(r"[^0-9a-zA-Z]+", "-", ascii_text).strip("-")
    return slug.lower()


def _listar_imagenes_area(
    area_id: Optional[object],
    nombre_area: Optional[str],
    visita_id: Optional[object] = None,
) -> List[Path]:
    base_dir = Path("imagenes_pdf") / "areas"
    if not base_dir.exists():
        return []

    candidatos: List[str] = []
    for value in (area_id, nombre_area):
        texto = str(value or "").strip()
        if texto:
            candidatos.append(texto)

    if not candidatos:
        return []

    slug_candidatos = {_slugify(valor) for valor in candidatos if valor}
    imagenes: List[Path] = []
    rutas_visitadas = set()

    def _agregar_imagenes(desde: Path) -> None:
        try:
            archivos = [
                archivo
                for archivo in desde.iterdir()
                if archivo.is_file() and archivo.suffix.lower() in AREA_IMAGE_EXTENSIONS
            ]
        except FileNotFoundError:
            return

        for archivo in natsorted(archivos, key=lambda p: p.name):
            if archivo in rutas_visitadas:
                continue
            rutas_visitadas.add(archivo)
            imagenes.append(archivo)

    visitas_candidatas: List[Path] = []

    if visita_id is not None:
        texto = str(visita_id).strip()
        if texto:
            visita_dir = base_dir / texto
            if visita_dir.exists() and visita_dir.is_dir():
                visitas_candidatas.append(visita_dir)

    if not visitas_candidatas:
        visitas_candidatas.append(base_dir)

    for visita_dir in visitas_candidatas:
        try:
            carpetas = [ruta for ruta in visita_dir.iterdir() if ruta.is_dir()]
        except FileNotFoundError:
            continue

        for carpeta in carpetas:
            nombre_carpeta = carpeta.name
            slug_carpeta = _slugify(nombre_carpeta)
            if nombre_carpeta in candidatos or slug_carpeta in slug_candidatos:
                _agregar_imagenes(carpeta)
                continue

            try:
                subcarpetas = [ruta for ruta in carpeta.iterdir() if ruta.is_dir()]
            except FileNotFoundError:
                continue

            for subcarpeta in subcarpetas:
                nombre_subcarpeta = subcarpeta.name
                slug_subcarpeta = _slugify(nombre_subcarpeta)
                if nombre_subcarpeta in candidatos or slug_subcarpeta in slug_candidatos:
                    _agregar_imagenes(subcarpeta)

    return imagenes


def _formatear_observaciones_general(group: pd.DataFrame, columnas: Iterable[str]) -> str:
    observaciones: List[str] = []
    for columna in columnas:
        if columna not in group.columns:
            continue
        valores = group[columna].dropna().astype(str).str.strip()
        for valor in valores:
            if not valor or valor.lower() == "nan":
                continue
            observaciones.append(valor)

    if not observaciones:
        return "Sin observaciones registradas."

    return "\n".join(OrderedDict.fromkeys(observaciones))


def _agregar_seccion_anexos(doc: Document, informacion_areas: List[dict]):
    doc.add_heading("Anexos", level=2)
    tabla = doc.add_table(rows=1, cols=3)
    tabla.style = "Table Grid"

    encabezados = ["Área", "Observaciones", "Registro fotográfico"]
    for indice, texto in enumerate(encabezados):
        tabla.rows[0].cells[indice].text = texto

    format_row(tabla.rows[0])
    set_column_width(tabla, 0, Cm(3))
    set_column_width(tabla, 1, Cm(5))
    set_column_width(tabla, 2, Cm(8.5))

    if not informacion_areas:
        row_cells = tabla.add_row().cells
        row_cells[0].text = "Sin información disponible"
        merged = row_cells[0].merge(row_cells[1]).merge(row_cells[2])
        for paragraph in merged.paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        return

    for area in informacion_areas:
        row_cells = tabla.add_row().cells
        row_cells[0].text = area.get("nombre", "")
        row_cells[1].text = area.get("observaciones", "")

        imagenes = area.get("imagenes") or []
        fotos_cell = row_cells[2]
        fotos_cell.text = ""

        if imagenes:
            for ruta in imagenes:
                run = fotos_cell.paragraphs[0].add_run()
                try:
                    run.add_picture(str(ruta), width=Cm(8.5))
                except Exception:
                    continue
                run.add_break()
        else:
            fotos_cell.text = "Sin registro fotográfico."


def _agregar_anexo_equipos(
    doc: Document, df_visitas: pd.DataFrame, df_equipos: pd.DataFrame
) -> None:
    """Agrega el anexo de instrumentos utilizados en la evaluación.

    Este bloque replica la estructura utilizada en el informe de confort térmico
    para listar los equipos de medición y sus respaldos.
    """

    if df_visitas.empty or df_equipos.empty:
        doc.add_paragraph("No se encontró información de la visita o de los equipos.")
        return

    row_visita = df_visitas.iloc[0]

    codigos_en_uso = []

    # Los equipos pueden estar registrados en la evaluación de confort o en la
    # tabla específica de ventilación (ev_ventilacion). Se prioriza esta
    # última para los informes de ventilación.
    for key in ("evv_equipo_temp", "evv_equipo_vel_air", "equipo_temp", "equipo_vel_air"):
        valor = str(row_visita.get(key, "")).strip()
        print("INSTRUMENTO", valor)
        if valor:
            codigos_en_uso.append(valor)
            print("INSTRUMENTOS", codigos_en_uso)


    df_equipos = df_equipos.copy()
    print("EQUIPOS", df_equipos)
    if not df_equipos.empty and "id_equipo" in df_equipos.columns:
        df_equipos["id_equipo"] = df_equipos["id_equipo"].astype(str)

    df_equipos_filtrado = df_equipos[df_equipos["id_equipo"].isin(codigos_en_uso)]

    field_mapping = {
        "nombre_equipo": "Tipo de equipo",
        "cod_equipo": "Código",
        "n_serie_equipo": "Número de serie",
        "marca_equipo": "Marca",
        "modelo_equipo": "Modelo",
        "fecha_calibracion": "Última calibración",
        "prox_calibracion": "Próxima calibración",
        "empresa_certificadora": "Empresa certificadora",
        "num_certificado": "Número de certificado",
        "url_certificado": "Respaldo certificado",
    }

    if df_equipos_filtrado.empty:
        instrumentos_manuales = []
        for indice in ("1", "2"):
            instrumento = {
                "nombre_equipo": row_visita.get(f"instru_nombre_{indice}", ""),
                "marca_equipo": row_visita.get(f"instru_marca_{indice}", ""),
                "modelo_equipo": row_visita.get(f"instru_modelo_{indice}", ""),
                "n_serie_equipo": row_visita.get(f"instru_nserie_{indice}", ""),
                "num_certificado": row_visita.get(f"instru_ncertificado_{indice}", ""),
            }

            if any(str(valor).strip() for valor in instrumento.values()):
                instrumentos_manuales.append(instrumento)

        if instrumentos_manuales:
            field_mapping_manual = {
                "nombre_equipo": "Instrumento",
                "marca_equipo": "Marca",
                "modelo_equipo": "Modelo",
                "n_serie_equipo": "Número de serie",
                "num_certificado": "Número de certificado",
            }

            for instrumento in instrumentos_manuales:
                tabla_equipo = doc.add_table(rows=len(field_mapping_manual), cols=2)
                tabla_equipo.style = "Table Grid"
                for row_num, (key, display_name) in enumerate(
                    field_mapping_manual.items()
                ):
                    tabla_equipo.rows[row_num].cells[0].text = display_name
                    tabla_equipo.rows[row_num].cells[1].text = str(
                        instrumento.get(key, "")
                    )

                set_column_width(tabla_equipo, 0, Cm(3.5))
                set_column_width(tabla_equipo, 1, Cm(13.5))
                doc.add_paragraph("")
            return

        doc.add_paragraph(
            "No se encontró información de equipos de medición relacionados con la visita."
        )
        return

    for _, row_eq in df_equipos_filtrado.iterrows():
        tabla_equipo = doc.add_table(rows=len(field_mapping), cols=2)
        tabla_equipo.style = "Table Grid"
        for row_num, (key, display_name) in enumerate(field_mapping.items()):
            tabla_equipo.rows[row_num].cells[0].text = display_name
            if key == "url_certificado":
                url = str(row_eq.get(key, ""))
                if url.strip():
                    qr_img = generate_qr_code(url)
                    cell = tabla_equipo.rows[row_num].cells[1]
                    cell.text = ""
                    run = cell.paragraphs[0].add_run()
                    run.add_break()
                    run.add_picture(qr_img, width=Inches(1))
                    run.add_break()
                else:
                    tabla_equipo.rows[row_num].cells[1].text = ""
            else:
                tabla_equipo.rows[row_num].cells[1].text = str(row_eq.get(key, ""))

        set_column_width(tabla_equipo, 0, Cm(3.5))
        set_column_width(tabla_equipo, 1, Cm(13.5))
        doc.add_paragraph("")

    for row_eq in df_equipos_filtrado.itertuples():
        id_equipo = str(row_eq.id_equipo)
        img_dir = os.path.join("imagenes_pdf", id_equipo)

        try:
            if os.path.exists(img_dir) and os.path.isdir(img_dir):
                imagenes = natsorted(
                    [
                        os.path.join(img_dir, f)
                        for f in os.listdir(img_dir)
                        if f.lower().endswith((".png", ".jpg", ".jpeg"))
                    ]
                )

                for img_path in imagenes:
                    doc.add_picture(img_path, width=Cm(17))
            #else:
                #doc.add_paragraph(
                #    f"No se encontraron imágenes para el equipo {id_equipo}"
                #)
        except Exception as exc:
            doc.add_paragraph(
                f"Error al cargar imágenes para equipo {id_equipo}: {exc}"
            )
def agregar_medidas_correctivas(doc, df_mediciones, areas_no_cumplen):
    # 1. Medidas Ingenieriles (solo para áreas no conformes)
    medidas_ingenieriles = []
    if areas_no_cumplen:
        for area in areas_no_cumplen:
            grupo = df_mediciones[df_mediciones['nombre_area'] == area]
            avg_params = {
                'tdb': grupo['t_bul_seco'].mean(),
                'tr': grupo['t_globo'].mean(),
                'vr': grupo['vel_air'].mean(),
                'rh': grupo['hum_rel'].mean(),
                'met': grupo['met'].mean(),
                'clo': grupo['clo'].mean(),
            }

            try:
                results = pmv_ppd_iso(
                    tdb=avg_params['tdb'],
                    tr=avg_params['tr'],
                    vr=avg_params['vr'],
                    rh=avg_params['rh'],
                    met=avg_params['met'],
                    clo=avg_params['clo'],
                    model="7730-2005",
                    limit_inputs=False,
                    round_output=True
                )
            except Exception as e:
                logging.error("Error al calcular pmv_ppd_iso para el área %s: %s", area, e)
                results = None

            if results is not None:
                if isinstance(results, dict):
                    avg_ppd = float(results.get("ppd", 0))
                    avg_pmv = float(results.get("pmv", 0))
                elif hasattr(results, "ppd") and hasattr(results, "pmv"):
                    avg_ppd = float(results.ppd)
                    avg_pmv = float(results.pmv)
                else:
                    avg_ppd = 0
                    avg_pmv = 0
            else:
                avg_ppd = 0
                avg_pmv = 0

            avg_params.update({
                'pmv': avg_pmv,
                'ppd': avg_ppd
            })

            recs = generar_recomendaciones(
                pmv=avg_params['pmv'],
                tdb_initial=avg_params['tdb'],
                tr_initial=avg_params['tr'],
                vr=avg_params['vr'],
                rh=avg_params['rh'],
                met=avg_params['met'],
                clo=avg_params['clo']
            )

            for rec in recs:
                if rec['tipo'] in ['ventilacion', 'enfriamiento', 'aislamiento', 'calefaccion']:
                    medidas_ingenieriles.append({
                        # 'tipo_medida': rec['categoria'],
                        'areas': [area],
                        'acciones': rec['acciones'],
                        'plazo': rec['plazo']
                    })

    # 2. Medidas Administrativas (siempre se incluyen)

    ##CAMBIO
    medidas_administrativas = [
        {
            # 'tipo_medida': 'Comunicación',
            'areas': ['Todas las áreas evaluadas'],
            'acciones': [
                "- Informar a cada persona trabajadora acerca de los riesgos que entrañan sus labores, de las medidas preventivas, de los métodos y/o procedimientos de trabajo correctos, acorde a lo identificado por la empresa. Además de lo señalado previamente, la entidad empleadora deberá informar de manera oportuna y adecuada el resultado del presente informe técnico.",
                "- Realizar capacitaciones (teóricas/prácticas) periódicas en prevención de riesgos laborales, con la finalidad de garantizar el aprendizaje efectivo y eficaz, dejando registro de dichas capacitaciones y evaluaciones. En el marco de los artículos 15° y 16° del Párrafo IV del D.S 44 “Aprueba nuevo reglamento sobre gestión preventiva de los riesgos laborales para un entorno de trabajo seguro y saludable."
            ],
            'plazo': '30 días desde la recepción del presente informe técnico'
        },
        {
            # 'tipo_medida': 'Mantenimiento',
            'areas': areas_no_cumplen if areas_no_cumplen else ['Todas'],
            'acciones': [
                "- Consultar con el proveedor el óptimo uso del equipo por ejemplo: periodicidad de suministrar agua helada, hielo o implemento refrigerante autorizado para el equipamiento adquirido, con el fin de estar constantemente enfriando durante toda la jornada laboral el área, especialmente en periodo de mayor temperaturas o época estival."
                "- Realizar mantención preventiva en los equipos de climatización, con el fin de identificar desgastes y prevenir fallas. Se debe seguir un cronograma establecido y registrar cada intervención."
                "- Ejecutar reparaciones en equipos de climatización al detectar fallas en su funcionamiento, restableciendo su operatividad de manera oportuna y registrando las acciones realizadas."
                "- Implementar un monitoreo continuo de los parámetros de confort térmico entre 23 a 26°C en verano manteniendo un registro sistemático de las mediciones y ajustes efectuados."
                "- Llevar una Bitácora o Registro de la actividad en lo referido al uso de los equipos."
            ],
            'plazo': '[__] dias desde la recepción del presente informe técnico'
        }
    ]

    # Agregar secciones al documento
    doc.add_heading("4.1 Medidas de Carácter Técnico", level=3)
    if medidas_ingenieriles:
        crear_tabla_recomendaciones(doc, "Ingenieril", medidas_ingenieriles)
    else:
        doc.add_paragraph("No se requieren medidas técnicas para este caso.")
    ##FINCAMBIO

    doc.add_paragraph()
    doc.add_heading("4.2 Medidas de Carácter Administrativo", level=3)
    crear_tabla_recomendaciones(doc, "Administrativa", medidas_administrativas)


# Función para generar texto de áreas cumplen/no cumplen
def generar_texto_areas(areas, tipo):
    if not areas:
        return ""
    areas_formateadas = join_with_and(areas)
    if tipo == "cumplen":
        return f"las áreas {areas_formateadas} cumplen con el estándar de confort térmico"
    else:
        return f"las áreas {areas_formateadas} no cumplen con el estándar"


def _formatear_texto_areas(areas, singular_template, plural_template):
    if not areas:
        return None
    if len(areas) == 1:
        return singular_template.format(areas[0])
    return plural_template.format(join_with_and(areas))


def redactar_conclusiones(doc, nombre_ct, areas_cumplen, areas_no_cumplen):
    cumplen_text = _formatear_texto_areas(areas_cumplen, "el área de {}", "las áreas {}")
    no_cumplen_text = _formatear_texto_areas(areas_no_cumplen, "el área {}", "las áreas {}")

    if cumplen_text and no_cumplen_text:
        doc.add_paragraph(
            f"Efectuadas mediciones de confort térmico en el local {nombre_ct}, se concluye que {cumplen_text} "
            f"{'cumple' if len(areas_cumplen) == 1 else 'cumplen'} con el estándar de confort térmico, establecido mediante la metodología de Fanger. "
            f"Esto significa que, al registrarse un PMV entre -1 y +1, el PPD, que expresa el porcentaje de personas que experimentan disconfort con la temperatura, "
            f"resulta inferior al 25%. Por ello, se recomienda mantener y/o mejorar las condiciones actuales o similares."
        )
        doc.add_paragraph()
        doc.add_paragraph(
            f"Respecto de {no_cumplen_text} {'NO cumple' if len(areas_no_cumplen) == 1 else 'NO cumplen'} con el mismo estándar, ya que el PMV se encuentra fuera del rango de -1 a +1, "
            f"lo que implica que el PPD resulta superior al 25%. Por lo tanto, se deben adoptar las medidas correctivas para alcanzar los estándares en la referencia técnica."
        )
    elif cumplen_text:
        doc.add_paragraph(
            f"Efectuadas mediciones de confort térmico en el local {nombre_ct}, se concluye que {cumplen_text} "
            f"{'cumple' if len(areas_cumplen) == 1 else 'cumplen'} con el estándar de confort térmico, establecido mediante la metodología de Fanger. "
            f"Esto significa que, al registrarse un PMV entre -1 y +1, el PPD, que expresa el porcentaje de personas que experimentan disconfort con la temperatura, "
            f"resulta inferior al 25%. Por ello, se recomienda mantener y/o mejorar las condiciones actuales o similares."
        )
    elif no_cumplen_text:
        doc.add_paragraph(
            f"Efectuadas mediciones de confort térmico en el local {nombre_ct}, se concluye que {no_cumplen_text} "
            f"{'NO cumple' if len(areas_no_cumplen) == 1 else 'NO cumplen'} con el estándar de confort térmico, establecido mediante la metodología de Fanger. "
            f"Esto significa que, al registrarse un PMV fuera del rango de -1 a +1, el PPD, que expresa el porcentaje de personas que experimentan disconfort con la temperatura, "
            f"resulta superior al 25%. Por lo tanto, se deben adoptar las medidas correctivas para alcanzar los estándares en la referencia técnica"
        )
    else:
        doc.add_paragraph(
            "No se encontró información suficiente en las mediciones para emitir una conclusión sobre el confort térmico."
        )


######

# -----------------------------------------------
# FUNCIÓN PARA GENERAR EL DOCUMENTO WORD
# -----------------------------------------------
def generar_informe_en_word(df_centros, df_visitas, df_mediciones, df_equipos) -> BytesIO:
    """
    Genera el informe en Word utilizando:
      - df_centros: información del centro de trabajo (tabla higiene_Centros_Trabajo)
      - df_visitas: información de visitas (tabla higiene_Visitas); se selecciona la visita más reciente.
      - df_mediciones: mediciones asociadas a la visita (tabla higiene_Mediciones)
      - df_equipos: información de equipos de medición (tabla higiene_Equipos_Medicion)
    """

    format_columns(
        df_visitas, ['nombre_personal_visita', 'consultor_name_complete'], mode="title"
    )
    format_columns(df_visitas, 'cargo_personal_visita', mode="capitalize")
    format_columns(df_mediciones, ['nombre_area', 'sector_especifico', 'puesto_trabajo'], mode="capitalize")

    doc = Document()
    look_informe(doc)
    set_vertical_alignment(doc, section_index=0, alignment='top')

    configurar_encabezado(doc)
    agregar_titulo_y_codigo(doc, "INFORME EVALUACIÓN CONFORT TÉRMICO", "CODIGO: [COMPLETAR]")

    # A partir de aquí, el contenido se alineará a la izquierda (valor por defecto)
    paragraph = doc.add_heading("1. Antecedentes", level=2)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    doc.add_paragraph(
        "Por solicitud del área de prevención de la empresa, se realiza evaluación de Confort Térmico, para determinar condición en que se encuentran los trabajadores que se desempeñan en áreas o sectores de trabajo, con el fin de que la empresa pueda adoptar oportuna y eficazmente medidas que permitan mejorar las condiciones evaluadas, según corresponda.")
    doc.add_paragraph()

    # -------------------------------
    # 1) IDENTIFICACIÓN ACTIVIDAD
    # -------------------------------

    table_empresa = None
    table_centro = None
    table_visita = None

    if not df_centros.empty and not df_visitas.empty:
        row_centro = df_centros.iloc[0]
        row_visita = df_visitas.iloc[0]

        add_table_with_rows(
            doc,
            "1.1 Información empresa",
            [
                ("Razón Social", row_centro.get('razon_social', '').lower().title()),
                ("RUT", row_centro.get('rut', '')),
                (
                    "CIIU",
                    row_centro.get(
                        'CIIU',
                        '521111. Grandes establecimientos (venta de alimentos); hipermercados'
                    ),
                ),
            ],
        )

        doc.add_paragraph()

        add_table_with_rows(
            doc,
            "1.2 Información centro de trabajo",
            [
                ("CUV/CECO/Código IST", row_centro.get('cuv', '')),
                ("Nombre de Local", row_centro.get('nombre_ct', '').lower().title()),
                ("Dirección", row_centro.get('direccion_ct', '')),
                ("Comuna", row_centro.get('comuna_ct', '')),
                ("Región", row_centro.get('region_ct', '')),
            ],
        )

        doc.add_paragraph()

        temperatura_exterior = ftemp(row_visita.get('temperatura_dia', '')) + "°C"
        add_table_with_rows(
            doc,
            "1.3 Información de la visita",
            [
                ("Motivo de la actividad", "Programa de trabajo"),
                ("Fecha actividad de terreno", formatear_fecha(row_visita.get('fecha_visita', ''))),
                ("Hora actividad de terreno", row_visita.get('hora_visita', '')),
                ("Temperatura ambiental exterior", temperatura_exterior),
                ("Fecha emisión informe", "[COMPLETAR]"),
                (
                    "Profesional consultor/a de IST",
                    row_visita.get('consultor_name_complete', '').lower().title(),
                ),
                ("Acompañante empresa", row_visita.get('nombre_personal_visita', '').lower().title()),
                (
                    "Cargo de la persona que acompaña visita",
                    row_visita.get('cargo_personal_visita', '').lower().title(),
                ),
                ("Revisor del informe", "Rodrigo Novoa Miranda"),
                ("Jefatura responsable IST", "[COMPLETAR]"),
                ("Destinatario informe", "[COMPLETAR]"),
            ],
        )

    else:
        add_table_with_rows(
            doc,
            "Información de empresa/centro",
            [("Resultado", "No se encontró información para este CUV.")],
        )

    # Encabezado principal del contenido: Metodología
    paragraph = doc.add_heading("2. Metodología", level=2)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # Primer párrafo
    doc.add_paragraph(
        "El concepto de “confort térmico” describe el estado mental de una persona en términos de percibir un ambiente demasiado caluroso o demasiado frío. "
        "Por otro lado, también se puede definir como una manifestación subjetiva de conformidad o satisfacción entre el trabajador con el ambiente térmico existente."
    )

    # Segundo párrafo con palabras en negrita
    doc.add_paragraph()
    p = doc.add_paragraph(
        "El presente informe utiliza la metodología de FANGER para evaluación de confort térmico en espacios interiores de acuerdo a la Nota técnica N°47 del Instituto de Salud Pública. "
        "De esta forma, los diferentes puestos de trabajo son evaluados y calificados en cada caso como "
    )
    p.add_run('"Cumple"').bold = True
    p.add_run(" o ")
    p.add_run('"No cumple"').bold = True

    # Tercer párrafo con palabras en negrita
    doc.add_paragraph()
    p2 = doc.add_paragraph(
        "Los alcances de las calificaciones específicas para cada área o sector evaluado, corresponde al cumplimiento del Voto Medio Estimado "
    )
    p2.add_run("PMV").bold = True
    p2.add_run(
        " (Predicted Mean Vote), equivalente a una condición media deseable que indican la sensación térmica media de un entorno y ")
    p2.add_run("PPD").bold = True
    p2.add_run(
        " (Predicted Percentage Dissatisfied) correspondiente al porcentaje de personas que sentirán algún grado de disconfort en un ambiente de trabajo evaluado.")

    # -------------------------------
    # Resultados de mediciones y evaluación
    # -------------------------------
    # Crear listas para las áreas que cumplen y no cumplen
    areas_cumplen = []
    areas_no_cumplen = []

    paragraph = doc.add_heading("3. Resultados de las mediciones y evaluación", level=2)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    def generar_tabla_resumen(doc, df_mediciones):
        # Verificamos si el DataFrame tiene datos
        if df_mediciones.empty:
            doc.add_paragraph("No se encontraron mediciones para detallar.")
            return

        # Definimos las columnas en el orden requerido
        columnas_resumen = [
            "Área",
            "Estándar confortabilidad",
            "Puesto de trabajo",
            "Temp. bulbo seco (°C)",
            "Temp. globo (°C)",
            "Humedad relativa (%)",
            "Velocidad del aire (m/s)",
            "PPD (%)",
            "PMV"
        ]

        # Creamos la tabla con tantas columnas como la lista anterior
        tabla_resumen = doc.add_table(rows=1, cols=len(columnas_resumen))
        tabla_resumen.style = 'Table Grid'

        # Encabezados
        hdr_cells = tabla_resumen.rows[0].cells
        for idx, col_name in enumerate(columnas_resumen):
            hdr_cells[idx].text = col_name

        # Agrupamos por área
        grouped = df_mediciones.groupby("nombre_area")

        # Recorremos cada grupo (cada área)
        for area, group in grouped:
            # Si hay múltiples mediciones en el área, calculamos promedio;
            # si solo hay una, usamos directamente esa.
            if len(group) > 1:
                # Calculamos promedios
                avg_t_bul = group["t_bul_seco"].astype(float).mean()
                avg_t_globo = group["t_globo"].astype(float).mean()
                avg_hum = group["hum_rel"].astype(float).mean()
                avg_vel = group["vel_air"].astype(float).mean()
                avg_met = group["met"].astype(float).mean()
                avg_clo = group["clo"].astype(float).mean()

                # Calcular pmv/ppd a partir de la función pmv_ppd_iso
                # (ajusta según tu propia lógica)
                avg_ppd = 0
                avg_pmv = 0
                try:
                    resultados = pmv_ppd_iso(
                        tdb=avg_t_bul,
                        tr=avg_t_globo,
                        vr=avg_vel,
                        rh=avg_hum,
                        met=avg_met,
                        clo=avg_clo,
                        model="7730-2005",
                        limit_inputs=False,
                        round_output=True,
                    )
                    if isinstance(resultados, dict):
                        avg_ppd = float(resultados.get("ppd", 0))
                        avg_pmv = float(resultados.get("pmv", 0))
                    else:
                        avg_ppd = float(getattr(resultados, "ppd", 0))
                        avg_pmv = float(getattr(resultados, "pmv", 0))
                except Exception as e:
                    logging.error("Error al calcular pmv_ppd_iso para el área %s: %s", area, e)

                analisis = interpret_pmv(avg_pmv)
                valores_unicos = list(OrderedDict.fromkeys(group["puesto_trabajo"].dropna()))
                puesto_trabajo = "\n".join(str(x) for x in valores_unicos)
                metrics = {
                    "t_bul": avg_t_bul,
                    "t_globo": avg_t_globo,
                    "hum": avg_hum,
                    "vel": avg_vel,
                    "ppd": avg_ppd,
                    "pmv": avg_pmv,
                }
                add_summary_row(tabla_resumen, str(area), analisis, puesto_trabajo, metrics)

            else:
                # Solo hay una medición en el área, la usamos directamente
                row = group.iloc[0]
                analisis = str(row.get("resultado_medicion", "")).upper()
                puesto_trabajo = str(row.get("puesto_trabajo", ""))
                metrics = {
                    "t_bul": row.get("t_bul_seco"),
                    "t_globo": row.get("t_globo"),
                    "hum": row.get("hum_rel"),
                    "vel": row.get("vel_air"),
                    "ppd": row.get("ppd"),
                    "pmv": row.get("pmv"),
                }
                add_summary_row(tabla_resumen, str(area), analisis, puesto_trabajo, metrics)

        # Opcional: Ajustar anchos de columna si lo deseas
        set_column_width(tabla_resumen, 0, Cm(3))
        set_column_width(tabla_resumen, 1, Cm(3))
        set_column_width(tabla_resumen, 2, Cm(3))
        set_column_width(tabla_resumen, 3, Cm(1.5))
        set_column_width(tabla_resumen, 4, Cm(1.5))
        set_column_width(tabla_resumen, 5, Cm(1.5))
        set_column_width(tabla_resumen, 6, Cm(1.5))
        set_column_width(tabla_resumen, 7, Cm(1.5))
        set_column_width(tabla_resumen, 8, Cm(1.5))

        # Formato de la fila de encabezado (opcional)
        format_row(tabla_resumen.rows[0])

    generar_tabla_resumen(doc, df_mediciones)

    # Procesar áreas antes del resumen
    areas_cumplen, areas_no_cumplen = procesar_areas(df_mediciones)

    # Encabezado principal del contenido: Conclusiones
    doc.add_paragraph()
    nombre_ct = df_centros.iloc[0].get("nombre_ct", "") if not df_centros.empty else ""
    redactar_conclusiones(doc, nombre_ct, areas_cumplen, areas_no_cumplen)

    '''
    if not df_centros.empty:
        row_centro = df_centros.iloc[0]
        nombre_ct = row_centro.get("nombre_ct", "")
        # Construir las cadenas de texto según las áreas que cumplen y las que no cumplen

        if areas_cumplen:
            if len(areas_cumplen) == 1:
                # Forma singular
                cumplen_text = f"el área de {areas_cumplen[0]}"
            else:
                # Forma plural, usando la función para unir con "y"
                cumplen_text = f"las áreas {join_with_and(areas_cumplen)}"
        else:
            cumplen_text = None  # O dejarlo en cadena vacía, según convenga

        if areas_no_cumplen:
            if len(areas_no_cumplen) == 1:
                no_cumplen_text = f"el área {areas_no_cumplen[0]}"
            else:
                no_cumplen_text = f"las áreas {join_with_and(areas_no_cumplen)}"
        else:
            no_cumplen_text = None

        # Generar la redacción final de las conclusiones
        if cumplen_text and no_cumplen_text:
            # Caso 2: Existen áreas que cumplen y áreas que no cumplen.
            doc.add_paragraph(
                f"Efectuadas mediciones de confort térmico en el local {nombre_ct}, es posible concluir que {cumplen_text} "
                f"{'cumple' if len(areas_cumplen) == 1 else 'cumplen'} con el estándar de confort térmico, por lo que se recomienda mantener las condiciones actuales o similares."
            )
            doc.add_paragraph(
                f"Respecto a {no_cumplen_text} que NO {'cumple' if len(areas_no_cumplen) == 1 else 'cumplen'} con el estándar, se deben adoptar las medidas prescritas a continuación para corregir las condiciones."
            )
        elif cumplen_text and not no_cumplen_text:
            # Caso 1: Sólo existen áreas que cumplen.
            doc.add_paragraph(
                f"Efectuadas mediciones de confort térmico en el local {nombre_ct}, se concluye que {cumplen_text} "
                f"{'cumple' if len(areas_cumplen) == 1 else 'cumplen'} con el estándar de confort térmico, por lo que se recomienda mantener las condiciones actuales o similares."
            )
        elif no_cumplen_text and not cumplen_text:
            # Caso 3: Sólo existen áreas que no cumplen.
            doc.add_paragraph(
                f"Efectuadas mediciones de confort térmico en el local {nombre_ct}, se concluye que {no_cumplen_text} "
                f"{'NO cumple' if len(areas_no_cumplen) == 1 else 'NO cumplen'} con el estándar de confort térmico, por lo que se deben adoptar las medidas prescritas a continuación para corregir las condiciones."
            )
        else:
            # En caso de que no haya información suficiente
            doc.add_paragraph(
                "No se encontró información suficiente en las mediciones para emitir una conclusión sobre el confort térmico."
            )

    '''

    # -------------------------------
    # 4) MEDIDAS CORRECTIVAS
    # -------------------------------

    paragraph = doc.add_heading("4. Prescripción de medidas", level=2)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # Asumiendo que row_centro ya está definido y contiene la información de la empresa:
    razon_social = row_centro.get('razon_social', 'RENDIC HERMANOS S.A.')

    ##CAMBIO
    # Luego, en el cuerpo del documento:
    # doc.add_paragraph(
    #    "Conforme al artículo 68 de la Ley N° 16.744, la implementación de las medidas prescritas por este organismo "
    #    "administrador es de carácter obligatoria, por lo que su incumplimiento podrá ser sancionado con el recargo de "
    #    "la cotización adicional diferenciada, sin perjuicio de las demás sanciones que correspondan."
    # )
    # doc.add_paragraph()
    # doc.add_paragraph(
    #    f"No obstante, {razon_social} podrá implementar otras medidas técnicas y/o administrativas equivalentes a las "
    #    "señaladas en el presente informe y que contribuyan a disminuir la exposición de sus trabajadores, debiendo "
    #    "informar a IST, quien evaluará su efectividad una vez implementadas. Adicionalmente, en el caso de que las áreas "
    #    "de trabajo sean operadas por contratistas, el mandante debe informar obligatoriamente a todos sus contratistas los "
    #    "riesgos a los que están expuestos."
    # )
    # doc.add_paragraph()
    ##FINCAMBIO

    doc.add_paragraph(
        "Acorde a las condiciones existentes al momento de las mediciones, al resultado de las mismas y a las conclusiones "
        "obtenidas, se establecen las siguientes medidas de control:"
    )

    agregar_medidas_correctivas(doc, df_mediciones, areas_no_cumplen)

    doc.add_paragraph()

    # -------------------------------
    # 5) VIGENCIA DEL INFORME
    # -------------------------------
    # Encabezado principal del contenido: Vigencia del informe
    paragraph = doc.add_heading("5. Vigencia del informe", level=2)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    if not df_visitas.empty:
        doc.add_paragraph(
            "En términos generales, el presente informe tiene validez de 3 años, a excepción que existan cambios en la situación, del tipo ingenieril o administrativo, que presupongan modificación a las condiciones encontradas al momento de la medición, lo cual implicará realizar una nueva evaluación en un plazo menor al señalado.")
        doc.add_paragraph()
        if len(areas_no_cumplen) > 0:
            doc.add_paragraph(
                "Cuando se concreten los cambios indicados, la empresa deberá informar al IST el detalle de los mismos, de forma tal de programar las gestiones a realizar, las que considerarán previamente un informe de Verificación y Control y posteriormente, de acuerdo a sus resultados, la nueva evaluación de higiene ocupacional correspondiente.")
            doc.add_paragraph()
        doc.add_paragraph(
            "Estos resultados de evaluación representan las condiciones existentes del ambiente y lugar de trabajo al momento de realizar las mediciones.")
    doc.add_paragraph()
    doc.add_paragraph()
    doc.add_paragraph()
    doc.add_paragraph()

    consultor_nombre = ""
    consultor_name_complete = ""
    consultor_cargo = ""
    consultor_zonal = ""

    if not df_visitas.empty:
        row_visita = df_visitas.iloc[0]
        consultor_nombre = row_visita.get("consultor_nombre", "")
        consultor_name_complete = (
            row_visita.get("consultor_name_complete") or consultor_nombre
        )
        consultor_cargo = row_visita.get("consultor_cargo", "")
        consultor_zonal = row_visita.get("consultor_zonal", "")

    consultor_ist_limpio = (
        consultor_nombre
        .replace(" ", "-")
        .replace("ñ", "n")
        .replace("Ñ", "N")
        .replace("á", "a")
        .replace("é", "e")
        .replace("í", "i")
        .replace("ó", "o")
        .replace("ú", "u")
    )

    firma_path = os.path.join("imagenes-firma", consultor_ist_limpio + ".png")

    try:
        paragraph = doc.add_paragraph()
        run = paragraph.add_run()
        run.add_picture(firma_path, width=Cm(4))
        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    except FileNotFoundError:
        print(f"La imagen no existe en la ruta especificada: {firma_path}")
    except Exception as e:
        print(f"Ocurrió un error al agregar la imagen: {e}")

    # Agregar párrafo para el consultor, centrado y en negrita
    p_consultor = doc.add_paragraph()
    p_consultor.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_consultor = p_consultor.add_run(consultor_name_complete)
    run_consultor.bold = True

    # Agregar párrafo para la profesión, centrado
    p_profesion = doc.add_paragraph()
    p_profesion.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_profesion = p_profesion.add_run(consultor_cargo)

    # Agregar párrafo para el zonal, centrado
    p_zonal = doc.add_paragraph()
    p_zonal.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_zonal = p_zonal.add_run(consultor_zonal)

    anexos_info: List[dict] = []
    if not df_mediciones.empty and "nombre_area" in df_mediciones.columns:
        obs_cols = [col for col in df_mediciones.columns if col.startswith("obs_")]
        for extra in ("observaciones", "caract_constructivas", "ingreso_salida_aire"):
            if extra in df_mediciones.columns:
                obs_cols.append(extra)

        for area, group in df_mediciones.groupby("nombre_area"):
            area_id = None
            if "area_id" in group.columns:
                disponibles = group["area_id"].dropna()
                if not disponibles.empty:
                    area_id = disponibles.iloc[0]

            visita_id = None
            for columna in ("visita_id", "id_visita"):
                if columna not in group.columns:
                    continue
                disponibles = group[columna].dropna()
                if disponibles.empty:
                    continue
                visita_id = disponibles.iloc[0]
                break

            observaciones = _formatear_observaciones_general(group, obs_cols)
            imagenes = _listar_imagenes_area(area_id, area, visita_id)

            anexos_info.append(
                {
                    "nombre": str(area) if area is not None else "",
                    "observaciones": observaciones,
                    "imagenes": imagenes,
                }
            )

    doc.add_page_break()
    _agregar_seccion_anexos(doc, anexos_info)
    doc.add_paragraph()
    doc.add_heading("Anexo 1. Consideraciones técnicas de la evaluación", level=2)

    # Agrega un párrafo con el título para la tabla
    doc.add_heading("a)     Verificación en terreno de parámetros de los equipos", level=3)

    # Crea la tabla con 4 columnas y aplica un estilo
    table_calib = doc.add_table(rows=0, cols=4)
    table_calib.style = 'Table Grid'

    # Agrega la fila de encabezado
    hdr_cells = table_calib.add_row().cells
    hdr_cells[0].text = "Temperatura"
    hdr_cells[1].text = "Patrón"
    hdr_cells[2].text = "Verificación inicial"
    hdr_cells[3].text = "Verificación final"

    # Si hay datos en df_visitas, agrega las filas para cada equipo
    if not df_visitas.empty:
        # Para TBS
        row_cells = table_calib.add_row().cells
        row_cells[0].text = "TBS"
        row_cells[1].text = ftemp(row_visita.get('patron_tbs', ''))
        row_cells[2].text = ftemp(row_visita.get('ver_tbs_ini', ''))
        row_cells[3].text = ftemp(row_visita.get('ver_tbs_fin', ''))

        # Para TBH
        row_cells = table_calib.add_row().cells
        row_cells[0].text = "TBH"
        row_cells[1].text = ftemp(row_visita.get('patron_tbh', ''))
        row_cells[2].text = ftemp(row_visita.get('ver_tbh_ini', ''))
        row_cells[3].text = ftemp(row_visita.get('ver_tbh_fin', ''))

        # Para TG
        row_cells = table_calib.add_row().cells
        row_cells[0].text = "TG"
        row_cells[1].text = ftemp(row_visita.get('patron_tg', ''))
        row_cells[2].text = ftemp(row_visita.get('ver_tg_ini', ''))
        row_cells[3].text = ftemp(row_visita.get('ver_tg_fin', ''))
    else:
        # Si no hay información, se agrega una fila de error
        row_cells = table_calib.add_row().cells
        row_cells[0].text = "D. Detalles de equipos y calibración"
        row_cells[1].text = "No se encontró información de visita."
        # Se pueden unir las celdas restantes para que el mensaje quede centrado
        merged = row_cells[0].merge(row_cells[1])
        merged = merged.merge(row_cells[2]).merge(row_cells[3])

    # Configurar el ancho de cada columna (opcional)
    set_column_width(table_calib, 0, Cm(4.25))
    set_column_width(table_calib, 1, Cm(4.25))
    set_column_width(table_calib, 2, Cm(4.25))
    set_column_width(table_calib, 3, Cm(4.25))
    format_row(table_calib.rows[0])

    doc.add_paragraph()
    doc.add_heading("b)     Caracterización de vestimenta utilizada y tasa metabólica", level=3)

    doc.add_paragraph(
        "Para efectos del presente informe, se considera el valor de 0,5 Clo, que es equivalente a ropa normal de trabajo y una tasa metabólica (Mets) de 1,1 a 1,2 (valor equivalente a 109 kcal/hrs.) en función de los componentes de la actividad para el área Línea de Cajas o Sala de ventas, y una tasa metabólica de 1,89 Mets (valor equivalente a 187 kcal/hrs.), en función de la profesión, para el área Bodega o Recepción.")

    doc.add_paragraph()
    doc.add_heading("c)     Características generales de las areas evaluadas", level=3)

    # Crear la tabla con 3 columnas: Área, Características constructivas y Condiciones de ventilación
    tabla_caract = doc.add_table(rows=1, cols=3)
    tabla_caract.style = 'Table Grid'

    # Agregar la fila de encabezado
    hdr_cells = tabla_caract.rows[0].cells
    hdr_cells[0].text = "Área"
    hdr_cells[1].text = "Características constructivas"
    hdr_cells[2].text = "Condiciones de ventilación"

    # Cambio caracteristicas constructivas
    # Diccionario actualizado con frases base
    condiciones = {
        "cond_techumbre": "cuenta con techumbre aislante",
        "cond_paredes": "cuenta con paredes aislantes",
        "cond_vantanal": "cuenta con ventanas aislantes",
        "cond_aire_acond": "cuenta con aire acondicionado",
        "cond_ventiladores": "cuenta con ventiladores",
        "cond_inyeccion_extraccion": "cuenta con sistema de inyección/extracción",
        "cond_ventanas": "tiene ventanas abiertas",
        "cond_puertas": "tiene puertas abiertas",
        "cond_otras": "existen otras condiciones generadoras de disconfort térmico"
    }

    grouped = df_mediciones.groupby("nombre_area")
    for area, group in grouped:
        registro = group.iloc[0]

        campos_constructivas = [
            "cond_techumbre", "obs_techumbre", "cond_paredes", "obs_paredes",
            "cond_vantanal", "obs_ventanal", "cond_otras", "obs_otras"
        ]
        campos_aire = [
            "cond_aire_acond", "obs_aire_acond", "cond_ventiladores", "obs_ventiladores",
            "cond_inyeccion_extraccion", "obs_inyeccion_extraccion", "cond_ventanas",
            "cond_puertas", "obs_ventanas", "obs_puertas"
        ]

        def procesar_campos(campos):
            resultado = []
            for campo in campos:
                valor = registro[campo]
                # Se filtra si el valor es nulo o está vacío (o solo espacios)
                if pd.notna(valor) and str(valor).strip() != "":
                    valor_str = str(valor).strip()
                    # Si el campo está en condiciones, es numérico (0 o 1) y se convierte a "No"/"Si"
                    if campo in condiciones:
                        if valor_str == "0":
                            valor_str = "No"
                        elif valor_str == "1":
                            valor_str = "Si"
                        # Se concatena la frase predefinida
                        resultado.append(f"{valor_str} {condiciones[campo]}")
                    else:
                        # Para campos observación, se conserva el texto tal cual
                        resultado.append(valor_str)
            return ", ".join(resultado)

        caract_constructivas = procesar_campos(campos_constructivas)
        ingreso_salida_aire = procesar_campos(campos_aire)

        row_cells = tabla_caract.add_row().cells
        row_cells[0].text = area
        row_cells[1].text = caract_constructivas
        row_cells[2].text = ingreso_salida_aire

    set_column_width(tabla_caract, 0, Cm(3))
    set_column_width(tabla_caract, 1, Cm(7))
    set_column_width(tabla_caract, 2, Cm(7))
    format_row(tabla_caract.rows[0])

    # Salto de página y título del anexo
    doc.add_page_break()
    doc.add_heading("Anexo 2. Instrumentos de medición utilizados", level=2)
    _agregar_anexo_equipos(doc, df_visitas, df_equipos)

    # (Continúa el resto del script si es necesario)

    # -------------------------------
    # Finaliza el documento y lo retorna como BytesIO
    # -------------------------------
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


def _texto_cumplimiento(valor):
    if valor in (1, "1", True):
        return "Cumple"
    if valor in (0, "0", False):
        return "No cumple"
    return ""


def generar_informe_ventilacion_en_word(
    df_centros, df_visitas, df_areas, df_puntos, df_equipos=None
) -> BytesIO:
    """Genera un informe en Word para la evaluación de ventilación."""

    if not isinstance(df_visitas, pd.DataFrame):
        df_visitas = pd.DataFrame(df_visitas or [])
    if not isinstance(df_areas, pd.DataFrame):
        df_areas = pd.DataFrame(df_areas or [])
    if not isinstance(df_puntos, pd.DataFrame):
        df_puntos = pd.DataFrame(df_puntos or [])
    if not isinstance(df_equipos, pd.DataFrame):
        df_equipos = pd.DataFrame(df_equipos or [])

    df_visitas = df_visitas.copy()
    df_areas = df_areas.copy()
    df_puntos = df_puntos.copy()
    df_equipos = df_equipos.copy()

    format_columns(
        df_visitas,
        ["nombre_personal_visita", "consultor_name_complete", "consultor_zonal"],
        mode="title",
    )
    if not df_areas.empty:
        format_columns(df_areas, ["nombre_area", "uso", "ventilacion_tipo", "ventilacion_estado"], mode="title")

    doc = Document()
    look_informe(doc)
    set_vertical_alignment(doc, section_index=0, alignment='top')

    configurar_encabezado(doc)
    agregar_titulo_y_codigo(doc, "INFORME EVALUACIÓN VENTILACIÓN", "CÓDIGO: [COMPLETAR]")
    # Centrar los dos últimos párrafos (título y código)
    for p in doc.paragraphs[-2:]:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER


    doc.add_heading("1. Antecedentes", level=2)
    doc.add_paragraph(
        "Por solicitud del área de prevención de la empresa, se realiza la evaluación de ventilación de acuerdo con lo "
        "establecido en el Artículo N°34 del Decreto Supremo N°594/1999, realizando la evaluación de las condiciones de ventilación en las "
        "áreas o sectores del centro de trabajo, de modo que la organización pueda implementar medidas oportunas que "
        "resguarden la salud de las personas trabajadoras conforme al citado decreto."                                                                                                                  )
    doc.add_paragraph()

    table_empresa = None
    table_centro = None
    table_visita = None

    if not df_centros.empty and not df_visitas.empty:
        row_centro = df_centros.iloc[0]
        row_visita = df_visitas.iloc[0]

        table_empresa = add_table_with_rows(
            doc,
            "1.1 Información empresa",
            [
                ("Razón Social", row_centro.get('razon_social', '').lower().title()),
                ("RUT", row_centro.get('rut', '')),
                ("CIIU", '471100 -  Venta al por menor en comercios de alimentos, bebidas o tabaco (supermercado)'),
            ],
        )
        set_column_width(table_empresa, 0, Cm(7))
        set_column_width(table_empresa, 1, Cm(10))

        doc.add_paragraph()

        table_centro = add_table_with_rows(
            doc,
            "1.2 Información centro de trabajo",
            [
                ("CUV/CECO/Código IST", row_centro.get('cuv', '')),
                ("Nombre de Local", row_centro.get('nombre_ct', '').lower().title()),
                ("Dirección", row_centro.get('direccion_ct', '')),
                ("Comuna", row_centro.get('comuna_ct', '')),
                ("Región", row_centro.get('region_ct', '')),
            ],
        )
        set_column_width(table_centro, 0, Cm(7))
        set_column_width(table_centro, 1, Cm(10))

        doc.add_paragraph()

        motivo = row_visita.get('motivo_evaluacion', '') or "Programa de trabajo"
        table_visita = add_table_with_rows(
            doc,
            "1.3 Información de la visita",
            [
                ("Motivo de la actividad", motivo),
                ("Fecha actividad de terreno", formatear_fecha(row_visita.get('fecha_visita', ''))),
                ("Hora actividad de terreno", row_visita.get('hora_visita', '')),
                (
                    "Profesional consultor/a de IST",
                    row_visita.get('consultor_nombre', '').lower().title(),
                ),
                ("Acompañante empresa", row_visita.get('nombre_personal_visita', '').lower().title()),
                ("Cargo de la persona que acompaña la visita", row_visita.get('cargo_personal_visita', '').lower().title()),
                ("Tipo de evaluación", row_visita.get('tipo_evaluacion', '').lower().title()),
                ("Fecha emisión informe", "[COMPLETAR]"),
                ("Revisor del informe", "Rodrigo Novoa"),
                ("Destinatario informe", "Cristian Fernandez"),
            ],
        )
        set_column_width(table_visita, 0, Cm(7))
        set_column_width(table_visita, 1, Cm(10))

    else:
        doc.add_paragraph("No se encontró información suficiente del centro de trabajo o la visita para completar el informe.")

    doc.add_paragraph()
    doc.add_heading("2. Metodología", level=2)


    ## Metodología breve

    doc.add_paragraph(
        "La evaluación consideró el levantamiento de antecedentes físicos y funcionales de cada recinto, incluyendo sus "
        "dimensiones, uso, número máximo de personas (NMP) y condiciones de ventilación disponibles. A partir de las "
        "mediciones de largo, ancho y alto se determinó el volumen total del recinto (m³), lo que permitió estimar el volumen "
        "disponible por persona y verificar si cumple con los valores mínimos exigidos por el D.S. N°594."
    )

    ## Metodología extendida
    '''
    doc.add_paragraph(
        "La evaluación consideró el levantamiento de antecedentes físicos y funcionales de cada recinto, incluyendo sus "
        "dimensiones, uso, número máximo de personas (NMP) y condiciones de ventilación disponibles. Posteriormente, se procedió "
        "a la obtención de los cambios de aire por hora, considerando el volumen del recinto y los sistemas de inyección o extracción "
        "de aire. Para ello, se realizaron mediciones de velocidad del aire en los puntos representativos de ventilación."
    )
    
    doc.add_paragraph(
        "Con los datos obtenidos, se calcularon los caudales volumétricos (m³/h) y se determinaron indicadores normativos como las "
        "renovaciones de aire por hora, el volumen disponible por persona y los m³ por persona por hora, verificando el cumplimiento "
        "de los valores mínimos exigidos por el D.S. N°594."
    )
    '''

    doc.add_paragraph()
    doc.add_heading("3. Resultados", level=2)
    doc.add_heading("3.1 Resumen por áreas evaluadas", level=3)

    if not df_areas.empty:
        df_areas = df_areas.sort_values(by=['nombre_area', 'codigo_area'], na_position='last') if 'codigo_area' in df_areas.columns else df_areas.sort_values(by='nombre_area')

        # --- NUEVO BLOQUE DE TABLA (reemplaza el bloque anterior de tabla_areas) ---
        # Tabla con encabezado de 2 filas y 8 columnas:
        # Área | NMP | (Dimensiones -> Largo | Ancho | Alto) | m³ sector | m³/persona | Evaluación
        tabla_areas = doc.add_table(rows=2, cols=8)
        tabla_areas.style = 'Table Grid'

        # Anchos de columnas
        set_column_width(tabla_areas, 0, Cm(3))  # Área
        set_column_width(tabla_areas, 1, Cm(3))  # NMP
        set_column_width(tabla_areas, 2, Cm(0.8))  # Largo
        set_column_width(tabla_areas, 3, Cm(0.8))  # Ancho
        set_column_width(tabla_areas, 4, Cm(0.8))  # Alto
        set_column_width(tabla_areas, 5, Cm(1.5))  # m³ sector
        set_column_width(tabla_areas, 6, Cm(1.5))  # m³/persona
        set_column_width(tabla_areas, 7, Cm(4))  # Evaluación

        # --- Alturas exactas de filas (encabezados) ---
        # Fila 0: títulos (incluye la celda fusionada "Dimensiones del recinto evaluado")
        tabla_areas.rows[0].height = Cm(2)
        tabla_areas.rows[0].height_rule = WD_ROW_HEIGHT_RULE.EXACTLY

        # Fila 1: subtítulos (Largo, Ancho, Alto)
        tabla_areas.rows[1].height = Cm(1)
        tabla_areas.rows[1].height_rule = WD_ROW_HEIGHT_RULE.EXACTLY

        # Encabezado fila 0
        hdr0 = tabla_areas.rows[0].cells
        hdr0[0].text = "Área y/o sector"
        hdr0[1].text = "Número máximo de personas\n(NMP) en el área y/o sector"
        hdr0[2].text = "Dimensiones del recinto evaluado"
        hdr0[5].text = "Metros cúbicos del sector (m³)"
        hdr0[6].text = "Metros cúbicos por persona"
        hdr0[7].text = "Evaluación estándar de ventilación D.S. N° 594\n10 metros cúbicos por persona\nCUMPLE/NO CUMPLE"

        # Encabezado fila 1 (subtítulos de Dimensiones)
        hdr1 = tabla_areas.rows[1].cells
        hdr1[2].text = "Largo (m)"
        hdr1[3].text = "Ancho (m)"
        hdr1[4].text = "Alto (m)"

        # Merges para replicar el diseño de la foto
        tabla_areas.cell(0, 0).merge(tabla_areas.cell(1, 0))  # Área
        tabla_areas.cell(0, 1).merge(tabla_areas.cell(1, 1))  # NMP
        tabla_areas.cell(0, 5).merge(tabla_areas.cell(1, 5))  # m³ sector
        tabla_areas.cell(0, 6).merge(tabla_areas.cell(1, 6))  # m³/persona
        tabla_areas.cell(0, 7).merge(tabla_areas.cell(1, 7))  # Evaluación
        tabla_areas.cell(0, 2).merge(tabla_areas.cell(0, 4))  # Dimensiones (Largo, Ancho, Alto)

        # Centrar toda la fila de títulos (fila 0)
        for c in tabla_areas.rows[0].cells:
            center_cell(c)

        # Centrar toda la fila de subtítulos (fila 1)
        for c in tabla_areas.rows[1].cells:
            center_cell(c)

        # Estilo de encabezados (tus colores/estilo corporativo)
        format_row(tabla_areas.rows[0], shading_color="4F0B7B")
        format_row(tabla_areas.rows[1], shading_color="4F0B7B")

        def _f(x):
            try:
                return float(x)
            except (TypeError, ValueError):
                return None

        # Orden como ya hacías (por nombre/código)
        df_areas = df_areas.sort_values(by=['nombre_area', 'codigo_area'], na_position='last') \
            if 'codigo_area' in df_areas.columns else df_areas.sort_values(by='nombre_area')

        for _, a in df_areas.iterrows():
            c = a.to_dict()
            row = tabla_areas.add_row().cells

            # 1) Área
            nombre = c.get('nombre_area') or c.get('codigo_area') or ''
            row[0].text = str(nombre)

            # 2) NMP
            nmp = _f(c.get('aforo_permitido')) or _f(c.get('nmp')) or 0
            row[1].text = str(int(nmp)) if nmp and nmp == int(nmp) else (format_decimal(nmp) if nmp else "")

            # 3) Dimensiones
            largo = _f(c.get('largo_m'))
            ancho = _f(c.get('ancho_m'))
            alto = _f(c.get('alto_m'))
            row[2].text = format_decimal(largo) if largo is not None else ""
            row[3].text = format_decimal(ancho) if ancho is not None else ""
            row[4].text = format_decimal(alto) if alto is not None else ""

            # 4) Volumen del sector (m³)
            volumen = _f(c.get('volumen_m3'))
            if volumen is None and all(v is not None for v in (largo, ancho, alto)):
                volumen = largo * ancho * alto
            row[5].text = format_decimal(volumen)

            # 5) m³ por persona
            m3_pp = _f(c.get('m3_porpersona'))
            if m3_pp is None and volumen is not None and nmp:
                m3_pp = volumen / nmp
            row[6].text = format_decimal(m3_pp)

            # 6) Evaluación (CUMPLE/NO CUMPLE según 10 m³/persona, salvo que venga explícito)
            eval_exp = c.get('m3_porpersona_cumple', None)
            if str(eval_exp) in ("1", "True", "true"):
                evaluacion = "CUMPLE"
            elif str(eval_exp) in ("0", "False", "false"):
                evaluacion = "NO CUMPLE"
            else:
                evaluacion = "CUMPLE" if (m3_pp is not None and m3_pp >= 10) else (
                    "NO CUMPLE" if m3_pp is not None else "")
            row[7].text = evaluacion

        # Alineación: izquierda Área, centrado el resto
        from docx.enum.text import WD_ALIGN_PARAGRAPH as _WDA
        for i, r in enumerate(tabla_areas.rows):
            # omite fila 0 (cabecera) ya con estilo
            if i == 0:
                continue
            for j, cell in enumerate(r.cells):
                for p in cell.paragraphs:
                    p.alignment = _WDA.LEFT if j == 0 else _WDA.CENTER

    doc.add_paragraph()

    if not df_puntos.empty:
        doc.add_heading("3.2 Puntos de medición", level=3)

        area_lookup = {}
        area_info_lookup = {}
        if not df_areas.empty and 'area_id' in df_areas.columns:
            area_lookup = {row['area_id']: row.get('nombre_area', row['area_id']) for _, row in df_areas.iterrows()}
            area_info_lookup = {
                row['area_id']: row.to_dict()
                for _, row in df_areas.iterrows()
            }

        df_puntos = df_puntos.sort_values(by=['area_id', 'codigo_punto']) if 'codigo_punto' in df_puntos.columns else df_puntos

        tabla_puntos = doc.add_table(rows=1, cols=8)
        tabla_puntos.style = 'Table Grid'
        headers_puntos = [
            "Área",
            "Tipo",
            "Caudal (m³/h)",
            "Número máximo de personas\n(NMP) en el área y/o sector",
            "Metros cúbicos por persona y por hora",
            "Evaluación estándar de ventilación D.S. N° 594.\nDe 20 metros cúbicos por persona por hora\nCUMPLE/ NO CUMPLE",
            "Cambios de aire por hora",
            "Evaluación estándar de ventilación D.S. N° 594.\nDe 6 hasta 60 cambios de aire por hora\nCUMPLE/ NO CUMPLE",
        ]
        for idx, texto in enumerate(headers_puntos):
            tabla_puntos.cell(0, idx).text = texto
        format_row(tabla_puntos.rows[0])

        tipo_map = {"Inyeccion": "Inyección", "Extraccion": "Extracción"}

        for _, punto in df_puntos.iterrows():
            punto_dict = punto.to_dict()
            row_cells = tabla_puntos.add_row().cells

            area_nombre = area_lookup.get(punto_dict.get('area_id'), punto_dict.get('area_id', ''))
            row_cells[0].text = str(area_nombre)
            row_cells[1].text = tipo_map.get(punto_dict.get('tipo_punto'), punto_dict.get('tipo_punto', ''))
            row_cells[2].text = format_decimal(punto_dict.get('caudal'))

            area_info = area_info_lookup.get(punto_dict.get('area_id'), {})
            nmp = _f(area_info.get('aforo_permitido')) or _f(area_info.get('nmp')) or 0
            row_cells[3].text = str(int(nmp)) if nmp and nmp == int(nmp) else (format_decimal(nmp) if nmp else "")

            m3_pp_hora = _f(area_info.get('m3_porpersona_hora'))
            row_cells[4].text = format_decimal(m3_pp_hora)

            m3_pp_hora_ref = _f(area_info.get('m3_porpersona_hora_594')) or 20.0
            m3_pp_hora_eval = area_info.get('m3_porpersona_hora_cumple', None)
            if str(m3_pp_hora_eval) in ("1", "True", "true"):
                row_cells[5].text = "CUMPLE"
            elif str(m3_pp_hora_eval) in ("0", "False", "false"):
                row_cells[5].text = "NO CUMPLE"
            else:
                row_cells[5].text = (
                    "CUMPLE" if (m3_pp_hora is not None and m3_pp_hora >= m3_pp_hora_ref) else (
                        "NO CUMPLE" if m3_pp_hora is not None else ""
                    )
                )

            recambio_hora = _f(area_info.get('recambio_hora'))
            row_cells[6].text = format_decimal(recambio_hora)

            recambio_min = _f(area_info.get('recambio_hora_594_min')) or 6.0
            recambio_max = _f(area_info.get('recambio_hora_594_max'))
            recambio_eval = area_info.get('recambio_hora_cumple', None)
            if str(recambio_eval) in ("1", "True", "true"):
                row_cells[7].text = "CUMPLE"
            elif str(recambio_eval) in ("0", "False", "false"):
                row_cells[7].text = "NO CUMPLE"
            else:
                if recambio_hora is None:
                    row_cells[7].text = ""
                else:
                    limite_max = recambio_max if recambio_max not in (None, 0) else 60.0
                    cumple_recambio = recambio_hora >= recambio_min and recambio_hora <= limite_max
                    row_cells[7].text = "CUMPLE" if cumple_recambio else "NO CUMPLE"

        set_column_width(tabla_puntos, 0, Cm(3.5))
        set_column_width(tabla_puntos, 1, Cm(2.4))
        set_column_width(tabla_puntos, 2, Cm(2.6))
        set_column_width(tabla_puntos, 3, Cm(2.6))
        set_column_width(tabla_puntos, 4, Cm(2.6))
        set_column_width(tabla_puntos, 5, Cm(3.0))
        set_column_width(tabla_puntos, 6, Cm(2.8))
        set_column_width(tabla_puntos, 7, Cm(3.2))

    else:
        doc.add_heading("3.2 Puntos de medición", level=3)
        requiere_mediciones = False
        if not df_areas.empty and 'm3_porpersona_cumple' in df_areas.columns:
            cumple_series = pd.to_numeric(
                df_areas['m3_porpersona_cumple'], errors='coerce'
            ).fillna(0)
            requiere_mediciones = (cumple_series.astype(int) == 0).any()

        if requiere_mediciones:
            doc.add_paragraph("No se registraron puntos de medición asociados a la visita.")
        else:
            doc.add_paragraph(
                "Todas las áreas evaluadas cumplen con la referencia de m³/persona, por lo que no fue necesario registrar puntos de medición."
            )

    doc.add_paragraph()

    total_areas = len(df_areas) if not df_areas.empty else 0
    areas_m3_no = []
    areas_m3_si = []
    areas_m3_h_no = []
    areas_recambio_no = []
    areas_no_cumplen = []
    mostrar_anexo_equipos = True

    if not df_areas.empty and 'nombre_area' in df_areas.columns:
        otros_indicadores_medidos = False
        if "m3_porpersona_hora" in df_areas.columns:
            otros_indicadores_medidos = (
                pd.to_numeric(df_areas["m3_porpersona_hora"], errors="coerce")
                .fillna(0)
                .gt(0)
                .any()
            )
        if not otros_indicadores_medidos and "recambio_hora" in df_areas.columns:
            otros_indicadores_medidos = (
                pd.to_numeric(df_areas["recambio_hora"], errors="coerce")
                .fillna(0)
                .gt(0)
                .any()
            )

        if 'm3_porpersona_cumple' in df_areas.columns:
            areas_m3_no = df_areas.loc[df_areas['m3_porpersona_cumple'] == 0, 'nombre_area'].tolist()
            areas_m3_si = df_areas.loc[df_areas['m3_porpersona_cumple'] == 1, 'nombre_area'].tolist()
            mostrar_anexo_equipos = len(areas_m3_no) > 0
        if 'm3_porpersona_hora_cumple' in df_areas.columns and otros_indicadores_medidos:
            areas_m3_h_no = df_areas.loc[df_areas['m3_porpersona_hora_cumple'] == 0, 'nombre_area'].tolist()
        else:
            areas_m3_h_no = []
        if 'recambio_hora_cumple' in df_areas.columns and otros_indicadores_medidos:
            areas_recambio_no = df_areas.loc[df_areas['recambio_hora_cumple'] == 0, 'nombre_area'].tolist()
        else:
            areas_recambio_no = []

        # Consolidar un listado general de áreas que presentan algún incumplimiento.
        areas_no_cumplen = sorted(
            {*(areas_m3_no or []), *(areas_m3_h_no or []), *(areas_recambio_no or [])}
        )



    # --- CONCLUSIONES ajustadas a la lógica: se omitirá m³/persona·h y recambios si no se midieron ---
    doc.add_heading("4. Conclusiones", level=2)

    # ¿Existen datos reales (no vacíos) para otros indicadores?
    cols_otras = ["m3_porpersona_hora_cumple", "recambio_hora_cumple"]
    # Consideramos que los otros indicadores fueron medidos cuando existen valores > 0.
    otros_indicadores_medidos = False
    if not df_areas.empty:
        if "m3_porpersona_hora" in df_areas.columns:
            otros_indicadores_medidos = (
                pd.to_numeric(df_areas["m3_porpersona_hora"], errors="coerce")
                .fillna(0)
                .gt(0)
                .any()
            )
        if not otros_indicadores_medidos and "recambio_hora" in df_areas.columns:
            otros_indicadores_medidos = (
                pd.to_numeric(df_areas["recambio_hora"], errors="coerce")
                .fillna(0)
                .gt(0)
                .any()
            )

    cumplen_m3 = todas_cumplen_m3(df_areas)
    areas_cumplen_texto = ', '.join(areas_m3_si) if areas_m3_si else 'evaluadas'
    print(areas_cumplen_texto)

    if cumplen_m3 and not otros_indicadores_medidos:
        # Caso que describes: solo m³/persona y todo CUMPLE → no corresponde mencionar m³/persona·h ni recambios
        doc.add_paragraph(
            f"Acorde a los resultados alcanzados, las áreas: {areas_cumplen_texto} cumplen con lo establecido en el Decreto Supremo "
            "N° 594/99 del MINSAL respecto del volumen mínimo de aire disponible por persona (10 m³ por persona). Por lo "
            "anterior, se deberán seguir las indicaciones propuestas con el fin de mantener y/o fortalecer las condiciones "
            "de ventilación evaluadas."
        )

    else:
        # Resumen general sin forzar la mención de indicadores que no se midieron
        doc.add_paragraph(
            f"Acorde a los resultados alcanzados, las áreas: {areas_cumplen_texto} cumplen con lo establecido en el Decreto Supremo "
            "N° 594/99 del MINSAL respecto del volumen mínimo de aire disponible por persona (10 m³ por persona)."
        )

        # 1) m³ por persona (siempre que exista la evaluación de m³/persona)
        if areas_m3_no:
            doc.add_paragraph(
                f"El indicador de volumen por persona (10 m³/persona) no cumple en: {', '.join(areas_m3_no)}."
            )
        elif not otros_indicadores_medidos:
            # Si no hay otros indicadores medidos y no hay incumplimientos de m³/persona, cerramos aquí sin mencionar m³/h ni recambios
            doc.add_paragraph(
                "El indicador de volumen por persona cumple el criterio normativo establecido para las áreas revisadas."
            )

        # 2) Solo mencionar m³/persona·h y recambios si efectivamente se midieron (algún dato no nulo)
        if otros_indicadores_medidos:
            if areas_m3_h_no:
                doc.add_paragraph(
                    f"El indicador de m³ por persona por hora requiere ajuste en: {', '.join(areas_m3_h_no)}."
                )
            if areas_recambio_no:
                doc.add_paragraph(
                    f"El recambio de aire por hora es inferior al criterio normativo en: {', '.join(areas_recambio_no)}."
                )
            # Si todos los indicadores medidos cumplen (y hubo otros además de m³/persona), puedes cerrar con un consolidado:
            if (not areas_m3_no) and (not areas_m3_h_no) and (not areas_recambio_no):
                doc.add_paragraph(
                    "Los indicadores evaluados cumplen los criterios normativos establecidos para las áreas revisadas."
                )

    # --- NUEVO: prescripciones administrativas en tabla (reemplaza '5. Recomendaciones') ---
    doc.add_paragraph()
    doc.add_heading("5. Prescripciones", level=2)
    doc.add_heading("5.1 Medidas de carácter Administrativas", level=3)

    # Tabla: Medidas de control | Área y/o sector | Plazo de cumplimiento
    tabla_presc = doc.add_table(rows=1, cols=3)
    tabla_presc.style = 'Table Grid'
    hdr = tabla_presc.rows[0].cells
    hdr[0].text = "Medidas de control"
    hdr[1].text = "Área y/o sector"
    hdr[2].text = "Plazo de cumplimiento"
    format_row(tabla_presc.rows[0], shading_color="4F0B7B")

    # Medida 1
    r1 = tabla_presc.add_row().cells
    r1[0].text = (
        "Informar a cada persona trabajadora acerca de los riesgos que entrañan sus labores, de las medidas preventivas, "
        "de los métodos y/o procedimientos de trabajo correctos, acorde a lo identificado por la empresa. Además, la "
        "entidad empleadora deberá informar de manera oportuna y adecuada el resultado del presente informe técnico. "
        "En el marco del artículo 15° del Párrafo IV del D.S 44 “Aprueba nuevo reglamento sobre gestión preventiva de los "
        "riesgos laborales para un entorno de trabajo seguro y saludable”."
    )
    r1[1].text = "Todas las áreas evaluadas"
    r1[2].text = "30 días desde la recepción del presente informe técnico."

    # Medida 2
    r2 = tabla_presc.add_row().cells
    r2[0].text = (
        "Realizar capacitaciones (teóricas/prácticas) periódicas en prevención de riesgos laborales, con la finalidad de "
        "garantizar el aprendizaje efectivo y eficaz, dejando registro de dichas capacitaciones y evaluaciones. En el marco "
        "del artículo 16° del Párrafo IV del D.S 44 “Aprueba nuevo reglamento sobre gestión preventiva de los riesgos laborales "
        "para un entorno de trabajo seguro y saludable”."
    )
    r2[1].text = "Todas las áreas evaluadas"
    r2[2].text = "De acuerdo a Programa de Capacitación vigente, se debe ejecutar lo prescrito de manera permanente."

    # Anchos de columnas (aproximados a tu layout)
    set_column_width(tabla_presc, 0, Cm(12))
    set_column_width(tabla_presc, 1, Cm(4))
    set_column_width(tabla_presc, 2, Cm(4))

    # -------------------------------
    # 5) VIGENCIA DEL INFORME
    # -------------------------------
    # Encabezado principal del contenido: Vigencia del informe
    paragraph = doc.add_heading("6. Vigencia del informe", level=2)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    if not df_visitas.empty:
        doc.add_paragraph(
            "En términos generales, el presente informe tiene validez de 3 años, a excepción que existan cambios en la situación, del tipo ingenieril o administrativo, que presupongan modificación a las condiciones encontradas al momento de la medición, lo cual implicará realizar una nueva evaluación en un plazo menor al señalado.")
        doc.add_paragraph()
        if len(areas_no_cumplen) > 0:
            doc.add_paragraph(
                "Cuando se concreten los cambios indicados, la empresa deberá informar al IST el detalle de los mismos, de forma tal de programar las gestiones a realizar, las que considerarán previamente un informe de Verificación y Control y posteriormente, de acuerdo a sus resultados, la nueva evaluación de higiene ocupacional correspondiente.")
            doc.add_paragraph()
        doc.add_paragraph(
            "Estos resultados de evaluación representan las condiciones existentes del ambiente y lugar de trabajo al momento de realizar las mediciones.")
    doc.add_paragraph()
    doc.add_paragraph()
    doc.add_paragraph()
    doc.add_paragraph()

    consultor_nombre = ""
    consultor_name_complete = ""
    consultor_cargo = ""
    consultor_zonal = ""

    if not df_visitas.empty:
        row_visita = df_visitas.iloc[0]
        consultor_nombre = row_visita.get("consultor_nombre", "")
        consultor_name_complete = (
            row_visita.get("consultor_name_complete") or consultor_nombre
        )
        consultor_cargo = row_visita.get("consultor_cargo", "")
        consultor_zonal = row_visita.get("consultor_zonal", "")

    consultor_ist_limpio = (
        consultor_nombre
        .replace(" ", "-")
        .replace("ñ", "n")
        .replace("Ñ", "N")
        .replace("á", "a")
        .replace("é", "e")
        .replace("í", "i")
        .replace("ó", "o")
        .replace("ú", "u")
        .replace("Á", "A")
    )

    firma_path = os.path.join("imagenes-firma", consultor_ist_limpio + ".png")
    print (firma_path)


    try:
        paragraph = doc.add_paragraph()
        run = paragraph.add_run()
        run.add_picture(firma_path, width=Cm(4))
        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    except FileNotFoundError:
        print(f"La imagen no existe en la ruta especificada: {firma_path}")
    except Exception as e:
        print(f"Ocurrió un error al agregar la imagen: {e}")

    # Agregar párrafo para el consultor, centrado y en negrita
    p_consultor = doc.add_paragraph()
    p_consultor.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_consultor = p_consultor.add_run(consultor_name_complete)
    run_consultor.bold = True

    # Agregar párrafo para la profesión, centrado
    p_profesion = doc.add_paragraph()
    p_profesion.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_profesion =p_profesion.add_run (consultor_cargo)

    # Agregar párrafo para el zonal, centrado
    p_zonal = doc.add_paragraph()
    p_zonal.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_zonal = p_zonal.add_run(consultor_zonal)

    anexos_info_vent: List[dict] = []
    if not df_areas.empty and "nombre_area" in df_areas.columns:
        df_areas_sorted = df_areas.copy()
        if "nombre_area" in df_areas_sorted.columns:
            df_areas_sorted = df_areas_sorted.sort_values(by="nombre_area")

        for _, fila in df_areas_sorted.iterrows():
            nombre = str(fila.get("nombre_area", "") or "")
            observacion = str(fila.get("observaciones", "") or "").strip()
            if not observacion:
                observacion = "Sin observaciones registradas."
            area_id = fila.get("area_id")

            visita_id = None
            for columna in ("visita_id", "id_visita"):
                valor = fila.get(columna)
                if pd.notna(valor):
                    visita_id = valor
                    break

            imagenes = _listar_imagenes_area(area_id, nombre, visita_id)
            anexos_info_vent.append(
                {
                    "nombre": nombre,
                    "observaciones": observacion,
                    "imagenes": imagenes,
                }
            )

    doc.add_page_break()
    _agregar_seccion_anexos(doc, anexos_info_vent)

    if mostrar_anexo_equipos:
        doc.add_paragraph()
        doc.add_heading("Anexo 2. Instrumentos de medición utilizados", level=2)
        _agregar_anexo_equipos(doc, df_visitas, df_equipos)

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer
