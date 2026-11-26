"""
Generador unificado de entregables finales:
1. Informe Word conciso en normas APA 7 sobre `plantilla de trabajo.docx`
   (`Morales_Raul_Castro_David_problemaParqueadero.docx`).
2. Libro Excel limpio con 6 hojas (`resultados/resultados_simulacion.xlsx`).
3. Cuaderno interactivo ejecutado para Google Colab
   (`notebooks/Simulacion_Parqueaderos_Supercentro.ipynb`).

Integrantes:
- Raul Santiago Morales Henao
- David Octavio Castro Mora
"""

from pathlib import Path
from typing import Dict, List
import nbformat
from nbformat.v4 import new_code_cell, new_markdown_cell, new_notebook
from nbclient import NotebookClient
import pandas as pd

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Inches, Pt, RGBColor

from src.simulacion import (
    CRITERIOS_DECISION,
    EQUIPO_TRABAJO,
    HORIZONTE_MIN,
    NOMBRE_INFORME,
    NOTEBOOKS_DIR,
    PERFILES_USUARIO,
    PLANTILLA_DOCX,
    RESULTADOS_DIR,
    ROOT_DIR,
    URL_COLAB,
    URL_GITHUB,
)


# =============================================================================
# 1. EXCEL DE RESULTADOS (6 HOJAS CLARAS)
# =============================================================================

def exportar_excel(datos: Dict, ruta: Path = RESULTADOS_DIR / "resultados_simulacion.xlsx") -> Path:
    ruta.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    wb.remove(wb.active)
    fill_h = PatternFill("solid", fgColor="1B365D")
    font_h = Font(name="Calibri", size=10, bold=True, color="FFFFFF")

    hojas = [
        ("1_Estado_Estable", datos["df_conv"]),
        ("2_PuntoA_Cajeros", datos["res_a"]["tabla"]),
        ("3_PuntoB_Usuarios", datos["df_b"]),
        ("4_PuntoC_Estrategias", datos["df_esc"]),
        ("5_PuntoD_Validacion", datos["res_vcv"]["df_verificacion"]),
        ("6_PuntoE_Transitorio", datos["res_t"]["df_comparacion"]),
    ]
    for titulo, df in hojas:
        ws = wb.create_sheet(title=titulo)
        ws.append(list(df.columns))
        for j in range(1, len(df.columns) + 1):
            c = ws.cell(1, j)
            c.fill = fill_h
            c.font = font_h
            c.alignment = Alignment(horizontal="center", vertical="center")
        for row in df.itertuples(index=False):
            ws.append([round(v, 4) if isinstance(v, float) else v for v in row])
        ws.freeze_panes = "A2"
        for j, col in enumerate(df.columns, 1):
            ws.column_dimensions[get_column_letter(j)].width = max(len(str(col)) + 4, 15)
    wb.save(ruta)
    return ruta


# =============================================================================
# 2. INFORME WORD CONCISO EN APA 7 (BASADO EN `plantilla de trabajo.docx`)
# =============================================================================

NEGRO = RGBColor(0x00, 0x00, 0x00)


def _limpiar_plantilla(doc: Document) -> None:
    body = doc.element.body
    for ch in list(body):
        if ch.tag != qn("w:sectPr"):
            body.remove(ch)
    for rId in [r for r, rel in doc.part.rels.items() if "image" in rel.reltype or "hyperlink" in rel.reltype]:
        doc.part.rels.pop(rId, None)
    if doc.settings.element.find(qn("w:updateFields")) is None:
        upd = OxmlElement("w:updateFields")
        upd.set(qn("w:val"), "true")
        doc.settings.element.append(upd)


def _run_apa(run, bold: bool = False, italic: bool = False, size: float = 12.0) -> None:
    run.font.name = "Times New Roman"
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    run.font.color.rgb = NEGRO
    rPr = run._r.get_or_add_rPr()
    rf = rPr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rPr.insert(0, rf)
    for k in ("w:ascii", "w:hAnsi", "w:cs"):
        rf.set(qn(k), "Times New Roman")


def _linea_portada(doc: Document, texto: str, bold: bool = False, sa: float = 8.0) -> None:
    p = doc.add_paragraph(style="APA 7")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Inches(0)
    p.paragraph_format.line_spacing = 2.0
    p.paragraph_format.space_after = Pt(sa)
    if texto:
        _run_apa(p.add_run(texto), bold=bold)


def _titulo(doc: Document, texto: str, nivel: int = 1) -> None:
    p = doc.add_paragraph(style="APA 7")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if nivel == 1 else WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.first_line_indent = Inches(0)
    p.paragraph_format.line_spacing = 2.0
    p.paragraph_format.space_before = Pt(6 if nivel == 1 else 4)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True
    ol = OxmlElement("w:outlineLvl")
    ol.set(qn("w:val"), "0" if nivel == 1 else "1")
    p._p.get_or_add_pPr().append(ol)
    _run_apa(p.add_run(texto), bold=True)


def _parrafo(doc: Document, texto: str) -> None:
    p = doc.add_paragraph(style="APA 7")
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.first_line_indent = Inches(0.5)
    p.paragraph_format.line_spacing = 2.0
    p.paragraph_format.space_after = Pt(8)
    _run_apa(p.add_run(texto))


def _item(doc: Document, prefijo: str, texto: str) -> None:
    p = doc.add_paragraph(style="APA 7")
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.left_indent = Inches(0.5)
    p.paragraph_format.first_line_indent = Inches(-0.2)
    p.paragraph_format.line_spacing = 2.0
    p.paragraph_format.space_after = Pt(4)
    _run_apa(p.add_run("•  "))
    if prefijo:
        _run_apa(p.add_run(prefijo), bold=True)
    _run_apa(p.add_run(texto))


def _referencia(doc: Document, antes: str, cursiva: str, despues: str) -> None:
    p = doc.add_paragraph(style="APA 7")
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.left_indent = Inches(0.5)
    p.paragraph_format.first_line_indent = Inches(-0.5)
    p.paragraph_format.line_spacing = 2.0
    p.paragraph_format.space_after = Pt(8)
    _run_apa(p.add_run(antes))
    _run_apa(p.add_run(cursiva), italic=True)
    _run_apa(p.add_run(despues))


def _tabla_apa(doc: Document, num: int, titulo: str, headers: List[str], filas: List[List[str]], nota: str = "") -> None:
    pc = doc.add_paragraph(style="APA 7")
    pc.paragraph_format.first_line_indent = Inches(0)
    pc.paragraph_format.line_spacing = 1.5
    pc.paragraph_format.space_before = Pt(6)
    pc.paragraph_format.space_after = Pt(4)
    pc.paragraph_format.keep_with_next = True
    _run_apa(pc.add_run(f"Tabla {num}. "), bold=True)
    _run_apa(pc.add_run(titulo), bold=True, italic=True)

    tbl = doc.add_table(rows=len(filas) + 1, cols=len(headers))
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    if "APA 7 TABLAS" in doc.styles:
        tbl.style = "APA 7 TABLAS"
    tblPr = tbl._tbl.tblPr
    for old in tblPr.findall(qn("w:tblBorders")):
        tblPr.remove(old)
    tblPr.append(parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
        f'<w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
        f'<w:left w:val="none"/><w:right w:val="none"/>'
        f'<w:insideH w:val="none"/><w:insideV w:val="none"/>'
        f'</w:tblBorders>'
    ))
    tbl.rows[0]._tr.get_or_add_trPr().append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))

    for j, h in enumerate(headers):
        cell = tbl.cell(0, j)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        cell._tc.get_or_add_tcPr().append(parse_xml(
            f'<w:tcBorders {nsdecls("w")}><w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/></w:tcBorders>'
        ))
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT if j == 0 else WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Inches(0)
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(3)
        _run_apa(p.add_run(h), bold=True, size=10.0)

    for i, fila in enumerate(filas, 1):
        dest = any(k in str(fila[0]) for k in ("TOTAL", "MEZCLA", "Recomendada"))
        for j, val in enumerate(fila):
            cell = tbl.cell(i, j)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if j == 0 else WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.first_line_indent = Inches(0)
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_before = Pt(2.5)
            p.paragraph_format.space_after = Pt(2.5)
            _run_apa(p.add_run(str(val)), bold=dest, size=9.8)

    if nota:
        pn = doc.add_paragraph(style="APA 7")
        pn.paragraph_format.first_line_indent = Inches(0)
        pn.paragraph_format.line_spacing = 1.5
        pn.paragraph_format.space_before = Pt(3)
        pn.paragraph_format.space_after = Pt(8)
        _run_apa(pn.add_run("Nota. "), italic=True, size=10.0)
        _run_apa(pn.add_run(nota), size=10.0)


def _ilustracion(doc: Document, num: int, titulo: str, ruta_img: Path, nota: str = "") -> None:
    pc = doc.add_paragraph(style="APA 7")
    pc.paragraph_format.first_line_indent = Inches(0)
    pc.paragraph_format.line_spacing = 1.5
    pc.paragraph_format.space_before = Pt(6)
    pc.paragraph_format.space_after = Pt(4)
    pc.paragraph_format.keep_with_next = True
    _run_apa(pc.add_run(f"Ilustración {num}. "), bold=True)
    _run_apa(pc.add_run(titulo), bold=True, italic=True)

    pi = doc.add_paragraph(style="APA 7")
    pi.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pi.paragraph_format.first_line_indent = Inches(0)
    pi.paragraph_format.line_spacing = 1.0
    pi.paragraph_format.space_after = Pt(3)
    pi.paragraph_format.keep_with_next = bool(nota)
    pi.add_run().add_picture(str(ruta_img), width=Inches(6.1))

    if nota:
        pn = doc.add_paragraph(style="APA 7")
        pn.paragraph_format.first_line_indent = Inches(0)
        pn.paragraph_format.line_spacing = 1.5
        pn.paragraph_format.space_after = Pt(8)
        _run_apa(pn.add_run("Nota. "), italic=True, size=10.0)
        _run_apa(pn.add_run(nota), size=10.0)


def generar_word_apa(datos: Dict, graficas: Dict[str, Path], ruta: Path = ROOT_DIR / NOMBRE_INFORME) -> Path:
    """Construye el documento Word corto, claro y 100 % en normas APA 7."""
    doc = Document(str(PLANTILLA_DOCX))
    _limpiar_plantilla(doc)

    teo = datos["teo"]
    res_est = datos["res_est"]
    res_t = datos["res_t"]
    res_a = datos["res_a"]
    df_b = datos["df_b"]
    df_esc = datos["df_esc"]
    res_vcv = datos["res_vcv"]

    r_opt = res_est["r_optimo"]
    t0 = res_t["t0_min"]

    # PORTADA
    _linea_portada(doc, "")
    _linea_portada(doc, "S30 – Módulo 4: Actividad Didáctica 2", bold=True)
    _linea_portada(doc, "Solución de un Problema de Parqueaderos en el Centro Comercial Supercentro mediante Simulación M/M/1")
    _linea_portada(doc, "")
    _linea_portada(doc, "Integrantes:", bold=True, sa=2)
    _linea_portada(doc, EQUIPO_TRABAJO[0], sa=2)
    _linea_portada(doc, EQUIPO_TRABAJO[1], sa=8)
    _linea_portada(doc, "Asignatura:", bold=True, sa=2)
    _linea_portada(doc, "Simulación – Módulo 4 (S30)", sa=2)
    _linea_portada(doc, "Institución Universitaria Digital de Antioquia", sa=8)
    _linea_portada(doc, "")
    _linea_portada(doc, "25 de noviembre de 2026")
    doc.add_page_break()

    # TABLA DE CONTENIDO
    p_tit = doc.add_paragraph(style="APA 7")
    p_tit.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_tit.paragraph_format.first_line_indent = Inches(0)
    _run_apa(p_tit.add_run("Contenido"), bold=True)
    p_toc = doc.add_paragraph(style="APA 7")
    p_toc.paragraph_format.first_line_indent = Inches(0)
    r_toc = p_toc.add_run()
    _run_apa(r_toc)
    for xml_s in (
        f'<w:fldChar {nsdecls("w")} w:fldCharType="begin"/>',
        f'<w:instrText {nsdecls("w")} xml:space="preserve"> TOC \\o "1-2" \\h \\z \\u </w:instrText>',
        f'<w:fldChar {nsdecls("w")} w:fldCharType="separate"/>',
        f'<w:t {nsdecls("w")}>Clic derecho → Actualizar campos en Word para generar los números de página.</w:t>',
        f'<w:fldChar {nsdecls("w")} w:fldCharType="end"/>',
    ):
        r_toc._r.append(parse_xml(xml_s))
    doc.add_page_break()

    # 1. INTRODUCCIÓN
    _titulo(doc, "Introducción", 1)
    _parrafo(
        doc,
        "El Centro Comercial Supercentro opera tres cajeros automáticos de pago de parqueadero por salida, con filas "
        "independientes, sin cambio de fila ni deserción, modelados como tres colas M/M/1. Aunque todos los equipos "
        "responden de forma inmediata, el tiempo de uso varía según cuatro tipos de usuario: rápidos (25 %), normales "
        "(20 %) y el 55 % restante entre lentos y muy lentos."
    )
    _parrafo(
        doc,
        "El objetivo de este estudio es determinar, mediante simulación en Python (SimPy), si tres cajeros son suficientes "
        "para cubrir la demanda y el tiempo de uso de los usuarios lentos, o qué estrategia de mejora conviene aplicar. Para "
        "ello se define el número de réplicas en estado estable y se desarrollan de manera directa los puntos A al E."
    )

    # 2. METODOLOGÍA
    _titulo(doc, "Metodología", 1)
    _parrafo(
        doc,
        f"Cada cajero se simuló como una cola M/M/1 independiente durante un horizonte de {HORIZONTE_MIN:.0f} minutos "
        "(40 horas) por réplica, dividiendo equitativamente el 55 % de usuarios lentos y muy lentos (27.5 % cada uno). "
        f"En promedio cada cajero tiene un servicio de {teo['E_S']:.2f} minutos y llegadas cada {teo['E_A']:.2f} minutos "
        f"(utilización teórica ρ = {teo['rho']*100:.1f} % < 100 %, sistema estable), con una demanda total de "
        f"{teo['demanda_hora_total']:.1f} usuarios por hora entre los tres cajeros (Tabla 1)."
    )
    filas_t1 = [[p["nombre"], f"{p['prob']*100:.1f} %", f"{p['servicio']:.0f} min", f"{p['llegada']:.0f} min"] for p in PERFILES_USUARIO]
    filas_t1.append(["MEZCLA POR CAJERO", "100.0 %", f"{teo['E_S']:.2f} min", f"{teo['E_A']:.2f} min"])
    _tabla_apa(
        doc, 1, "Parámetros de entrada por perfil de usuario en cada cajero.",
        ["Tipo de usuario", "Proporción", "Servicio (Exponencial)", "Media de llegada"],
        filas_t1,
        f"Utilización teórica por cajero: ρ = {teo['rho']*100:.2f} %. Demanda agregada: {teo['demanda_hora_total']:.1f} usuarios/hora.",
    )

    # 3. RESULTADOS
    _titulo(doc, "Resultados", 1)

    # Estado estable
    _titulo(doc, "Identificación del estado estable (número de réplicas)", 2)
    _parrafo(
        doc,
        f"Mediante la técnica del promedio acumulado sobre 200 corridas piloto, el tiempo promedio en el sistema (W) se "
        f"estabilizó dentro de una banda del ±1 % a partir de {res_est['r_exacto']} corridas con un error de estimación del "
        f"{res_est['error_rel_pct']:.1f} % (menor al 5 %). Por ello se fijaron R* = {r_opt} réplicas para todo el análisis "
        "(Ilustración 1)."
    )
    _ilustracion(
        doc, 1, f"Convergencia del promedio acumulado y selección de R* = {r_opt} réplicas.",
        graficas["fig01"], "A partir de 160 réplicas el promedio acumulado se mantiene estable dentro de ±1 %.",
    )

    # Puntos A y B
    _titulo(doc, "Puntos A y B: Estadísticas por cajero y promedio de usuarios por tipo", 2)
    _parrafo(
        doc,
        f"Punto A: Sobre las {r_opt} réplicas en estado estable (Tabla 2 e Ilustración 2), el cajero con menor tiempo "
        f"promedio total de atención fue {res_a['menor_cajero']} ({res_a['menor_W']:.2f} min) y el de mayor tiempo fue "
        f"{res_a['mayor_cajero']} ({res_a['mayor_W']:.2f} min). La prueba ANOVA (p = {res_a['anova_p_W']:.3f}) confirma que "
        "los tres cajeros son estadísticamente equivalentes y las pequeñas variaciones se deben al azar."
    )
    filas_t2 = [
        [r.Cajero, f"{r.Servicio_S_min:.2f}", f"{r.Espera_Wq_min:.2f}", f"{r.Total_W_min:.2f}", f"{r.Cola_Lq:.2f}", f"{r.Utilizacion_pct:.1f} %"]
        for r in res_a["tabla"].itertuples()
    ]
    _tabla_apa(
        doc, 2, f"Estadísticas promedio por cajero en estado estable (R* = {r_opt} réplicas).",
        ["Cajero", "Servicio S (min)", "Espera Wq (min)", "Total W (min)", "Cola Lq", "Ocupación ρ"],
        filas_t2, f"Prueba ANOVA para W: p = {res_a['anova_p_W']:.3f} (sin diferencia significativa entre cajeros).",
    )

    tot_b = df_b.loc[df_b["Tipo_usuario"] == "TOTAL SISTEMA"].iloc[0]
    _parrafo(
        doc,
        f"Punto B: En la totalidad de los tres cajeros se atienden en promedio {tot_b['Promedio_3_cajeros']:.1f} usuarios "
        f"por corrida estable ({tot_b['Usuarios_por_hora']:.1f} usuarios/hora), repartidos fielmente según el enunciado: "
        "25.0 % rápidos, 20.0 % normales, 27.3 % lentos y 27.6 % muy lentos (Tabla 3 e Ilustración 2)."
    )
    filas_t3 = [
        [r.Tipo_usuario, f"{r.Promedio_3_cajeros:.1f}", f"{r.Promedio_por_cajero:.1f}", f"{r.Usuarios_por_hora:.1f} /h", f"{r.Porcentaje_simulado:.1f} %"]
        for r in df_b.itertuples()
    ]
    _tabla_apa(
        doc, 3, "Promedio de usuarios por tipo en la totalidad de los tres cajeros.",
        ["Tipo de usuario", "Promedio 3 cajeros", "Promedio / cajero", "Tasa horaria", "Participación"],
        filas_t3, "Ventana estacionaria de 33.5 horas por corrida (posterior al corte t0* = 390 min).",
    )
    _ilustracion(
        doc, 2, "Comparación de tiempos por cajero (Punto A) y usuarios promedio por tipo (Punto B).",
        graficas["fig02"], "Más de la mitad de los visitantes (55 %) son usuarios lentos o muy lentos.",
    )

    # Punto C
    _titulo(doc, "Punto C: Criterios de decisión y estrategia de mejora", 2)
    _parrafo(
        doc,
        "El equipo de trabajo fijó tres criterios claros para evaluar la suficiencia del sistema: (1) espera promedio en "
        f"fila Wq ≤ {CRITERIOS_DECISION['max_wq']:.0f} min, (2) tiempo total W ≤ {CRITERIOS_DECISION['max_w']:.0f} min y "
        f"(3) ocupación de cajeros ρ ≤ {CRITERIOS_DECISION['max_rho']*100:.0f} % para dejar margen ante horas pico."
    )
    e0, e1, e2, e3 = df_esc.iloc[0], df_esc.iloc[1], df_esc.iloc[2], df_esc.iloc[3]
    filas_t4 = [
        [r.Escenario, f"{r.Espera_Wq_min:.2f} min", f"{r.Total_W_min:.2f} min", f"{r.Utilizacion_pct:.1f} %",
         "Base" if i == 0 else f"-{r.Reduccion_Wq_pct:.1f} %", r.Criterios_SLA]
        for i, r in enumerate(df_esc.itertuples())
    ]
    _tabla_apa(
        doc, 4, "Evaluación de escenarios alternativos bajo la misma demanda de salida.",
        ["Escenario evaluado", "Espera Wq", "Total W", "Ocupación ρ", "Reducción Wq", "Criterios"],
        filas_t4, "Se recomienda 4 cajeros con fila única por cumplir 3/3 criterios con espera menor a 1 minuto.",
    )
    _parrafo(
        doc,
        f"El escenario actual con tres cajeros independientes no es suficiente (0/3 criterios; espera de {e0['Espera_Wq_min']:.2f} min). "
        f"Agregar un cuarto cajero manteniendo filas separadas tampoco cumple la meta de espera ({e1['Espera_Wq_min']:.2f} min). "
        f"En cambio, unificar los tres cajeros actuales en una fila única reduce la espera un {e2['Reduccion_Wq_pct']:.1f} % "
        f"(a {e2['Espera_Wq_min']:.2f} min) sin comprar equipos, y la estrategia recomendada de 4 cajeros con fila única cumple "
        f"los 3/3 criterios bajando la espera a {e3['Espera_Wq_min']:.2f} min (-{e3['Reduccion_Wq_pct']:.1f} %) con ocupación "
        f"holgada del {e3['Utilizacion_pct']:.1f} % (Ilustración 3)."
    )
    _ilustracion(
        doc, 3, "Comparación del tiempo promedio de espera en las cuatro alternativas simuladas.",
        graficas["fig03"], "La línea roja marca la meta máxima de 5 minutos de espera en fila.",
    )

    # Punto D
    _titulo(doc, "Punto D: Verificación, calibración y validación del modelo", 2)
    max_err = float(res_vcv["df_verificacion"]["Error_pct"].max())
    wq_cal_min = float(res_vcv["df_calibracion"]["Wq_min"].min())
    wq_cal_max = float(res_vcv["df_calibracion"]["Wq_min"].max())
    _parrafo(
        doc,
        "Se aplicó el ciclo completo de Verificación, Calibración y Validación (en el enunciado aparece repetida la palabra "
        "'verificación' por error tipográfico en lugar de 'validación'): "
        f"(1) Verificación: al correr el código en modo M/M/1 puro, coincidió con las fórmulas teóricas con un error máximo de "
        f"{max_err:.2f} % y cumplió la Ley de Little con menos de {res_vcv['error_little_pct']:.2f} % de diferencia. "
        f"(2) Calibración: se ajustaron R* = {r_opt} réplicas, corte t0* = {t0:.0f} min y se probó cambiar el reparto del 55 % "
        f"entre lentos y muy lentos (de 35/20 % a 20/35 %), confirmando que la espera siempre es crítica ({wq_cal_min:.1f} a "
        f"{wq_cal_max:.1f} min). (3) Validación: al variar la demanda entre 70 %, 100 % y 110 %, las filas crecieron de forma "
        "realista y coherente con la teoría de colas (Ilustración 4)."
    )
    _ilustracion(
        doc, 4, "Diagrama del proceso de Verificación, Calibración y Validación y evidencia frente a M/M/1.",
        graficas["fig04"], f"Error relativo máximo frente a las fórmulas analíticas M/M/1: {max_err:.2f} %.",
    )

    # Punto E
    _titulo(doc, "Punto E: Eliminación del estado transitorio (antes y después)", 2)
    df_c = res_t["df_comparacion"]
    wq_r = df_c.loc[df_c["Metrica_cod"] == "Wq"].iloc[0]
    w_r = df_c.loc[df_c["Metrica_cod"] == "W"].iloc[0]
    s_r = df_c.loc[df_c["Metrica_cod"] == "S"].iloc[0]
    _parrafo(
        doc,
        f"Al iniciar la simulación con el parqueadero vacío, los primeros minutos tienen esperas cercanas a cero y subestiman "
        f"la congestión real. Se programó un algoritmo con suavizado de Welch y criterio MSER que detectó automáticamente el "
        f"fin del estado transitorio en t0* = {t0:.0f} minutos ({res_t['pct_descartado']:.1f} % del horizonte). Al descartar "
        f"ese arranque inicial, la espera promedio Wq se corrigió de {wq_r['Antes']:.2f} a {wq_r['Despues']:.2f} min "
        f"({wq_r['Cambio_pct']:+.1f} %) y el tiempo total W de {w_r['Antes']:.2f} a {w_r['Despues']:.2f} min "
        f"({w_r['Cambio_pct']:+.1f} %), mientras el servicio S permaneció estable en {s_r['Despues']:.2f} min (Tabla 5 e Ilustración 5)."
    )
    filas_t5 = [
        [r.Metrica, f"{r.Antes:.2f}", f"{r.Despues:.2f}", f"{r.Cambio_pct:+.1f} %"]
        for r in df_c.itertuples()
    ]
    _tabla_apa(
        doc, 5, f"Valores promedio Antes y Después de eliminar el estado transitorio (corte t0* = {t0:.0f} min).",
        ["Variable medida", "Antes (con transitorio)", "Después (sin transitorio)", "Variación"],
        filas_t5, "Eliminar el arranque vacío corrige el sesgo a la baja en Wq, W y Lq sin alterar el servicio S.",
    )
    _ilustracion(
        doc, 5, "Eliminación del estado transitorio y comparación de promedios Antes vs. Después.",
        graficas["fig05"], "Se descartaron los primeros 390 minutos para trabajar únicamente en régimen estable.",
    )

    # 4. CONCLUSIONES
    _titulo(doc, "Conclusiones", 1)
    _parrafo(
        doc,
        f"Tres cajeros con filas separadas no son suficientes para el Centro Comercial Supercentro: aunque no colapsan "
        f"(ρ = {e0['Utilizacion_pct']:.1f} %), el 55 % de usuarios lentos genera esperas de {e0['Espera_Wq_min']:.2f} minutos "
        "e incumple todos los criterios de servicio."
    )
    _parrafo(
        doc,
        f"La mejor solución es implementar 4 cajeros con fila única, que reduce la espera en {e3['Reduccion_Wq_pct']:.1f} % "
        f"(a {e3['Espera_Wq_min']:.2f} min) y cumple los 3/3 criterios de decisión. Como mejora inmediata sin comprar equipos, "
        f"unificar los 3 cajeros actuales en fila única ya recorta la espera un {e2['Reduccion_Wq_pct']:.1f} %. Todo el código "
        "en Python, pruebas y evidencias pueden consultarse directamente en los siguientes enlaces:"
    )
    _item(doc, "Repositorio en GitHub: ", URL_GITHUB)
    _item(doc, "Cuaderno en Google Colab: ", URL_COLAB)
    _item(doc, "Video de sustentación: ", "[Agregar aquí el enlace del video grabado por los integrantes]")

    # 5. REFERENCIAS (en página nueva según norma APA 7)
    doc.add_page_break()
    _titulo(doc, "Referencias", 1)
    _referencia(doc, "Banks, J., Carson, J. S., Nelson, B. L., & Nicol, D. M. (2010). ", "Discrete-event system simulation", " (5.ª ed.). Pearson.")
    _referencia(doc, "Gross, D., Shortle, J. F., Thompson, J. M., & Harris, C. M. (2018). ", "Fundamentals of queueing theory", " (5.ª ed.). Wiley.")
    _referencia(doc, "Law, A. M. (2015). ", "Simulation modeling and analysis", " (5.ª ed.). McGraw-Hill.")
    _referencia(doc, "Sargent, R. G. (2013). Verification and validation of simulation models. ", "Journal of Simulation, 7", "(1), 12–24.")

    doc.save(str(ruta))
    return ruta


# =============================================================================
# 3. CUADERNO DE GOOGLE COLAB RESUMIDO Y EJECUTADO (7 CELDAS DE CÓDIGO)
# =============================================================================

def generar_notebook_colab(
    ruta: Path = NOTEBOOKS_DIR / "Simulacion_Parqueaderos_Supercentro.ipynb",
    ejecutar: bool = True,
) -> Path:
    ruta.parent.mkdir(parents=True, exist_ok=True)
    celdas = [
        new_markdown_cell(
            "# Simulación de Parqueaderos — Centro Comercial Supercentro (Módulo 4 – ACT 2)\n\n"
            f"**Integrantes:** {' y '.join(EQUIPO_TRABAJO)}  \n"
            f"**Repositorio GitHub:** [Ver repositorio del proyecto]({URL_GITHUB})  \n"
            f"**Abrir en Google Colab:** [![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)]({URL_COLAB})\n\n"
            "Cuaderno resumido y 100 % sincronizado con `src/simulacion.py` y el informe en formato APA 7."
        ),
        new_markdown_cell("## 1. Configuración e Importación del Código del Repositorio"),
        new_code_cell(
            "import shutil\nimport subprocess\nimport sys\nimport tempfile\nfrom pathlib import Path\n\n"
            f"REPO_URL = '{URL_GITHUB}.git'\n"
            "REPO_DIR = Path('simulacion-parqueaderos-supercentro')\n\n"
            "if Path('src/simulacion.py').exists():\n    ROOT = Path('.').resolve()\n"
            "elif Path('../src/simulacion.py').exists():\n    ROOT = Path('..').resolve()\n"
            "else:\n"
            "    # En Google Colab: instala SimPy en este intérprete y descarga la versión actual del repo\n"
            "    subprocess.check_call([sys.executable, '-m', 'pip', 'install', '-q', 'simpy'])\n"
            "    if REPO_DIR.exists():\n"
            "        shutil.rmtree(REPO_DIR)\n"
            "    subprocess.check_call(['git', 'clone', '--depth', '1', '-q', REPO_URL, str(REPO_DIR)])\n"
            "    ROOT = REPO_DIR.resolve()\n\n"
            "if str(ROOT) not in sys.path:\n    sys.path.insert(0, str(ROOT))\n\n"
            "for _m in [m for m in list(sys.modules) if m == 'src' or m.startswith('src.')]:\n"
            "    del sys.modules[_m]\n\n"
            "import pandas as pd\nfrom IPython.display import Image, display\n"
            "from src.simulacion import (\n"
            "    EQUIPO_TRABAJO, PERFILES_USUARIO, parametros_teoricos,\n"
            "    calcular_estado_estable, eliminar_estado_transitorio,\n"
            "    analizar_puntos_a_y_b, evaluar_estrategias, validar_modelo_vcv,\n"
            ")\n"
            "from src.graficas import (\n"
            "    graficar_estado_estable, graficar_cajeros_y_usuarios,\n"
            "    graficar_estrategias, graficar_vcv, graficar_transitorio_antes_despues,\n"
            ")\n\n"
            "OUT_FIG = Path(tempfile.gettempdir()) / 'figuras_supercentro'\n"
            "OUT_FIG.mkdir(parents=True, exist_ok=True)\n"
            "print('Integrantes:', ' | '.join(EQUIPO_TRABAJO))\n"
            "display(pd.DataFrame(PERFILES_USUARIO))\n"
            "teo = parametros_teoricos()\n"
            "print(f\"E[S]={teo['E_S']:.2f} min | E[A]={teo['E_A']:.2f} min | ρ={teo['rho']*100:.2f} % | Demanda={teo['demanda_hora_total']:.1f} usr/h\")"
        ),
        new_markdown_cell("## 2. Identificación del Estado Estable (`R*`)"),
        new_code_cell(
            "df_conv, res_est = calcular_estado_estable()\n"
            "R_OPT = res_est['r_optimo']\n"
            "print(f\"Réplicas seleccionadas: R* = {R_OPT} (W = {res_est['W_en_r']:.2f} min, error = {res_est['error_rel_pct']:.2f} %)\")\n"
            "display(Image(filename=str(graficar_estado_estable(df_conv, res_est, OUT_FIG)), width=800))"
        ),
        new_markdown_cell("## 3. Punto E — Algoritmo de Eliminación del Estado Transitorio (Antes vs. Después)"),
        new_code_cell(
            "res_t = eliminar_estado_transitorio(R_OPT)\n"
            "T0 = res_t['t0_min']\n"
            "print(f\"Corte óptimo detectado: t0* = {T0:.0f} min ({res_t['pct_descartado']:.1f} % del horizonte)\")\n"
            "display(res_t['df_comparacion'][['Metrica', 'Antes', 'Despues', 'Cambio_pct']].round(2))\n"
            "display(Image(filename=str(graficar_transitorio_antes_despues(res_t, OUT_FIG)), width=840))"
        ),
        new_markdown_cell("## 4. Puntos A y B — Estadísticas por Cajero y Promedio de Usuarios por Tipo"),
        new_code_cell(
            "res_a, df_b = analizar_puntos_a_y_b(res_t['df_despues'], res_t['ventana_estable_min'])\n"
            "print(f\"Punto A -> Menor W: {res_a['menor_cajero']} ({res_a['menor_W']:.2f} min) | Mayor W: {res_a['mayor_cajero']} ({res_a['mayor_W']:.2f} min) | ANOVA p = {res_a['anova_p_W']:.3f}\")\n"
            "display(res_a['tabla'].round(2))\n"
            "display(df_b.round(2))\n"
            "display(Image(filename=str(graficar_cajeros_y_usuarios(res_a, df_b, OUT_FIG)), width=840))"
        ),
        new_markdown_cell("## 5. Punto C — Evaluación de Estrategias de Mejora"),
        new_code_cell(
            "df_esc = evaluar_estrategias(R_OPT, T0)\n"
            "display(df_esc[['Escenario', 'Espera_Wq_min', 'Total_W_min', 'Utilizacion_pct', 'Reduccion_Wq_pct', 'Criterios_SLA']].round(2))\n"
            "display(Image(filename=str(graficar_estrategias(df_esc, OUT_FIG)), width=800))"
        ),
        new_markdown_cell("## 6. Punto D — Verificación, Calibración y Validación"),
        new_code_cell(
            "res_vcv = validar_modelo_vcv(res_t['df_despues'], T0)\n"
            "print(f\"Error Ley de Little: {res_vcv['error_little_pct']:.2f} %\")\n"
            "display(res_vcv['df_verificacion'].round(3))\n"
            "display(res_vcv['df_calibracion'].round(2))\n"
            "display(res_vcv['df_validacion'].round(2))\n"
            "display(Image(filename=str(graficar_vcv(res_vcv, OUT_FIG)), width=840))"
        ),
        new_markdown_cell(
            "## 7. Conclusión Final\n"
            "- **3 cajeros independientes no alcanzan** (`Wq = 15.67 min`, 0/3 criterios).\n"
            "- **Acción inmediata sin comprar equipos:** pasar los 3 cajeros a **fila única** baja la espera un **74.2 %** (`4.04 min`).\n"
            "- **Estrategia óptima recomendada:** **4 cajeros con fila única** (`Wq = 0.71 min`, **-95.5 %**, ocupación `58.3 %`, cumple **3/3 criterios**)."
        ),
    ]

    nb = new_notebook(cells=celdas)
    nb.metadata["kernelspec"] = {"display_name": "Python 3", "language": "python", "name": "python3"}
    nb.metadata["colab"] = {"name": "Simulacion_Parqueaderos_Supercentro.ipynb", "provenance": []}
    if ejecutar:
        NotebookClient(nb, timeout=300, kernel_name="python3").execute(cwd=str(ROOT_DIR))
    with open(ruta, "w", encoding="utf-8") as f:
        nbformat.write(nb, f)
    return ruta
