import io
import copy
import openpyxl
import streamlit as st
from collections import defaultdict
from pptx import Presentation
from pptx.util import Emu
from pptx.enum.shapes import MSO_SHAPE_TYPE
from lxml import etree

P_NS = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'

# Template geometry (EMU) — measured from PPTmuestra.pptx
CIRCLE_X = 11337402
CIRCLE_Y0 = 1203044   # top of first circle
CIRCLE_STEP = 731520  # vertical step per row
CIRCLE_SZ = 288000

FILL_RED  = f'<a:solidFill xmlns:a="{A_NS}"><a:srgbClr val="FF0000"/></a:solidFill>'
FILL_GRAY = f'<a:solidFill xmlns:a="{A_NS}"><a:schemeClr val="bg1"><a:lumMod val="65000"/></a:schemeClr></a:solidFill>'

# Separator rectangles (row indicators)
RECT_X    = 5313211
RECT_Y0   = 1614107   # top of first rect (row 1)
RECT_STEP = 716400    # = 1.99cm, matches data row height
RECT_SZ_W = 135089
RECT_SZ_H = 137865

FILL_TRANSVERSAL = f'<a:solidFill xmlns:a="{A_NS}"><a:srgbClr val="00B491"/></a:solidFill>'
FILL_ESPECIFICO  = f'<a:solidFill xmlns:a="{A_NS}"><a:schemeClr val="tx1"><a:lumMod val="75000"/><a:lumOff val="25000"/></a:schemeClr></a:solidFill>'


# ── Excel reading ────────────────────────────────────────────────────────────

def read_excel(excel_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))

    ws = wb['datos']
    headers = [c.value for c in ws[1]]
    datos = [
        dict(zip(headers, [c.value for c in row]))
        for row in ws.iter_rows(min_row=2)
        if any(c.value for c in row)
    ]

    ws = wb['metadata']
    headers = [c.value for c in ws[1]]
    meta = {}
    for row in ws.iter_rows(min_row=2):
        if any(c.value for c in row):
            d = dict(zip(headers, [c.value for c in row]))
            meta[(d['carrera'], d['segmento'], d['modalidad'])] = d

    return datos, meta


# ── Slide cloning ────────────────────────────────────────────────────────────

def clone_slide(output_prs, template_slide):
    """Append a clone of template_slide to output_prs, skipping OLE objects."""
    new_slide = output_prs.slides.add_slide(output_prs.slide_layouts[6])
    sp_tree = new_slide.shapes._spTree

    # Keep first 2 structural children (nvGrpSpPr, grpSpPr), remove the rest
    for child in list(sp_tree)[2:]:
        sp_tree.remove(child)

    for shape in template_slide.shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT:
            continue
        sp_tree.append(copy.deepcopy(shape._element))

    return new_slide


# ── Shape helpers ────────────────────────────────────────────────────────────

def find_shape(slide, name):
    return next((s for s in slide.shapes if s.name == name), None)


def set_cell(cell, text):
    """Overwrite cell text preserving first-run formatting."""
    tf = cell.text_frame
    if not tf.paragraphs:
        return
    para = tf.paragraphs[0]
    for p in tf.paragraphs[1:]:
        p._p.getparent().remove(p._p)
    if para.runs:
        para.runs[0].text = str(text) if text is not None else ''
        for r in para.runs[1:]:
            r._r.getparent().remove(r._r)
    else:
        para.add_run().text = str(text) if text is not None else ''


# ── Per-slide updaters ───────────────────────────────────────────────────────

def update_header(slide, carrera, segmento, modalidad, des_carrera, des_prom, pg, total):
    for shape in slide.shapes:
        if shape.has_text_frame and shape.top / 360000 < 5:
            if 'Título' in shape.name or 'Titulo' in shape.name:
                text = (
                    f"{carrera} – {segmento} – {modalidad}"
                    f"  |  Deserción Carrera: {des_carrera}"
                    f"   |   Deserción Prom.: {des_prom}\t    \t      ({pg} de {total})"
                )
                tf = shape.text_frame
                if tf.paragraphs and tf.paragraphs[0].runs:
                    tf.paragraphs[0].runs[0].text = text
                break


def update_comments(slide, comentarios):
    """comentarios can be a string with \n separating paragraphs."""
    for shape in slide.shapes:
        if shape.has_text_frame and shape.top / 360000 > 10:
            if 'Título' in shape.name or 'Titulo' in shape.name:
                tf = shape.text_frame
                lines = str(comentarios).split('\n') if comentarios else ['']

                # Ensure enough paragraphs
                while len(tf.paragraphs) < len(lines):
                    # Clone last paragraph
                    last_p = tf.paragraphs[-1]._p
                    new_p = copy.deepcopy(last_p)
                    last_p.addnext(new_p)

                # Remove excess paragraphs
                while len(tf.paragraphs) > len(lines):
                    tf.paragraphs[-1]._p.getparent().remove(tf.paragraphs[-1]._p)

                for para, line in zip(tf.paragraphs, lines):
                    if para.runs:
                        para.runs[0].text = line
                        for r in para.runs[1:]:
                            r._r.getparent().remove(r._r)
                    else:
                        para.add_run().text = line
                break


def update_main_table(slide, rows):
    """rows: list of (pain, hallazgo, accion) tuples."""
    shape = find_shape(slide, 'Tabla 87')
    if not shape or not shape.has_table:
        return
    table = shape.table

    # Adjust row count (header is row 0, data starts at row 1)
    while len(table.rows) - 1 < len(rows):
        new_tr = copy.deepcopy(table.rows[1]._tr)  # clone first data row
        table._tbl.append(new_tr)

    while len(table.rows) - 1 > len(rows):
        table._tbl.remove(table.rows[len(table.rows) - 1]._tr)

    for i, (pain, hallazgo, accion) in enumerate(rows):
        r = table.rows[i + 1]
        set_cell(r.cells[0], pain)
        set_cell(r.cells[1], hallazgo)
        set_cell(r.cells[2], accion)


def update_circles(slide, niveles):
    """niveles: list of 'profundizar' | 'no_accionable'."""
    circles = [s for s in slide.shapes if 'Elipse' in s.name]
    if not circles:
        return

    template_sp = circles[0]._element

    # Remove excess
    for c in circles[len(niveles):]:
        c._element.getparent().remove(c._element)

    # Add missing (clone first circle)
    circles = [s for s in slide.shapes if 'Elipse' in s.name]
    while len(circles) < len(niveles):
        new_sp = copy.deepcopy(template_sp)
        slide.shapes._spTree.append(new_sp)
        circles = [s for s in slide.shapes if 'Elipse' in s.name]

    # Position + color
    for i, circle in enumerate(circles):
        spPr = circle._element.find(f'{{{P_NS}}}spPr')
        if spPr is None:
            continue
        xfrm = spPr.find(f'{{{A_NS}}}xfrm')
        if xfrm is not None:
            off = xfrm.find(f'{{{A_NS}}}off')
            if off is not None:
                off.set('x', str(CIRCLE_X))
                off.set('y', str(CIRCLE_Y0 + i * CIRCLE_STEP))

        # Replace fill
        for tag in [f'{{{A_NS}}}solidFill', f'{{{A_NS}}}noFill', f'{{{A_NS}}}gradFill']:
            for el in spPr.findall(tag):
                spPr.remove(el)

        nivel = niveles[i] if i < len(niveles) else 'no_accionable'
        fill_xml = FILL_RED if (nivel or '').lower() == 'profundizar' else FILL_GRAY
        new_fill = etree.fromstring(fill_xml)
        geom = spPr.find(f'{{{A_NS}}}prstGeom')
        if geom is not None:
            geom.addnext(new_fill)
        else:
            spPr.append(new_fill)


def update_rects(slide, separadores):
    """separadores: list of 'transversal' | 'especifico'"""
    rects = [s for s in slide.shapes if 'Rect' in s.name and s.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE]
    if not rects:
        return

    template_sp = rects[0]._element

    # Remove excess
    for r in rects[len(separadores):]:
        r._element.getparent().remove(r._element)

    # Add missing
    rects = [s for s in slide.shapes if 'Rect' in s.name and s.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE]
    while len(rects) < len(separadores):
        new_sp = copy.deepcopy(template_sp)
        slide.shapes._spTree.append(new_sp)
        rects = [s for s in slide.shapes if 'Rect' in s.name and s.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE]

    # Position + color
    for i, rect in enumerate(rects):
        spPr = rect._element.find(f'{{{P_NS}}}spPr')
        if spPr is None:
            continue
        xfrm = spPr.find(f'{{{A_NS}}}xfrm')
        if xfrm is not None:
            off = xfrm.find(f'{{{A_NS}}}off')
            ext = xfrm.find(f'{{{A_NS}}}ext')
            if off is not None:
                off.set('x', str(RECT_X))
                off.set('y', str(RECT_Y0 + i * RECT_STEP))
            if ext is not None:
                ext.set('cx', str(RECT_SZ_W))
                ext.set('cy', str(RECT_SZ_H))

        for tag in [f'{{{A_NS}}}solidFill', f'{{{A_NS}}}noFill', f'{{{A_NS}}}gradFill']:
            for el in spPr.findall(tag):
                spPr.remove(el)

        sep = (separadores[i] or '').lower() if i < len(separadores) else 'transversal'
        fill_xml = FILL_ESPECIFICO if sep == 'especifico' else FILL_TRANSVERSAL
        new_fill = etree.fromstring(fill_xml)
        geom = spPr.find(f'{{{A_NS}}}prstGeom')
        if geom is not None:
            geom.addnext(new_fill)
        else:
            spPr.append(new_fill)


def update_career_list(slide, carrera, tipo_carrera):
    """Move the arrow + highlight rect to the row matching the current carrera."""
    # Geometry constants (EMU) — from PPTmuestra.pptx analysis
    GROUP_LOCAL_OFFSET = 112549  # absolute_y + this = local_y in group XML
    ROW3_TOP  = 1795546          # cumulative top of row 3 in Tabla 8
    ROW3_H    = 296091
    ROW5_TOP  = 2529921          # cumulative top of row 5 (No Masivas list)
    PARA_H    = 265045           # height per career paragraph in row 5
    RECT_H    = 199429
    ARROW_H   = 318221

    tipo_lower = (tipo_carrera or '').lower()
    no_masivas = 'no masivas' in tipo_lower

    if no_masivas:
        # Find carrera index in row 5 paragraphs
        shape = find_shape(slide, 'Tabla 8')
        if not shape or not shape.has_table:
            return
        paras = shape.table.rows[5].cells[0].text_frame.paragraphs
        names = [p.text.strip() for p in paras]
        try:
            idx = names.index(carrera)
        except ValueError:
            return
        center_y = int(ROW5_TOP + (idx + 0.5) * PARA_H)
    else:
        # Update row 3 text and point there
        shape = find_shape(slide, 'Tabla 8')
        if shape and shape.has_table:
            set_cell(shape.table.rows[3].cells[0], carrera)
        center_y = ROW3_TOP + ROW3_H // 2

    # Move the highlight rect and arrow (both are inside Grupo 20)
    grupo = find_shape(slide, 'Grupo 20')
    if grupo is None:
        return

    for child in grupo.shapes:
        is_rect  = 'Rect' in child.name
        is_arrow = 'Flecha' in child.name or 'flecha' in child.name
        if not (is_rect or is_arrow):
            continue

        child_h = RECT_H if is_rect else ARROW_H
        new_abs_top = center_y - child_h // 2
        new_local_y = new_abs_top + GROUP_LOCAL_OFFSET

        spPr = child._element.find(f'{{{P_NS}}}spPr')
        if spPr is None:
            continue
        xfrm = spPr.find(f'{{{A_NS}}}xfrm')
        if xfrm is None:
            continue
        off = xfrm.find(f'{{{A_NS}}}off')
        if off is not None:
            off.set('y', str(new_local_y))


def update_stats(slide, muestra, desertores, alumnos):
    shape = find_shape(slide, 'Tabla 45')
    if not shape or not shape.has_table:
        return
    t = shape.table
    set_cell(t.rows[0].cells[1], muestra)
    set_cell(t.rows[1].cells[1], desertores)
    set_cell(t.rows[2].cells[1], alumnos)


def update_var_table(slide, var_des_prom, var_periodo_ant):
    shape = find_shape(slide, 'Table 11')
    if not shape or not shape.has_table:
        return
    t = shape.table
    set_cell(t.rows[1].cells[0], var_des_prom)
    set_cell(t.rows[1].cells[1], var_periodo_ant)


# ── Main generator ───────────────────────────────────────────────────────────

def generate_ppt(template_bytes, datos, metadata):
    template_prs = Presentation(io.BytesIO(template_bytes))
    template_slide = template_prs.slides[0]

    output_prs = Presentation(io.BytesIO(template_bytes))

    groups = defaultdict(list)
    for row in datos:
        key = (row.get('carrera'), row.get('segmento'), row.get('modalidad'))
        groups[key].append(row)

    for i, (key, rows) in enumerate(groups.items()):
        carrera, segmento, modalidad = key
        meta = metadata.get(key, {})
        slide = output_prs.slides[0] if i == 0 else clone_slide(output_prs, template_slide)

        update_header(slide, carrera, segmento, modalidad,
                      meta.get('desercion_carrera', ''),
                      meta.get('desercion_prom', ''),
                      1, 1)

        table_rows = [(r.get('pain'), r.get('hallazgo'), r.get('accion')) for r in rows]
        update_main_table(slide, table_rows)

        niveles     = [r.get('nivel', 'no_accionable') for r in rows]
        separadores = [r.get('separador', 'transversal') for r in rows]
        update_circles(slide, niveles)
        update_rects(slide, separadores)

        tipo_carrera = rows[0].get('tipo_carrera', '')
        update_career_list(slide, carrera, tipo_carrera)

        update_stats(slide, meta.get('muestra'), meta.get('desertores'), meta.get('alumnos'))
        update_var_table(slide, meta.get('var_des_prom'), meta.get('var_periodo_ant'))
        update_comments(slide, meta.get('comentarios'))

    buf = io.BytesIO()
    output_prs.save(buf)
    return buf.getvalue()


# ── Streamlit UI ─────────────────────────────────────────────────────────────

st.set_page_config(page_title="AutoPPT", page_icon="📊")
st.title("📊 AutoPPT")
st.caption("Cargá tu Excel con datos y tu plantilla PPT. Descargá el resultado.")

with st.expander("📋 Formato del Excel requerido"):
    st.markdown("""
**Hoja `datos`** — una fila por pain/hallazgo:

| tipo_carrera | carrera | segmento | modalidad | pain | hallazgo | accion | nivel |
|---|---|---|---|---|---|---|---|
| Masivas | Ing. Electrónica | Nuevos | Presencial | Pagos... | Se le generó... | No aplica | no_accionable |

`nivel`: `profundizar` → círculo rojo · `no_accionable` → círculo gris

---

**Hoja `metadata`** — una fila por combinación carrera × segmento × modalidad:

| carrera | segmento | modalidad | desercion_carrera | desercion_prom | muestra | desertores | alumnos | comentarios | var_des_prom | var_periodo_ant |
|---|---|---|---|---|---|---|---|---|---|---|

`comentarios`: usá Alt+Enter dentro de la celda para múltiples líneas.
    """)

col1, col2 = st.columns(2)
with col1:
    excel_file = st.file_uploader("📄 Excel con datos", type=["xlsx"])
with col2:
    ppt_file = st.file_uploader("📑 Plantilla PPT", type=["pptx"])

if excel_file and ppt_file:
    try:
        datos, metadata = read_excel(excel_file.read())
        keys = list({(r.get('carrera'), r.get('segmento'), r.get('modalidad')) for r in datos})
        st.info(f"Se generarán **{len(keys)} slides** para {len(datos)} filas de datos.")

        if st.button("🚀 Generar PPT"):
            with st.spinner("Generando..."):
                result = generate_ppt(ppt_file.read(), datos, metadata)
            st.success("¡Listo!")
            st.download_button(
                label="⬇️ Descargar PPT generado",
                data=result,
                file_name="resultado.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )
    except Exception as e:
        st.error(f"Error: {e}")
        import traceback
        st.code(traceback.format_exc())
