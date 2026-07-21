"""Check the footer variation table: cell background per value, white text.

Run: python test_footer.py
"""
import io

from pptx import Presentation

import main


def test_var_table():
    datos, metadata, secciones = main.read_excel(open('ejemplo.xlsx', 'rb').read())

    key = next(k for k in metadata if k in {
        (r.get('carrera'), r.get('segmento'), r.get('modalidad')) for r in datos
    })
    metadata[key]['var_des_prom'] = 'Muy Encima'
    metadata[key]['var_periodo_ant'] = 'Debajo'

    out = main.generate_ppt(open(main.TEMPLATE_PATH, 'rb').read(), datos, metadata, secciones)
    prs = Presentation(io.BytesIO(out))

    found = False
    for slide in prs.slides:
        shape = main.find_shape(slide, 'Table 11')
        cells = shape.table.rows[1].cells
        if cells[0].text != 'Muy Encima':
            continue
        found = True
        assert cells[0].fill.fore_color.rgb == main.VAR_COLORS['muy encima']
        assert cells[1].fill.fore_color.rgb == main.VAR_COLORS['debajo']
        for cell in cells:
            assert cell.text_frame.paragraphs[0].runs[0].font.color.rgb == main.VAR_TEXT_COLOR
    assert found, 'no slide got the injected variation values'


if __name__ == '__main__':
    test_var_table()
    print('ok')
