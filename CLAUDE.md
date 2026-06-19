# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Running the app

```bash
streamlit run main.py
```

Opens at `http://localhost:8501`.

## Architecture

Single-file app (`main.py`). No modules, no classes — purely functional pipeline:

```
Excel upload → read_excel() → generate_ppt() → download
PPT upload   ↗
```

**`read_excel(bytes)`** — reads three sheets from the uploaded Excel and returns `(datos, metadata, secciones)`:
- `datos`: one row per pain/hallazgo; grouped by `(carrera, segmento, modalidad)` to produce one slide per group.
- `metadata`: one row per `(carrera, segmento, modalidad)` key; used for header stats, comments, and variation tables.
- `secciones` (via `read_secciones`): the `secciones` sheet, the **single source of truth** for which careers belong to which section. Returns `{order, careers, by_career}`. Section membership is NOT derived from `tipo_carrera`/`area`.

**`generate_ppt(template_bytes, datos, metadata, secciones)`** — clones slide 0 of the uploaded `.pptx` template for each group, then calls a set of `update_*` functions to mutate each cloned slide in-place via `python-pptx` + direct lxml XML manipulation.

## Shape targeting

All shape mutations rely on **exact shape names** from `PPTmuestra.pptx`. If the template changes, these names must match:

| Shape name | Updater |
|---|---|
| `Tabla 87` | `update_main_table` — pain/hallazgo/accion rows |
| `Tabla 8` | `update_career_list` — rows 0-3 = section cards (dynamic, active section bolded); row 4 = careers of the active section, active career bolded |
| `Tabla 45` | `update_stats` — muestra/desertores/alumnos |
| `Table 11` | `update_var_table` — var_des_prom/var_periodo_ant |
| `Grupo 20` | `update_career_list` — arrow + highlight rect, moved to the active career in the list row (row 4) |
| `Elipse *` | `update_circles` — one per data row, red=profundizar / gray=no_accionable |
| `Rect *` (AUTO_SHAPE) | `update_rects` — one per data row, teal=transversal / dark=especifico |
| `Título` / `Titulo` in name + top < 5cm | `update_header` |
| `Título` / `Titulo` in name + top > 10cm | `update_comments` |

## Geometry constants

All EMU (English Metric Units) coordinates at the top of `main.py` were measured from `PPTmuestra.pptx`. If the template layout changes, these must be re-measured and updated:

- `CIRCLE_*` — circle position/step per row
- `RECT_*` — separator rectangle position/step per row
- `SIDEBAR_PARA_H` in `update_career_list` (row tops and the group coordinate offset are computed at runtime from the template, so they self-adjust)

## Key constraint: `set_cell`

`set_cell` preserves the first run's formatting (font, color, size) and strips all other runs. Never use `cell.text = ...` directly — it drops formatting.

## Excel schema

`datos` sheet columns: `tipo_carrera`, `area`, `carrera`, `segmento`, `modalidad`, `pain`, `hallazgo`, `accion`, `nivel`, `separador`

`metadata` sheet columns: `carrera`, `segmento`, `modalidad`, `desercion_carrera`, `desercion_prom`, `muestra`, `desertores`, `alumnos`, `comentarios`, `var_des_prom`, `var_periodo_ant`

`secciones` sheet columns: `seccion`, `carrera` — one row per career. Source of truth for the sidebar grouping; order of rows = display order. Up to 4 sections fit the template menu (rows 0-3 of `Tabla 8`; row 4 is the career list).
