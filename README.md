# AutoPPT

Genera presentaciones de PowerPoint a partir de un Excel de datos y una plantilla
`.pptx`. Una app de Streamlit que clona un slide modelo, una vez por cada combinación
de **carrera × segmento × modalidad**, y rellena tablas, círculos, comentarios y la
barra lateral de carreras a partir del Excel.

```
Excel (datos)  ─┐
                ├─►  read_excel()  ─►  generate_ppt()  ─►  resultado.pptx
Plantilla .pptx ┘
```

## Cómo correrlo

```bash
streamlit run main.py
```

Abre en `http://localhost:8501`. Subís el Excel y la plantilla `.pptx`, y descargás el
PPT generado.

Dependencias: `streamlit`, `python-pptx`, `openpyxl`, `lxml`.

```bash
pip install streamlit python-pptx openpyxl lxml
```

## Formato del Excel

El Excel tiene **tres hojas**: `datos`, `metadata` y `secciones`.

### Hoja `datos`

Una fila por cada pain/hallazgo. Las filas se agrupan por
`(carrera, segmento, modalidad)` → **un slide por grupo**.

| columna | descripción |
|---|---|
| `tipo_carrera` | clasificación interna (no decide la barra lateral) |
| `area` | área de la carrera (informativo) |
| `carrera` | nombre de la carrera |
| `segmento` | p. ej. `Nuevos` |
| `modalidad` | p. ej. `Presencial` |
| `pain` | dolor / hallazgo (fila de la tabla principal) |
| `hallazgo` | detalle del hallazgo |
| `accion` | acción propuesta |
| `nivel` | `profundizar` → círculo rojo · `no_accionable` → círculo gris |
| `separador` | `transversal` → rect. verde · `especifico` → rect. oscuro |

### Hoja `metadata`

Una fila por cada combinación `(carrera, segmento, modalidad)`. Aporta los datos de
cabecera, los comentarios y la tabla de variaciones.

| columna | descripción |
|---|---|
| `carrera`, `segmento`, `modalidad` | clave que matchea con `datos` |
| `desercion_carrera`, `desercion_prom` | porcentajes de cabecera |
| `muestra`, `desertores`, `alumnos` | estadísticas |
| `comentarios` | texto libre; usá **Alt+Enter** para separar líneas |
| `var_des_prom`, `var_periodo_ant` | celdas de la tabla de variaciones |

### Hoja `secciones` — agrupación de la barra lateral

Define **a qué sección pertenece cada carrera**. Es la **fuente única de verdad**:
la agrupación NO se deriva de `tipo_carrera` ni de `area`, se escribe acá a mano.

| seccion | carrera |
|---|---|
| Carreras Masivas | Contabilidad |
| Carreras Masivas | Ing. de Sistemas |
| Salud | Medicina |
| Salud | Enfermería |

Para cada slide, la barra lateral muestra **las tarjetas de todas las secciones** y,
debajo, **solo las carreras de la sección a la que pertenece esa carrera**, resaltando
la carrera actual:

```
╔════════════════════╗
║ Carreras Masivas   ║
╠════════════════════╣
║ Salud           ◀  ║  ← sección activa (negrita + color)
╠════════════════════╣
║ Carreras No Masivas║
╚════════════════════╝
······················
  Medicina         ◀    ← carrera activa (negrita + barra/flecha)
  Farm. y Bioq
  Enfermería
  Psicología
```

**Reglas:**
- El **orden de las filas** define el orden de aparición (secciones y carreras).
- La barra + flecha apuntan a la **carrera activa**; la **sección activa** se resalta en
  negrita y color.
- El menú soporta hasta **5 secciones** (las filas 0-4 del template). Si hay más, las
  que sobran no aparecen como tarjeta (la app lo avisa).
- Si una carrera de `datos` no está en `secciones`, la app avisa y esa carrera no se
  agrupa.

## Arquitectura

App de **un solo archivo** (`main.py`), pipeline puramente funcional — sin clases ni
módulos, por decisión de diseño (ver `CLAUDE.md`).

| función | rol |
|---|---|
| `read_excel` | lee las hojas `datos`, `metadata` y delega en `read_secciones` |
| `read_secciones` | construye `{order, careers, by_career}` desde la hoja `secciones` |
| `clone_slide` | clona el slide modelo (salta objetos OLE) |
| `set_cell` / `set_cell_lines` | escriben texto preservando el formato del template |
| `update_*` | mutan cada shape del slide clonado (cabecera, tablas, círculos, etc.) |
| `update_career_list` | renderiza la barra lateral (secciones + carreras de la activa) |
| `generate_ppt` | orquesta: agrupa `datos`, clona un slide por grupo, llama a los `update_*` |

### Dependencia de la plantilla

Todas las mutaciones dependen de los **nombres exactos de los shapes** y de
**constantes de geometría en EMU** medidas sobre `PPTmuestra.pptx`. Si cambia el layout
de la plantilla, hay que re-medir.

| shape | updater |
|---|---|
| `Tabla 87` | `update_main_table` — filas pain/hallazgo/accion |
| `Tabla 8` | `update_career_list` — menú de secciones + lista de carreras |
| `Grupo 20` | `update_career_list` — barra + flecha que apuntan a la carrera activa |
| `Tabla 45` | `update_stats` — muestra/desertores/alumnos |
| `Table 11` | `update_var_table` — variaciones |
| `Elipse *` | `update_circles` — un círculo por fila de datos |
| `Rect *` (AUTO_SHAPE) | `update_rects` — un rectángulo separador por fila |
| `Título`/`Titulo`, top < 5cm | `update_header` |
| `Título`/`Titulo`, top > 10cm | `update_comments` |

> **`set_cell` preserva el formato del primer run** y descarta el resto. Nunca uses
> `cell.text = ...` directo: pierde fuente, color y tamaño.

## Limitaciones conocidas

- La fila 5 de `Tabla 8` fue dibujada para ~8 carreras. Si una sección tiene muchas
  más, el texto puede desbordar visualmente; el ajuste fino de altura es de diseño en
  el `.pptx`.
- El menú de secciones está acotado a 5 tarjetas (estructura del template).
- `ejemplo.xlsx` incluye las tres hojas con datos de muestra para probar la app.
