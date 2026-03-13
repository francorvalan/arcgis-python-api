# Survey123: exportar respuestas a Word con fotos

Este ejemplo agrega funciones reutilizables para usar en notebooks y generar un `.docx` desde la capa de respuestas de un *Survey123*.

## Flujo básico

```python
from arcgis.gis import GIS
from survey123_word_report import get_survey_layer, list_layer_fields, export_survey_to_docx

gis = GIS("home")
survey_item = gis.content.get("<survey_item_id>")
rel_fs = survey_item.related_items("Survey2Service", "forward")[0]
layer = get_survey_layer(rel_fs)

# 1) Ver columnas disponibles
campos = list_layer_fields(layer)
print(campos)

# 2) Elegir columnas y exportar
export_survey_to_docx(
    layer,
    output_path="reporte_survey.docx",
    id_field="id_sitios",
    columns=["fecha", "tecnico", "estado"],
)
```

## ¿Cómo seleccionar las columnas?

- Ejecuta `list_layer_fields(layer)` para listar campos válidos.
- Pasa la lista de nombres en `columns=[...]`.
- `id_field` se incluye siempre como primera columna (por defecto `id_sitios`).

## Notas

- Las fotos se toman desde los **attachments** de cada registro y se insertan en la última columna.
- Requiere `python-docx` (`pip install python-docx`).
