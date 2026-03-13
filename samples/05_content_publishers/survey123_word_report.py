"""Utilidades para exportar registros de Survey123 a un reporte Word (.docx)."""

from __future__ import annotations

import random
from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Iterable, Sequence

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.shared import Inches


SYSTEM_FIELDS = {"OBJECTID", "GLOBALID", "SHAPE", "SHAPE_LENGTH", "SHAPE_AREA"}


def get_survey_layer(related_feature_service, layer_index: int = 0):
    """Devuelve la capa de respuestas del Feature Service relacionado a un Survey."""

    layers = getattr(related_feature_service, "layers", None)
    if not layers:
        raise ValueError("El Feature Service relacionado no contiene capas.")

    try:
        return layers[layer_index]
    except IndexError as exc:
        raise ValueError(
            f"No existe la capa en el índice {layer_index}. "
            f"Capas disponibles: 0..{len(layers) - 1}."
        ) from exc


def list_layer_fields(layer, include_system_fields: bool = False) -> list[str]:
    """Lista nombres de campos para elegir columnas del reporte."""

    field_names = [field["name"] for field in layer.properties.fields]
    if include_system_fields:
        return field_names
    return [name for name in field_names if name.upper() not in SYSTEM_FIELDS]


def export_survey_to_docx(
    layer,
    output_path: str | Path,
    id_field: str = "id_sitios",
    columns: Sequence[str] | None = None,
    where: str = "1=1",
    order_by_field: str | None = None,
    photo_column_title: str = "fotos",
    photo_mode: str = "all",
    orientation: str = "vertical",
    image_width_inches: float = 1.3,
) -> Path:
    """Genera un Word con atributos en columnas y fotos en la última columna."""

    available_fields = list_layer_fields(layer, include_system_fields=True)
    if id_field not in available_fields:
        raise ValueError(f"El campo id_field '{id_field}' no existe en la capa.")

    if columns is None:
        columns = list_layer_fields(layer)
    else:
        _validate_columns(columns, available_fields)

    if order_by_field is not None and order_by_field not in available_fields:
        raise ValueError(
            f"El campo order_by_field '{order_by_field}' no existe en la capa."
        )

    photo_mode = photo_mode.lower()
    valid_photo_modes = {"primera", "ultima", "random", "all"}
    if photo_mode not in valid_photo_modes:
        raise ValueError(
            "photo_mode no válido. Use uno de: " + ", ".join(sorted(valid_photo_modes))
        )

    orientation = orientation.lower()
    if orientation not in {"vertical", "horizontal"}:
        raise ValueError("orientation no válido. Use 'vertical' o 'horizontal'.")

    selected_columns = [id_field] + [col for col in columns if col != id_field]
    out_fields = ",".join(selected_columns)

    query_kwargs = {
        "where": where,
        "out_fields": out_fields,
        "return_geometry": False,
        "return_all_records": True,
    }
    if order_by_field is not None:
        query_kwargs["order_by_fields"] = order_by_field

    features = layer.query(**query_kwargs).features

    document = Document()
    _set_document_orientation(document, orientation)
    document.add_heading("Reporte Survey123", level=1)
    document.add_paragraph(f"Total de registros: {len(features)}")

    table = document.add_table(rows=1, cols=len(selected_columns) + 1)
    table.style = "Table Grid"

    header_cells = table.rows[0].cells
    for idx, field_name in enumerate(selected_columns):
        header_cells[idx].text = field_name
    header_cells[-1].text = photo_column_title

    oid_field = layer.properties.objectIdField
    with TemporaryDirectory() as tmp_dir:
        tmp_dir_path = Path(tmp_dir)

        for feature in features:
            row_cells = table.add_row().cells
            attrs = feature.attributes or {}

            for idx, field_name in enumerate(selected_columns):
                value = attrs.get(field_name)
                row_cells[idx].text = "" if value is None else str(value)

            oid_value = attrs.get(oid_field)
            if oid_value is None:
                row_cells[-1].text = "Sin OBJECTID"
                continue

            image_paths = _download_image_attachments(
                layer, oid_value, tmp_dir_path, photo_mode=photo_mode
            )
            if not image_paths:
                row_cells[-1].text = "Sin fotos"
                continue

            paragraph = row_cells[-1].paragraphs[0]
            for image_path in image_paths:
                paragraph.add_run().add_picture(
                    str(image_path), width=Inches(image_width_inches)
                )
                paragraph.add_run("\n")

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    document.save(str(output_path))
    return output_path


def _validate_columns(columns: Iterable[str], available_fields: Sequence[str]) -> None:
    missing = [col for col in columns if col not in available_fields]
    if missing:
        raise ValueError(
            "Las siguientes columnas no existen en la capa: "
            + ", ".join(sorted(missing))
        )


def _download_image_attachments(
    layer, object_id: int, temp_dir: Path, photo_mode: str = "all"
) -> list[Path]:
    """Descarga adjuntos de imagen del registro indicado."""

    attachments = _safe_get_attachments(layer, object_id)
    image_attachments = []

    for attachment in attachments:
        content_type = (attachment.get("contentType") or "").lower()
        if not content_type.startswith("image/"):
            continue
        image_attachments.append(attachment)

    image_attachments = _select_attachments(image_attachments, photo_mode)
    image_paths: list[Path] = []

    for attachment in image_attachments:
        attachment_id = attachment["id"]
        file_name = attachment.get("name") or f"attachment_{attachment_id}.jpg"
        target_path = temp_dir / f"{object_id}_{attachment_id}_{file_name}"

        download_result = layer.attachments.download(
            oid=object_id,
            attachment_id=attachment_id,
            save_path=str(temp_dir),
        )

        resolved_path = _materialize_download(download_result, target_path)
        if resolved_path and resolved_path.exists():
            image_paths.append(resolved_path)

    return image_paths


def _select_attachments(attachments: list[dict], photo_mode: str) -> list[dict]:
    if not attachments:
        return []

    if photo_mode == "primera":
        return [attachments[0]]
    if photo_mode == "ultima":
        return [attachments[-1]]
    if photo_mode == "random":
        return [random.choice(attachments)]
    return attachments


def _set_document_orientation(document: Document, orientation: str) -> None:
    section = document.sections[0]
    if orientation == "horizontal":
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width, section.page_height = section.page_height, section.page_width
        return

    section.orientation = WD_ORIENT.PORTRAIT


def _safe_get_attachments(layer, object_id: int):
    """Compatibilidad entre variantes de firma para get_list()."""

    try:
        return layer.attachments.get_list(oid=object_id)
    except TypeError:
        return layer.attachments.get_list(object_id=object_id)


def _materialize_download(download_result, target_path: Path) -> Path | None:
    """Normaliza la salida de `attachments.download` a un path existente."""

    if isinstance(download_result, (str, Path)):
        src_path = Path(download_result)
        if src_path.exists() and src_path != target_path:
            src_path.rename(target_path)
        return target_path if target_path.exists() else src_path

    if isinstance(download_result, list) and download_result:
        src_path = Path(download_result[0])
        if src_path.exists() and src_path != target_path:
            src_path.rename(target_path)
        return target_path if target_path.exists() else src_path

    if isinstance(download_result, bytes):
        target_path.write_bytes(download_result)
        return target_path

    if isinstance(download_result, BytesIO):
        target_path.write_bytes(download_result.getvalue())
        return target_path

    return target_path if target_path.exists() else None
