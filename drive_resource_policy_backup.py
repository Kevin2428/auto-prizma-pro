from __future__ import annotations

from dataclasses import dataclass, field
import os
import re
import unicodedata
from typing import Any, Iterable


@dataclass(frozen=True)
class WarningItem:
    code: str
    message: str


@dataclass
class MatchResult:
    entry: dict[str, Any] | None = None
    warnings: list[WarningItem] = field(default_factory=list)
    error: str | None = None


@dataclass
class DriveChoiceResult:
    item: dict[str, Any] | None = None
    warnings: list[WarningItem] = field(default_factory=list)
    error: str | None = None


def _base_name(value: Any) -> str:
    text = str(value or "").replace("\\", "/")
    return text.rsplit("/", 1)[-1].strip()


def _normalized_name(value: Any) -> str:
    text = unicodedata.normalize("NFC", _base_name(value))
    text = text.replace("–", "-").replace("—", "-")
    text = re.sub(r"\s+", " ", text).strip()
    return text.casefold()


def _extension(value: Any) -> str:
    return os.path.splitext(_base_name(value))[1].lower()


def _size(value: Any) -> int | None:
    try:
        if value is None:
            return None
        return int(value)
    except (TypeError, ValueError):
        return None


def _same_entry(a: dict[str, Any], b: dict[str, Any]) -> bool:
    return (
        _normalized_name(a.get("nombre")) == _normalized_name(b.get("nombre"))
        and _size(a.get("tamano")) == _size(b.get("tamano"))
        and str(a.get("miembro") or "") == str(b.get("miembro") or "")
    )


def _filter_expected_extension(
    entries: Iterable[dict[str, Any]],
    expected_extension: str,
) -> list[dict[str, Any]]:
    expected = str(expected_extension or "").lower().strip()
    items = list(entries)
    if expected not in {".h5p", ".pdf"}:
        return items
    filtered = [item for item in items if _extension(item.get("nombre")) == expected]
    return filtered or items


def choose_drive_item(
    items: Iterable[dict[str, Any]],
    expected_extension: str = "",
) -> DriveChoiceResult:
    candidates = _filter_expected_extension(items, expected_extension)
    if not candidates:
        return DriveChoiceResult(error="NO_DRIVE_RESOURCE")

    candidates = sorted(
        candidates,
        key=lambda item: (
            _normalized_name(item.get("nombre")),
            _size(item.get("tamano")) if _size(item.get("tamano")) is not None else -1,
        ),
    )

    if len(candidates) == 1:
        return DriveChoiceResult(item=candidates[0])

    signatures = {
        (_normalized_name(item.get("nombre")), _size(item.get("tamano")))
        for item in candidates
    }
    if len(signatures) == 1:
        selected = candidates[0]
        return DriveChoiceResult(
            item=selected,
            warnings=[
                WarningItem(
                    "IDENTICAL_DRIVE_DUPLICATE",
                    "La carpeta de Drive contiene recursos repetidos con el mismo nombre y tamaño; se continuará con uno de ellos y se deja la alerta para la revisión final.",
                )
            ],
        )

    return DriveChoiceResult(error="MULTIPLE_DIFFERENT_DRIVE_RESOURCES")


def match_drive_resource_to_zip(
    drive_item: dict[str, Any],
    zip_entries: Iterable[dict[str, Any]],
    expected_extension: str = "",
    fallback_entry: dict[str, Any] | None = None,
) -> MatchResult:
    entries = _filter_expected_extension(zip_entries, expected_extension)
    drive_name = _base_name(drive_item.get("nombre"))
    drive_name_key = _normalized_name(drive_name)
    drive_size = _size(drive_item.get("tamano"))

    same_name = [
        item for item in entries
        if _normalized_name(item.get("nombre")) == drive_name_key
    ]

    warnings: list[WarningItem] = []

    if same_name:
        if drive_size is not None:
            same_size = [item for item in same_name if _size(item.get("tamano")) == drive_size]
        else:
            same_size = []

        if len(same_size) == 1:
            return MatchResult(entry=same_size[0])

        if len(same_size) > 1:
            chosen = sorted(same_size, key=lambda item: str(item.get("miembro") or item.get("nombre") or ""))[0]
            warnings.append(
                WarningItem(
                    "IDENTICAL_ZIP_DUPLICATE",
                    f'El ZIP contiene más de una copia de "{drive_name}" con {drive_size} bytes; se continuará porque son equivalentes por nombre y tamaño.',
                )
            )
            return MatchResult(entry=chosen, warnings=warnings)

        if len(same_name) == 1:
            chosen = same_name[0]
            zip_size = _size(chosen.get("tamano"))
            if drive_size is not None and zip_size is not None and drive_size != zip_size:
                warnings.append(
                    WarningItem(
                        "SIZE_MISMATCH",
                        f'"{drive_name}" tiene {drive_size} bytes en Drive y {zip_size} bytes en el ZIP; se continuará y se mostrará esta diferencia al final.',
                    )
                )
            return MatchResult(entry=chosen, warnings=warnings)

        if fallback_entry is not None and any(_same_entry(fallback_entry, item) for item in same_name):
            chosen = next(item for item in same_name if _same_entry(fallback_entry, item))
            zip_size = _size(chosen.get("tamano"))
            if drive_size is not None and zip_size is not None and drive_size != zip_size:
                warnings.append(
                    WarningItem(
                        "SIZE_MISMATCH",
                        f'"{drive_name}" tiene {drive_size} bytes en Drive y {zip_size} bytes en el ZIP; se continuará y se mostrará esta diferencia al final.',
                    )
                )
            return MatchResult(entry=chosen, warnings=warnings)

        return MatchResult(error="AMBIGUOUS_SAME_NAME")

    # Si Drive truncó o cambió el nombre, el peso puede identificar de forma
    # inequívoca el binario local dentro del tipo esperado.
    if drive_size is not None:
        same_size = [item for item in entries if _size(item.get("tamano")) == drive_size]
        if len(same_size) == 1:
            chosen = same_size[0]
            warnings.append(
                WarningItem(
                    "NAME_MISMATCH_SIZE_MATCH",
                    f'Drive muestra "{drive_name}", pero por tamaño ({drive_size} bytes) se relacionó con "{_base_name(chosen.get("nombre"))}" del ZIP.',
                )
            )
            return MatchResult(entry=chosen, warnings=warnings)

    # La relación de la fila y su carpeta G ya es autoritativa. Si el resolver
    # histórico de esa misma actividad encuentra un único recurso local,
    # seguimos con él y dejamos visibles las diferencias como advertencias.
    if fallback_entry is not None:
        chosen = fallback_entry
        warnings.append(
            WarningItem(
                "NAME_MISMATCH_ACTIVITY_FALLBACK",
                f'Drive muestra "{drive_name}", mientras que la actividad quedó relacionada con "{_base_name(chosen.get("nombre"))}" en el ZIP.',
            )
        )
        zip_size = _size(chosen.get("tamano"))
        if drive_size is not None and zip_size is not None and drive_size != zip_size:
            warnings.append(
                WarningItem(
                    "SIZE_MISMATCH",
                    f'El recurso relacionado tiene {drive_size} bytes en Drive y {zip_size} bytes en el ZIP; se continuará y se mostrará esta diferencia al final.',
                )
            )
        return MatchResult(entry=chosen, warnings=warnings)

    return MatchResult(error="NO_ZIP_MATCH")


def detect_duplicate_assignments(
    resources: Iterable[dict[str, Any]],
) -> list[WarningItem]:
    groups: dict[tuple[str, int], list[dict[str, Any]]] = {}
    for resource in resources:
        name = _normalized_name(resource.get("archivo_drive"))
        size = _size(resource.get("tamano_drive"))
        if not name or size is None:
            continue
        groups.setdefault((name, size), []).append(resource)

    warnings: list[WarningItem] = []
    for (_, size), items in groups.items():
        if len(items) < 2:
            continue
        ordered = sorted(items, key=lambda item: (str(item.get("hoja") or ""), int(item.get("fila") or 0)))
        locations = ", ".join(
            f'{item.get("hoja", "")} fila {int(item.get("fila") or 0)}'
            for item in ordered
        )
        name = _base_name(ordered[0].get("archivo_drive"))
        warnings.append(
            WarningItem(
                "CROSS_ACTIVITY_DUPLICATE",
                f'El mismo recurso de Drive "{name}" ({size} bytes) aparece relacionado con más de una actividad: {locations}. El análisis continuó normalmente.',
            )
        )
    return warnings
