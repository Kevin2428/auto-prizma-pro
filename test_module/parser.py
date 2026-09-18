from __future__ import annotations

import html
import re
from io import BytesIO

from docx import Document as DocxDocument

from .models import DocumentTest, Option, Pair, Question, Section, TestDocument, ValidationIssue, Severity


_ACTIVITY_HEADERS = {
    "trueOrFalse": re.compile(
        r"^(?:preguntas?\s+)?(?:verdadero\s*(?:/|y|o)\s*falso|true\s*(?:/|or)\s*false|vrai\s*(?:/|ou)\s*faux)$",
        re.I,
    ),
    "multipleChoice": re.compile(
        r"^(?:preguntas?\s+de\s+)?(?:única\s+respuesta|unica\s+respuesta|selección\s+(?:múltiple|unica)|seleccion\s+(?:multiple|unica)|opción\s+múltiple|opcion\s+multiple|multiple\s+choice|single\s+(?:answer|choice)|choix\s+multiple|réponse\s+unique|reponse\s+unique)$",
        re.I,
    ),
    "matching": re.compile(
        r"^(?:emparejamiento|matching|match\s+the\s+following|relacione?|relaciona|asocie|association|appariement|reliez|associez)$",
        re.I,
    ),
    "writeAnswer": re.compile(
        r"^(?:pregunta\s+abierta|respuesta\s+corta|resposta\s+aberta|open\s+(?:question|answer)|short\s+answer|question\s+ouverte|réponse\s+courte|reponse\s+courte)$",
        re.I,
    ),
}

_NUMBERED = re.compile(r"^(\d+)\s*(?:[.)]|[-.]|-)\s*(.+)$")
_WEEK = re.compile(r"\b(?:semana|week|semaine)\s*[/\-]?\s*(\d+)\b", re.I)
_CUT = re.compile(r"\b(?:corte|cut|periodo|période)\s*[/\-]?\s*(\d+)\b", re.I)


def _clean(value: str) -> str:
    value = html.unescape(str(value or ""))
    replacements = {
        "&sup0;": "⁰",
        "&sup1;": "¹",
        "&sup2;": "²",
        "&sup3;": "³",
        "&sup4;": "⁴",
        "&sup5;": "⁵",
        "&sup6;": "⁶",
        "&sup7;": "⁷",
        "&sup8;": "⁸",
        "&sup9;": "⁹",
        "&minus;": "−",
        "&times;": "×",
        "&divide;": "÷",
        "&radic;": "√",
        "&le;": "≤",
        "&ge;": "≥",
        "&ne;": "≠",
    }
    for source, target in replacements.items():
        value = value.replace(source, target)
    return re.sub(r"[ \t\xa0]+", " ", value.replace("\ufeff", "").strip())


def _lines(text: str) -> list[str]:
    return [_clean(line) for line in str(text or "").replace("\r", "").split("\n") if _clean(line)]


def _is_header(line: str, code: str | None = None) -> bool:
    if code:
        return bool(_ACTIVITY_HEADERS[code].match(_clean(line)))
    return any(pattern.match(_clean(line)) for pattern in _ACTIVITY_HEADERS.values())


def _language(text: str) -> str:
    normalized = _clean(text).lower()
    if re.search(r"\b(?:vrai|faux|reliez|associez|question ouverte|réponse)\b", normalized):
        return "fr"
    if re.search(r"\b(?:true|false|matching|open question|short answer)\b", normalized):
        return "en"
    return "es"


def _split_blocks(text: str) -> list[str]:
    raw = str(text or "").replace("\r", "")
    lines = [l.strip() for l in raw.split("\n") if l.strip()]
    if not lines:
        return [raw]

    blocks = []
    current = []
    has_questions = False

    is_boundary_start = re.compile(
        r"^(?:semana\s*[\/\-]?\s*corte|semana\s*[\/\-]?\s*\d+|week\s+\d+|semaine\s+\d+|test\s+\d+|(?:t[íi]tulo|titulo)\s+del\s+test)\b",
        re.I,
    )
    is_q_header = re.compile(
        r"^(?:preguntas?\s+)?(?:verdadero|true|vrai|única|unica|selección|seleccion|opción|opcion|multiple|single|choix|réponse|reponse|emparejamiento|matching|relacione|asocie|pregunta\s+abierta|respuesta\s+corta|open\s+question)\b",
        re.I,
    )

    for idx, line in enumerate(lines):
        if is_q_header.match(line):
            has_questions = True

        future_has_questions = any(is_q_header.match(fut) for fut in lines[idx:])

        if is_boundary_start.match(line) and has_questions and future_has_questions:
            blocks.append("\n".join(current))
            current = [line]
            has_questions = False
        else:
            current.append(line)

    if current:
        blocks.append("\n".join(current))

    return blocks if blocks else [raw]


def _section_lines(lines: list[str], code: str) -> list[str]:
    try:
        start = next(i for i, line in enumerate(lines) if _is_header(line, code))
    except StopIteration:
        return []
    end = len(lines)
    for index in range(start + 1, len(lines)):
        if _is_header(lines[index]):
            end = index
            break
    return lines[start + 1 : end]


def _extract_title(lines: list[str], index: int) -> str:
    title_regex = re.compile(
        r"^(?:t[íi]tulo|titulo)\s+(?:del\s+test|\d+)|nombre\s+del\s+test|test\s+title|title|titre",
        re.I,
    )
    for position, line in enumerate(lines):
        clean_line = _clean(line)
        # Caso 1: "Título del TEST: Nombre del test" o "Título del TEST - Nombre del test"
        colon_match = re.match(
            r"^(?:t[íi]tulo|titulo)\s+(?:del\s+test|\d+)[ \t]*[:\-–—][ \t]*(.+)$",
            clean_line,
            re.I,
        )
        if colon_match and colon_match.group(1).strip():
            candidate = colon_match.group(1).strip().strip(".'\"")
            if len(candidate) > 2:
                return candidate

        # Caso 2: "Título del TEST" solo en la línea, y el título en la línea siguiente
        if title_regex.match(clean_line):
            # Verificar si hay contenido adicional en la misma línea
            sub = title_regex.sub("", clean_line).strip(" :-–—\t.'\"")
            if len(sub) > 2:
                return sub
            if position + 1 < len(lines):
                next_line = lines[position + 1].strip().strip(".'\"")
                if not _is_header(next_line) and not _WEEK.search(next_line) and not _CUT.search(next_line):
                    return next_line

    # Fallback inteligente
    for line in lines[:15]:
        clean_line = _clean(line)
        if not _is_header(clean_line) and not _WEEK.search(clean_line) and not _CUT.search(clean_line) and not re.match(r"^(?:indicaciones|instructions?|programme|programa|campo de conocimientos|semana|corte)\b", clean_line, re.I):
            candidate = clean_line.strip().strip(".'\"")
            if 4 < len(candidate) < 160:
                return candidate

    return f"Test {index}"


def _extract_description(lines: list[str]) -> str:
    for position, line in enumerate(lines):
        if re.match(r"^(?:estimados estudiantes|dear students|chers étudiants|chers etudiants|instructions?|indicaciones?)", line, re.I):
            description = []
            for candidate in lines[position + 1 :]:
                if _is_header(candidate):
                    break
                description.append(candidate)
            if description:
                return _clean(" ".join(description))
    return "Responde con base en los contenidos estudiados."


_TF_OPTIONS_LINE = re.compile(
    r"^(?:(?:\([ xX]?\)|\[[ xX]?\])\s*)?(?:verdadero|falso|true|false|vrai|faux|v\b|f\b)"
    r"(?:\s*[\/\-]\s*|\s+)"
    r"(?:(?:\([ xX]?\)|\[[ xX]?\])\s*)?(?:verdadero|falso|true|false|vrai|faux|v\b|f\b)\s*$",
    re.I,
)

_TF_TRAILING_PATTERN = re.compile(
    r"\s*(?:(?:\([ xX]?\)|\[[ xX]?\])\s*)?(?:verdadero|falso|true|false|vrai|faux|v\b|f\b)"
    r"(?:\s*[\/\-]\s*|\s+)"
    r"(?:(?:\([ xX]?\)|\[[ xX]?\])\s*)?(?:verdadero|falso|true|false|vrai|faux|v\b|f\b)\s*[.]?\s*$",
    re.I,
)


def _clean_tf_prompt_text(text: str) -> str:
    cleaned = _clean(text)
    return _TF_TRAILING_PATTERN.sub("", cleaned).strip()


def _parse_true_false(lines: list[str]) -> list[Question]:
    questions: list[Question] = []
    current: Question | None = None
    for line in lines:
        clean_line = _clean(line)
        match = _NUMBERED.match(clean_line)
        if match:
            if current:
                questions.append(current)
            prompt = _clean(match.group(2))
            answer = None
            marker = re.search(
                r"(?:\((V|F|True|False|Verdadero|Falso|Vrai|Faux)\)|\[(?:R|A)\s*:\s*(Verdadero|Verdadeiro|Falso|True|False|Vrai|Faux|V|F)\s*\])\s*$",
                prompt,
                re.I,
            )
            if marker:
                answer_text = marker.group(1) or marker.group(2)
                answer = answer_text.lower() in {"v", "true", "verdadero", "vrai"}
                prompt = prompt[: marker.start()]
            if answer is None:
                if re.search(r"(?:\([xX]\)|\[[xX]\])\s*(?:verdadero|true|vrai|v\b)", prompt, re.I):
                    answer = True
                elif re.search(r"(?:\([xX]\)|\[[xX]\])\s*(?:falso|false|faux|f\b)", prompt, re.I):
                    answer = False
            prompt = _clean_tf_prompt_text(prompt)
            current = Question(prompt=prompt, answer=answer)
            continue
        if current is None:
            continue

        if _TF_OPTIONS_LINE.match(clean_line):
            if current.answer is None:
                if re.search(r"(?:\([xX]\)|\[[xX]\])\s*(?:verdadero|true|vrai|v\b)", clean_line, re.I):
                    current.answer = True
                elif re.search(r"(?:\([xX]\)|\[[xX]\])\s*(?:falso|false|faux|f\b)", clean_line, re.I):
                    current.answer = False
            continue

        marker = re.search(r"\[(?:R|A)\s*:\s*(Verdadero|Verdadeiro|Falso|True|False|Vrai|Faux|V|F)\s*\]", clean_line, re.I)
        if marker:
            current.answer = marker.group(1).lower() in {"v", "true", "verdadero", "vrai"}
            continue
        marker = re.match(r"^(?:answer|respuesta|resp|réponse|reponse)\s*:\s*(.+)$", clean_line, re.I)
        if marker and marker.group(1).strip().lower() in {"v", "f", "true", "false", "verdadero", "falso", "vrai", "faux"}:
            current.answer = marker.group(1).strip().lower() in {"v", "true", "verdadero", "vrai"}
            continue
        if clean_line.strip().lower() in {"v", "f", "true", "false", "verdadero", "falso", "vrai", "faux"} and current.answer is None:
            current.answer = clean_line.strip().lower() in {"v", "true", "verdadero", "vrai"}
            continue

        clean_part = _clean_tf_prompt_text(clean_line)
        if clean_part:
            current.prompt = _clean(f"{current.prompt} {clean_part}")
    if current:
        questions.append(current)
    for q in questions:
        q.prompt = _clean_tf_prompt_text(q.prompt)
    return questions


def _parse_multiple_choice(lines: list[str]) -> list[Question]:
    questions: list[Question] = []
    current: Question | None = None
    for line in lines:
        match = _NUMBERED.match(line)
        if match:
            if current:
                questions.append(current)
            current = Question(prompt=_clean(match.group(2)))
            continue
        if current is None:
            continue
        option = re.match(r"^([a-d])\s*[).]\s+(.+)$", line, re.I)
        if option:
            current.options.append(Option(option.group(1).lower(), _clean(option.group(2))))
            continue
        marker = re.search(r"\[(?:R|A)\s*:\s*([^\]]+)\]", line, re.I)
        if not marker:
            marker = re.match(r"^(?:answer|respuesta|resp|correcta|opción correcta|opcion correcta|correct answer|réponse correcte|reponse correcte)\s*:\s*([a-d](?:\s*,\s*[a-d])*)$", line, re.I)
        if marker:
            current.correct_option_letters = [letter.lower() for letter in re.findall(r"[a-d]", marker.group(1).lower())]
            continue
        if not current.options:
            current.prompt = _clean(f"{current.prompt} {line}")
    if current:
        questions.append(current)
    for question in questions:
        correct = set(question.correct_option_letters)
        for option in question.options:
            option.correct = option.letter in correct
    return questions


def _parse_matching(lines: list[str]) -> list[Question]:
    questions: list[Question] = []
    index = 0
    while index < len(lines):
        match = _NUMBERED.match(lines[index])
        if not match:
            index += 1
            continue
        prompt = _clean(match.group(2))
        index += 1
        while index < len(lines) and not re.match(r"^(?:columna|column|colonne)\s+a$", lines[index], re.I):
            if _NUMBERED.match(lines[index]):
                break
            index += 1
        if index >= len(lines) or not re.match(r"^(?:columna|column|colonne)\s+a$", lines[index], re.I):
            continue
        index += 1
        left: list[str] = []
        while index < len(lines) and not re.match(r"^(?:columna|column|colonne)\s+b$", lines[index], re.I):
            if re.match(r"^\[(?:R|A)\s*:", lines[index], re.I):
                break
            left.append(lines[index])
            index += 1
        if index >= len(lines) or not re.match(r"^(?:columna|column|colonne)\s+b$", lines[index], re.I):
            continue
        index += 1
        right: list[str] = []
        answer_text = ""
        while index < len(lines) and not _NUMBERED.match(lines[index]):
            marker = re.match(r"^\[(?:R|A)\s*:\s*(.+)\]$", lines[index], re.I)
            if marker:
                answer_text = marker.group(1)
                index += 1
                break
            right.append(lines[index])
            index += 1
        left_map = {}
        for i, value in enumerate(left):
            if not value:
                continue
            num_match = _NUMBERED.match(value)
            if num_match:
                left_map[num_match.group(1)] = _clean(num_match.group(2))
            else:
                left_map[str(i + 1)] = _clean(value)
        right_map = {}
        for i, value in enumerate(right):
            if not value:
                continue
            option = re.match(r"^([a-z])\s*[).]\s+(.+)$", value, re.I)
            if option:
                right_map[option.group(1).lower()] = _clean(option.group(2))
            else:
                right_map[chr(ord('a') + i)] = _clean(value)
        pairs = []
        for left_key, right_key in re.findall(r"(\d+)\s*[-–—]\s*([a-z])", answer_text, re.I):
            if left_map.get(left_key) and right_map.get(right_key.lower()):
                pairs.append(Pair(left_map[left_key], right_map[right_key.lower()]))
        questions.append(Question(prompt=prompt, pairs=pairs))
    return questions


def _parse_open(lines: list[str]) -> list[Question]:
    return [Question(prompt=_clean(match.group(2))) for line in lines if (match := _NUMBERED.match(line))][:1]


def _parse_block(block: str, index: int) -> DocumentTest:
    lines = _lines(block)
    sections = []
    labels = {
        "trueOrFalse": "Verdadero o falso",
        "multipleChoice": "Selección múltiple",
        "matching": "Emparejamiento",
        "writeAnswer": "Respuesta Corta",
    }
    parsers = {
        "trueOrFalse": _parse_true_false,
        "multipleChoice": _parse_multiple_choice,
        "matching": _parse_matching,
        "writeAnswer": _parse_open,
    }
    for code, parser in parsers.items():
        questions = parser(_section_lines(lines, code))
        if questions:
            sections.append(Section(code=code, name=labels[code], questions=questions))
    week_match = _WEEK.search(block)
    cut_match = _CUT.search(block)
    return DocumentTest(
        title=_extract_title(lines, index),
        description=_extract_description(lines),
        week=week_match.group(1) if week_match else "",
        cut=cut_match.group(1) if cut_match else "",
        language=_language(block),
        sections=sections,
    )


def parse_document_text(text: str, source_name: str = "") -> TestDocument:
    document = TestDocument(source_name=source_name)
    blocks = _split_blocks(text)
    tests = [_parse_block(block, index) for index, block in enumerate(blocks, start=1) if _lines(block)]
    tests_con_preguntas = [t for t in tests if t.sections]
    document.tests = tests_con_preguntas if tests_con_preguntas else tests
    return document


def _extraer_codigo_programa(texto: str) -> str:
    """Extrae el texto del programa solo si contiene un código unívoco (ej. COD111, P104, 104 - Nombre).
    Si solo contiene nombre genérico sin código (ej. 'Negocios Internacionales'), retorna cadena vacía.
    """
    if not texto:
        return ""
    texto_limpio = texto.strip()
    match_codigo = re.search(r"\b([A-Z]{2,}\s*[\-_]?\s*\d+|\d{3,})\b", texto_limpio, re.I)
    if match_codigo:
        return texto_limpio
    return ""


def parse_docx_bytes(content: bytes, source_name: str = "") -> TestDocument:
    source = DocxDocument(BytesIO(content))
    parts = []
    programa_detectado = ""
    try:
        from docx.text.paragraph import Paragraph
        from docx.table import Table

        for el in source.element.body:
            if el.tag.endswith("p"):
                p = Paragraph(el, source)
                txt = p.text.strip()
                if txt:
                    parts.append(txt)
            elif el.tag.endswith("tbl"):
                tbl = Table(el, source)
                for row in tbl.rows:
                    raw_cells = [c.text.strip() for c in row.cells]
                    cells = []
                    for c in raw_cells:
                        if not cells or c != cells[-1]:
                            cells.append(c)
                    if len(cells) == 2 and "\n" not in cells[0] and len(cells[0]) < 50:
                        header_lower = cells[0].lower().strip()
                        if re.match(r"^programa\b", header_lower) and not programa_detectado:
                            programa_detectado = _extraer_codigo_programa(cells[1])
                        parts.append(f"{cells[0]}: {cells[1]}")
                    else:
                        parts.append("\n".join(cells))
    except Exception:
        parts = [paragraph.text for paragraph in source.paragraphs if paragraph.text.strip()]
        for table in source.tables:
            for row in table.rows:
                parts.append("\t".join(cell.text for cell in row.cells))

    if not parts:
        parts = [paragraph.text for paragraph in source.paragraphs if paragraph.text.strip()]

    document = parse_document_text("\n".join(parts), source_name=source_name)
    if programa_detectado:
        document.program = programa_detectado
    if source.inline_shapes:
        document.issues.append(
            ValidationIssue(
                Severity.WARNING,
                f"El documento contiene {len(source.inline_shapes)} imagen(es) incrustada(s); verifica que no sean parte esencial de una pregunta.",
                "document.images",
            )
        )
    return document
