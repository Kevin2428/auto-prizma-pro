from __future__ import annotations

from .models import Severity, TestDocument, ValidationIssue


def validate_document(document: TestDocument) -> list[ValidationIssue]:
    issues = list(document.issues)
    if not document.tests:
        return issues + [ValidationIssue(Severity.ERROR, "No se detectaron tests en el documento.", "tests")]
    for test_index, test in enumerate(document.tests):
        prefix = f"tests[{test_index}]"
        test_label = f"Test {test_index + 1}" if len(document.tests) > 1 else "Test"
        if not test.title.strip():
            issues.append(ValidationIssue(Severity.ERROR, f"{test_label}: El test no tiene nombre o título.", f"{prefix}.title"))
        for section_index, section in enumerate(test.sections):
            section_prefix = f"{prefix}.sections[{section_index}]"
            if not section.questions:
                issues.append(ValidationIssue(Severity.ERROR, f"{test_label} ({section.name}): La sección no contiene preguntas.", section_prefix))
            for question_index, question in enumerate(section.questions):
                path = f"{section_prefix}.questions[{question_index}]"
                q_num = question_index + 1
                q_label = f"{test_label} • {section.name} • Pregunta {q_num}"
                if not question.prompt.strip():
                    issues.append(ValidationIssue(Severity.ERROR, f"{q_label}: La pregunta está vacía.", path))
                if section.code == "trueOrFalse" and question.answer is None:
                    issues.append(ValidationIssue(Severity.ERROR, f"{q_label}: La pregunta verdadero/falso no tiene respuesta seleccionada.", path))
                elif section.code == "multipleChoice":
                    letters = {option.letter for option in question.options}
                    if len(question.options) < 2:
                        issues.append(ValidationIssue(Severity.ERROR, f"{q_label}: La pregunta debe tener al menos dos opciones.", path))
                    if not question.correct_option_letters:
                        issues.append(ValidationIssue(Severity.ERROR, f"{q_label}: La pregunta de selección no tiene respuesta correcta marcada.", path))
                    elif not set(question.correct_option_letters).issubset(letters):
                        issues.append(ValidationIssue(Severity.ERROR, f"{q_label}: La clave apunta a una opción inexistente.", path))
                elif section.code == "matching" and not question.pairs:
                    issues.append(ValidationIssue(Severity.ERROR, f"{q_label}: El emparejamiento no tiene parejas válidas.", path))
    return issues
