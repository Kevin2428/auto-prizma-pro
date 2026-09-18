from __future__ import annotations

from dataclasses import dataclass, field
from enum import StrEnum


class Severity(StrEnum):
    ERROR = "error"
    WARNING = "warning"
    INFO = "info"


@dataclass
class ValidationIssue:
    severity: Severity
    message: str
    path: str = ""


@dataclass
class Option:
    letter: str
    text: str
    correct: bool = False


@dataclass
class Pair:
    left: str
    right: str


@dataclass
class Question:
    prompt: str
    options: list[Option] = field(default_factory=list)
    answer: bool | None = None
    correct_option_letters: list[str] = field(default_factory=list)
    pairs: list[Pair] = field(default_factory=list)


@dataclass
class Section:
    code: str
    name: str
    questions: list[Question] = field(default_factory=list)
    description: str = ""


@dataclass
class DocumentTest:
    title: str
    description: str = ""
    week: str = ""
    cut: str = ""
    language: str = "es"
    sections: list[Section] = field(default_factory=list)


@dataclass
class TestDocument:
    tests: list[DocumentTest] = field(default_factory=list)
    source_name: str = ""
    program: str = ""
    issues: list[ValidationIssue] = field(default_factory=list)


def document_to_dict(document: TestDocument) -> dict:
    return {
        "source_name": document.source_name,
        "program": document.program,
        "issues": [
            {"severity": issue.severity.value, "message": issue.message, "path": issue.path}
            for issue in document.issues
        ],
        "tests": [
            {
                "title": test.title,
                "description": test.description,
                "week": test.week,
                "cut": test.cut,
                "language": test.language,
                "sections": [
                    {
                        "code": section.code,
                        "name": section.name,
                        "description": section.description,
                        "questions": [
                            {
                                "prompt": question.prompt,
                                "answer": question.answer,
                                "correct_option_letters": question.correct_option_letters,
                                "options": [
                                    {"letter": option.letter, "text": option.text, "correct": option.correct}
                                    for option in question.options
                                ],
                                "pairs": [{"left": pair.left, "right": pair.right} for pair in question.pairs],
                            }
                            for question in section.questions
                        ],
                    }
                    for section in test.sections
                ],
            }
            for test in document.tests
        ],
    }


def document_from_dict(data: dict) -> TestDocument:
    tests = []
    for test_data in data.get("tests", []):
        sections = []
        for section_data in test_data.get("sections", []):
            questions = []
            for question_data in section_data.get("questions", []):
                questions.append(
                    Question(
                        prompt=str(question_data.get("prompt") or ""),
                        answer=question_data.get("answer"),
                        correct_option_letters=[str(letter).lower() for letter in question_data.get("correct_option_letters", [])],
                        options=[
                            Option(str(option.get("letter") or "").lower(), str(option.get("text") or ""), bool(option.get("correct")))
                            for option in question_data.get("options", [])
                        ],
                        pairs=[Pair(str(pair.get("left") or ""), str(pair.get("right") or "")) for pair in question_data.get("pairs", [])],
                    )
                )
            sections.append(
                Section(
                    code=str(section_data.get("code") or ""),
                    name=str(section_data.get("name") or ""),
                    description=str(section_data.get("description") or ""),
                    questions=questions,
                )
            )
        tests.append(
            DocumentTest(
                title=str(test_data.get("title") or ""),
                description=str(test_data.get("description") or ""),
                week=str(test_data.get("week") or ""),
                cut=str(test_data.get("cut") or ""),
                language=str(test_data.get("language") or "es"),
                sections=sections,
            )
        )
    return TestDocument(tests=tests, source_name=str(data.get("source_name") or ""), program=str(data.get("program") or ""))
