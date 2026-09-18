from __future__ import annotations

from html import escape

from .models import DocumentTest, Question


ACTIVITY_IDS = {
    "trueOrFalse": "122d0f18-7aa9-44e1-9dc1-2c0851a10940",
    "multipleChoice": "b78ebefa-222d-45ef-8436-99c6fee343f4",
    "matching": "54b45255-6fdc-43fe-8901-8b48e12162ab",
    "writeAnswer": "2f48b2b0-54a0-4bd9-af4e-84d389836974",
}


def _html_paragraph(value: str) -> str:
    return f"<p>{escape(value or '', quote=False)}</p>"


def _matching_activity(question: Question, ordering: int, name: str, language: str) -> dict:
    code = "matching"
    default_desc = {
        "en": "Match each concept with its corresponding description.",
        "fr": "Associez chaque concept à sa description correspondante.",
        "es": "Relacione cada concepto con su descripción correspondiente.",
    }.get(language, "Relacione cada concepto con su descripción correspondiente.")
    description = question.prompt or default_desc
    contents = [
        {
            "ordering": pair_index + 3,
            "content": pair.left,
            "content_options": [
                {
                    "ordering": pair_index + 3,
                    "text": pair.right,
                    "right_answer": True,
                }
            ],
        }
        for pair_index, pair in enumerate(question.pairs)
    ]
    return {
        "ordering": ordering,
        "activity_type": {"id": ACTIVITY_IDS[code], "code": code},
        "name_section": name,
        "description_section": _html_paragraph(description),
        "contents": contents,
    }


def _activity(section, ordering: int, language: str) -> dict:
    code = section.code
    names = {
        "trueOrFalse": {"es": "Verdadero o falso", "en": "True or False", "fr": "Vrai ou Faux"},
        "multipleChoice": {"es": "Selección múltiple", "en": "Multiple Choice", "fr": "Choix multiple"},
        "matching": {"es": "Emparejamiento", "en": "Matching", "fr": "Association"},
        "writeAnswer": {"es": "Respuesta Corta", "en": "Short Answer", "fr": "Réponse courte"},
    }
    name = names[code].get(language, names[code]["es"])
    if code == "matching":
        default_desc = {
            "en": "Match each concept with its corresponding description.",
            "fr": "Associez chaque concept à sa description correspondante.",
            "es": "Relacione cada concepto con su descripción correspondiente.",
        }.get(language, "Relacione cada concepto con su descripción correspondiente.")
        description = section.description or (section.questions[0].prompt if section.questions else default_desc)
        all_pairs = []
        for q in section.questions:
            all_pairs.extend(q.pairs)
        contents = [
            {
                "ordering": pair_index + 3,
                "content": pair.left,
                "content_options": [
                    {
                        "ordering": pair_index + 3,
                        "text": pair.right,
                        "right_answer": True,
                    }
                ],
            }
            for pair_index, pair in enumerate(all_pairs)
        ]
        return {
            "ordering": ordering,
            "activity_type": {"id": ACTIVITY_IDS[code], "code": code},
            "name_section": name,
            "description_section": _html_paragraph(description),
            "contents": contents,
        }

    contents = []
    for index, question in enumerate(section.questions):
        options = []
        if code == "trueOrFalse":
            true_text, false_text = ({"es": ("Verdadero", "Falso"), "en": ("True", "False"), "fr": ("Vrai", "Faux")}.get(language, ("Verdadero", "Falso")))
            options = [
                {"ordering": 1, "text": true_text, "right_answer": question.answer is True},
                {"ordering": 2, "text": false_text, "right_answer": question.answer is False},
            ]
        elif code == "multipleChoice":
            options = [
                {"ordering": option_index + 1, "text": option.text, "right_answer": option.letter in question.correct_option_letters}
                for option_index, option in enumerate(question.options)
            ]
        elif code == "writeAnswer":
            options = []

        contents.append(
            {
                "ordering": index + 2,
                "content": _html_paragraph(question.prompt),
                "content_options": options,
            }
        )
    return {
        "ordering": ordering,
        "activity_type": {"id": ACTIVITY_IDS[code], "code": code},
        "name_section": name,
        "description_section": name,
        "contents": contents,
    }


def serialize_test(test: DocumentTest) -> dict:
    activities = []
    for section in test.sections:
        if section.code == "matching":
            total_mq = len(section.questions)
            for q_idx, q in enumerate(section.questions, start=1):
                sec_name = f"Emparejamiento {q_idx}" if total_mq > 1 else section.name
                activities.append(
                    _matching_activity(
                        question=q,
                        ordering=len(activities) + 1,
                        name=sec_name,
                        language=test.language,
                    )
                )
        else:
            activities.append(
                _activity(
                    section=section,
                    ordering=len(activities) + 1,
                    language=test.language,
                )
            )

    return {
        "test_information": {
            "title": test.title,
            "description": _html_paragraph(test.description),
            "activities": activities,
        }
    }
