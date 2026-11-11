"""Билеты из tickets.docx.

Ожидаемый вид документа:

    Билет 1
    1. Вопрос первый
    2. Вопрос второй
    3. Вопрос третий

Вопросы могут быть и элементами нумерованного списка Word. Билет, у
которого меньше трёх вопросов или нераспознан заголовок, пропускается;
причина попадает в ``problems``.
"""

import re
from dataclasses import dataclass, field
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

PER_TICKET = 3

TITLE = re.compile(r"Билет\s*№?\s*(\d+)\s*[.:]?", re.IGNORECASE)
LOOKS_LIKE_TITLE = re.compile(r"Билет(?![а-яё])", re.IGNORECASE)
NUMBERED = re.compile(r"\d+\s*[.)]\s*(?P<text>.+)")


@dataclass
class QuestionBank:
    tickets: dict[int, tuple[str, ...]] = field(default_factory=dict)
    problems: list[str] = field(default_factory=list)

    @property
    def numbers(self) -> list[int]:
        return sorted(self.tickets)

    @classmethod
    def from_docx(cls, path: Path) -> "QuestionBank":
        return cls.from_paragraphs(_paragraphs(Document(str(path))))

    @classmethod
    def from_paragraphs(cls, paragraphs) -> "QuestionBank":
        """``paragraphs`` — пары (текст, абзац_списка)."""
        collector = _Collector()
        for text, is_list_item in paragraphs:
            collector.feed(text.strip(), is_list_item)
        collector.close()
        return collector.bank


class _Collector:
    """Автомат разбора: текущий билет копит вопросы до следующего заголовка."""

    def __init__(self) -> None:
        self.bank = QuestionBank()
        self._number: int | None = None   # номер текущего билета
        self._broken = False              # заголовок текущего блока не распознан
        self._questions: list[str] = []
        self._inside = False

    def feed(self, text: str, is_list_item: bool) -> None:
        if not text:
            return
        title = TITLE.fullmatch(text)
        if title:
            self._finish()
            self._inside, self._broken, self._number = True, False, int(title.group(1))
        elif LOOKS_LIKE_TITLE.match(text):
            # Блок открываем, чтобы его вопросы не достались предыдущему билету.
            self._finish()
            self._inside, self._broken, self._number = True, True, None
            self.bank.problems.append(f"Заголовок «{text}» не распознан — билет пропущен.")
        elif self._inside and not self._broken:
            numbered = NUMBERED.fullmatch(text)
            if numbered:
                self._questions.append(numbered["text"].strip())
            elif is_list_item:
                self._questions.append(text)

    def close(self) -> None:
        self._finish()

    def _finish(self) -> None:
        number, questions = self._number, self._questions
        self._number, self._questions, self._inside = None, [], False
        if number is None:
            return
        if len(questions) < PER_TICKET:
            self.bank.problems.append(
                f"Билет {number}: вопросов {len(questions)} из {PER_TICKET} — билет пропущен.")
        elif number in self.bank.tickets:
            self.bank.problems.append(f"Билет {number} встречается повторно — второй пропущен.")
        else:
            if len(questions) > PER_TICKET:
                self.bank.problems.append(f"Билет {number}: вопросов больше {PER_TICKET}, взяты первые.")
            self.bank.tickets[number] = tuple(questions[:PER_TICKET])


def _paragraphs(document):
    """Абзацы документа (и из таблиц) как пары (текст, абзац_списка).
    Перенос строки внутри абзаца (Shift+Enter) делит его на строки."""
    list_styles = _list_styles(document)
    for paragraph in document.element.body.iter(qn("w:p")):
        props = paragraph.pPr
        style = props.pStyle.val if props is not None and props.pStyle is not None else None
        in_list = props is not None and (props.numPr is not None or style in list_styles)

        pieces = []
        for node in paragraph.iter(qn("w:t"), qn("w:br"), qn("w:cr"), qn("w:tab")):
            if node.tag == qn("w:t"):
                pieces.append(node.text or "")
            elif node.tag == qn("w:tab"):
                pieces.append(" ")
            else:
                pieces.append("\n")
        for line in "".join(pieces).split("\n"):
            yield line, in_list


def _list_styles(document) -> set[str]:
    """Стили абзацев с нумерацией — свои или унаследованные."""
    numbered, parents = set(), {}
    for style in document.styles.element.findall(qn("w:style")):
        style_id = style.get(qn("w:styleId"))
        if style.find(f"{qn('w:pPr')}/{qn('w:numPr')}") is not None:
            numbered.add(style_id)
        based_on = style.find(qn("w:basedOn"))
        if based_on is not None:
            parents[style_id] = based_on.get(qn("w:val"))

    for style_id in list(parents):
        chain, current = [], style_id
        while current in parents and current not in numbered and current not in chain:
            chain.append(current)
            current = parents[current]
        if current in numbered:
            numbered.update(chain)
    return numbered
