"""Список групп и студентов из students.xlsx.

Лист = группа (имя листа — название группы). Строка 1 — заголовок,
дальше: колонка A — фамилия, колонка B — имя.
"""

from dataclasses import dataclass, field
from pathlib import Path

from openpyxl import load_workbook


@dataclass(frozen=True, order=True)
class Person:
    surname: str
    name: str

    @property
    def label(self) -> str:
        """Как студент показывается в списке: «Фамилия Имя»."""
        return f"{self.surname} {self.name}".strip()


@dataclass
class Roster:
    groups: dict[str, list[Person]] = field(default_factory=dict)

    @property
    def group_names(self) -> list[str]:
        return list(self.groups)

    def people(self, group: str) -> list[Person]:
        """Студенты группы. Пустая или неизвестная группа — пустой список."""
        return list(self.groups.get(group, []))

    @classmethod
    def from_excel(cls, path: Path) -> "Roster":
        # read_only открывает файл только на чтение — работает, даже если
        # он сейчас открыт в Excel.
        book = load_workbook(path, read_only=True, data_only=True)
        try:
            roster = cls()
            for sheet in book.worksheets:
                people = []
                for row in sheet.iter_rows(min_row=2, max_col=2, values_only=True):
                    surname, name = (tuple(row) + (None, None))[:2]
                    person = Person(_text(surname), _text(name))
                    if person.label:
                        people.append(person)
                roster.groups[sheet.title] = people
            return roster
        finally:
            book.close()


def _text(cell) -> str:
    return "" if cell is None else str(cell).strip()
