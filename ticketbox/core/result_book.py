"""Книга результатов results.xlsx.

Строки только дописываются. Студент, уже получавший билет (та же группа,
фамилия и имя), получает номер из своей самой первой записи с пометкой
«да» в колонке «Повтор».
"""

import random
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Callable

from openpyxl import Workbook, load_workbook

from .roster import Person

HEADER = ("Группа", "Фамилия", "Имя", "Номер билета", "Дата и время", "Повтор (да/нет)")


class BookBusy(Exception):
    """results.xlsx занят другой программой (например, открыт в Excel)."""


@dataclass(frozen=True)
class Issue:
    number: int
    repeated: bool


class ResultBook:
    def __init__(self, path: Path, clock: Callable[[], datetime] = datetime.now) -> None:
        self.path = Path(path)
        self._clock = clock

    def ensure(self) -> None:
        if not self.path.exists():
            book = Workbook()
            book.active.title = "Results"
            book.active.append(HEADER)
            book.save(self.path)

    def issue(self, group: str, person: Person, numbers: list[int], chance: random.Random) -> Issue:
        """Выдаёт билет и сразу сохраняет запись."""
        self.ensure()
        try:
            book = load_workbook(self.path)
        except PermissionError as error:
            raise BookBusy(str(self.path)) from error
        try:
            sheet = book.worksheets[0]
            known = self._lookup(sheet, group, person)
            if known is None and not numbers:
                raise ValueError("Нет ни одного билета для выдачи.")
            issue = Issue(known, True) if known is not None else Issue(chance.choice(numbers), False)
            sheet.append([group, person.surname, person.name, issue.number,
                          f"{self._clock():%Y-%m-%d %H:%M:%S}", "да" if issue.repeated else "нет"])
            book.save(self.path)
            return issue
        except PermissionError as error:
            raise BookBusy(str(self.path)) from error
        finally:
            book.close()

    @staticmethod
    def _lookup(sheet, group: str, person: Person) -> int | None:
        wanted = (group, person.surname, person.name)
        for row in sheet.iter_rows(min_row=2, max_col=4, values_only=True):
            cells = tuple(str(value or "").strip().casefold() for value in row[:3])
            if cells == tuple(part.strip().casefold() for part in wanted) and isinstance(row[3], (int, float)):
                return int(row[3])
        return None
