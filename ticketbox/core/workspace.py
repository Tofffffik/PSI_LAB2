"""Папка с данными: входные файлы, журнал результатов, лог."""

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

BUNDLED = Path(__file__).resolve().parents[2] / "data"


@dataclass(frozen=True)
class Workspace:
    folder: Path

    students = property(lambda self: self.folder / "students.xlsx")
    tickets = property(lambda self: self.folder / "tickets.docx")
    results = property(lambda self: self.folder / "results.xlsx")
    log_file = property(lambda self: self.folder / "app.log")

    @classmethod
    def locate(cls) -> "Workspace":
        """Текущая папка, если в ней есть students.xlsx или tickets.docx,
        иначе — примеры данных рядом с программой."""
        here = cls(Path.cwd())
        return here if here.students.exists() or here.tickets.exists() else cls(BUNDLED)

    def log(self, message: str) -> None:
        """Пишет строку в app.log; сбой записи лога ничего не ломает."""
        try:
            with self.log_file.open("a", encoding="utf-8") as file:
                file.write(f"[{datetime.now():%d.%m.%Y %H:%M:%S}] {message}\n")
        except OSError:
            pass
