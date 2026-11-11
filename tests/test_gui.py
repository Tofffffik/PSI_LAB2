"""UI-тесты Ticketbox.

Настоящее окно открывается за пределами экрана на копии примеров данных:
ИС-21 — 5 студентов, ИС-22 — 4, ПИ-23 — 3, ПИ-24 — пустая; 10 билетов.
Тест выбирает пункты в списках, нажимает кнопки и смотрит, что видно.
Диалоги подменены: настоящий messagebox модальный.
"""

import random
import re
import shutil
import time
import tkinter as tk
from dataclasses import dataclass

import customtkinter as ctk
import pytest
from openpyxl import load_workbook

from ticketbox.core.workspace import BUNDLED, Workspace
from ticketbox.gui import look
from ticketbox.gui.main_window import NO_STUDENTS, PICK_GROUP, PICK_STUDENT, TicketBoxWindow
from ticketbox.gui.ticket_dialog import TicketDialog


@dataclass
class Seen:
    heading: str
    who: str
    repeated: bool
    questions: list[str]
    closed_by_escape: bool


class Screen:
    def __init__(self, host, workspace: Workspace) -> None:
        self.host = host
        self.workspace = workspace
        self.messages: list[tuple[str, str]] = []
        self.answers: list[bool] = []
        self.on_message = None
        self.window = TicketBoxWindow(host, workspace, chance=random.Random(3), notify=self._notify)
        self.wait(lambda: self.window.roster.group_names, "загрузка данных")

    def _notify(self, kind, title, text):
        self.messages.append((kind, text))
        if self.on_message:
            self.on_message()
        return self.answers.pop(0) if self.answers else True

    def wait(self, condition, what, timeout=10.0):
        deadline = time.monotonic() + timeout
        while not condition():
            if time.monotonic() > deadline:
                raise TimeoutError(f"Не дождались: {what}")
            self.host.update()
            time.sleep(0.01)

    @staticmethod
    def pick(menu: ctk.CTkOptionMenu, value: str) -> None:
        # Тот же путь, что при выборе пункта мышью: значение + обработчик.
        assert value in menu.cget("values"), f"нет пункта «{value}»"
        menu._dropdown_callback(value)

    def ready(self) -> bool:
        return self.window.issue_button.cget("state") == "normal"

    def press_issue(self) -> Seen | None:
        """Нажимает «Выдать билет». Окно билета модальное — его читаем и
        закрываем клавишей Esc из отложенного вызова. None — окна не было."""
        seen: list[Seen | None] = [None]

        def look_at_dialog():
            dialogs = [child for child in self.host.winfo_children() if isinstance(child, TicketDialog)]
            if not dialogs:
                return
            dialog = dialogs[0]
            dialog.update()
            result = Seen(dialog.heading.cget("text"), dialog.who.cget("text"), dialog.repeat_chip is not None,
                          list(dialog.questions), closed_by_escape=False)
            dialog.focus_force()
            dialog.update()
            dialog.event_generate("<Escape>")
            dialog.update()
            result.closed_by_escape = not dialog.winfo_exists()
            if not result.closed_by_escape:
                dialog.destroy()
            seen[0] = result

        self.host.after(300, look_at_dialog)
        assert self.ready(), "кнопка «Выдать билет» недоступна"
        self.window.issue_button.invoke()
        self.host.update()
        return seen[0]

    def journal(self):
        book = load_workbook(self.workspace.results)
        try:
            return [["" if v is None else str(v) for v in row]
                    for row in book.worksheets[0].iter_rows(min_row=2, values_only=True)]
        finally:
            book.close()


@pytest.fixture(scope="session")
def app():
    look.setup()
    root = ctk.CTk()
    root.withdraw()
    yield root
    root.destroy()


@pytest.fixture
def host(app):
    window = tk.Toplevel(app)
    window.geometry("+-20000+-20000")
    yield window
    window.destroy()


@pytest.fixture
def screen(tmp_path, host):
    for name in ("students.xlsx", "tickets.docx"):
        shutil.copy(BUNDLED / name, tmp_path / name)
    return Screen(host, Workspace(tmp_path))


def test_start_lists_groups_and_waits_for_choice(screen):
    window = screen.window

    assert list(window.group_menu.cget("values")) == ["ИС-21", "ИС-22", "ПИ-23", "ПИ-24"]
    assert window.group_menu.get() == PICK_GROUP and window.student_menu.get() == PICK_STUDENT
    assert not screen.ready()
    assert "билетов: 10" in window.status.cget("text")
    assert screen.messages == []


def test_group_choice_fills_students_and_button_needs_student(screen):
    screen.pick(screen.window.group_menu, "ИС-22")

    assert list(screen.window.student_menu.cget("values")) == [
        "Петров Иван", "Васильева Мария", "Морозов Никита", "Новикова Елена"]
    assert not screen.ready()

    screen.pick(screen.window.student_menu, "Новикова Елена")
    assert screen.ready()


def test_empty_group_disables_student_list_and_button(screen):
    screen.pick(screen.window.group_menu, "ПИ-24")

    assert screen.window.student_menu.get() == NO_STUDENTS
    assert screen.window.student_menu.cget("state") == "disabled"
    assert not screen.ready()


def test_issue_shows_ticket_writes_journal_and_escape_closes(screen):
    screen.pick(screen.window.group_menu, "ИС-21")
    screen.pick(screen.window.student_menu, "Иванов Пётр")

    seen = screen.press_issue()

    assert seen is not None
    assert re.fullmatch(r"Билет № ([1-9]|10)", seen.heading)
    assert seen.who == "Иванов Пётр · группа ИС-21"
    assert not seen.repeated
    assert len(seen.questions) == 3 and all(seen.questions)
    assert seen.closed_by_escape, "Esc не закрыл окно билета"

    assert screen.window.student_menu.get() == PICK_STUDENT and not screen.ready()
    [row] = screen.journal()
    assert row[:3] == ["ИС-21", "Иванов", "Пётр"]
    assert row[3] == seen.heading.split("№")[1].strip() and row[5] == "нет"


def test_same_student_twice_gets_same_ticket_marked_repeat(screen):
    screen.pick(screen.window.group_menu, "ПИ-23")
    screen.pick(screen.window.student_menu, "Алексеев Кирилл")
    first = screen.press_issue()
    screen.pick(screen.window.student_menu, "Алексеев Кирилл")
    second = screen.press_issue()

    assert first.heading == second.heading and first.questions == second.questions
    assert not first.repeated and second.repeated
    assert [row[5] for row in screen.journal()] == ["нет", "да"]


def test_busy_journal_cancel_writes_nothing(screen, opened_in_excel):
    screen.pick(screen.window.group_menu, "ИС-21")
    screen.pick(screen.window.student_menu, "Попов Артём")
    screen.answers = [False]

    with opened_in_excel(screen.workspace.results):
        seen = screen.press_issue()

    assert seen is None
    [(kind, text)] = screen.messages
    assert kind == "retry" and "открыт в Excel" in text
    assert screen.journal() == []


def test_busy_journal_retry_after_closing_file_succeeds(screen, opened_in_excel):
    screen.pick(screen.window.group_menu, "ИС-21")
    screen.pick(screen.window.student_menu, "Смирнова Ольга")
    lock = opened_in_excel(screen.workspace.results)
    lock.__enter__()
    screen.on_message = lambda: lock.__exit__(None, None, None)  # файл закрыли и нажали «Повторить»

    seen = screen.press_issue()

    assert [kind for kind, _ in screen.messages] == ["retry"]
    assert seen is not None and len(screen.journal()) == 1


def test_missing_files_are_reported_and_window_survives(tmp_path, host):
    messages = []
    window = TicketBoxWindow(host, Workspace(tmp_path), notify=lambda kind, title, text: messages.append(text) or True)
    deadline = time.monotonic() + 10
    while not messages and time.monotonic() < deadline:
        host.update()
        time.sleep(0.01)

    assert "Нет файла со списком студентов" in messages[0]
    assert "Нет файла с билетами" in messages[0]
    assert window.issue_button.cget("state") == "disabled"
