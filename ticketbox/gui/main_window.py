"""Главное окно: выбор группы и студента, кнопка «Выдать билет»."""

import random
from tkinter import messagebox
from typing import Callable, Literal

import customtkinter as ctk

from ..core.question_bank import QuestionBank
from ..core.result_book import BookBusy, Issue, ResultBook
from ..core.roster import Person, Roster
from ..core.workspace import Workspace
from . import look
from .ticket_dialog import TicketDialog

PICK_GROUP = "— выберите группу —"
PICK_STUDENT = "— выберите студента —"
NO_STUDENTS = "в группе нет студентов"

Kind = Literal["warning", "error", "retry"]
# (вид, заголовок, текст) -> True для «ОК»/«Повторить». В тестах подменяется:
# настоящий messagebox модальный и остановил бы тест.
Notify = Callable[[Kind, str, str], bool]


def message_boxes(parent) -> Notify:
    def notify(kind: Kind, title: str, text: str) -> bool:
        if kind == "retry":
            return messagebox.askretrycancel(title, text, parent=parent)
        (messagebox.showerror if kind == "error" else messagebox.showwarning)(title, text, parent=parent)
        return True

    return notify


class TicketBoxWindow:
    def __init__(self, master, workspace: Workspace, chance: random.Random | None = None,
                 notify: Notify | None = None) -> None:
        self.master = master
        self.workspace = workspace
        self.chance = chance or random.Random()
        self.notify = notify or message_boxes(master)
        self.book = ResultBook(workspace.results)

        self.roster = Roster()
        self.bank = QuestionBank()
        self._people: dict[str, Person] = {}

        self._build()
        master.after(150, self.reload)

    # --- вид ---------------------------------------------------------------

    def _build(self) -> None:
        header = ctk.CTkFrame(self.master, fg_color=look.ACCENT, corner_radius=0)
        header.pack(fill="x")
        titles = ctk.CTkFrame(header, fg_color="transparent")
        titles.pack(side="left", padx=28, pady=(22, 20))
        ctk.CTkLabel(titles, text="Экзаменационные билеты", text_color=look.ON_ACCENT,
                     font=look.font(24, bold=True)).pack(anchor="w")
        ctk.CTkLabel(titles, text="Выдача билетов и учёт в журнале", text_color=look.ON_ACCENT_SOFT,
                     font=look.font(13)).pack(anchor="w")
        self.theme_switch = ctk.CTkSwitch(header, text="Тёмная тема", text_color=look.ON_ACCENT,
                                          progress_color=look.ACCENT_HOVER, command=self._toggle_theme)
        self.theme_switch.pack(side="right", padx=28)
        if ctk.get_appearance_mode() == "Dark":
            self.theme_switch.select()

        body = ctk.CTkFrame(self.master, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=28, pady=24)

        card = ctk.CTkFrame(body, corner_radius=16)
        card.pack(fill="x")
        card.grid_columnconfigure((0, 1), weight=1, uniform="column")
        for column, (caption, attribute, placeholder, command) in enumerate([
            ("Группа", "group_menu", PICK_GROUP, self.on_group_picked),
            ("Студент", "student_menu", PICK_STUDENT, lambda _: self._refresh_button()),
        ]):
            ctk.CTkLabel(card, text=caption, text_color=look.MUTED, font=look.font(13)).grid(
                row=0, column=column, sticky="w", padx=(20 if column == 0 else 10, 20 if column else 10), pady=(18, 4))
            menu = ctk.CTkOptionMenu(card, values=[placeholder], command=command, height=38, corner_radius=10,
                                     dynamic_resizing=False, font=look.font(14), **look.option_menu_colors())
            menu.set(placeholder)
            menu.grid(row=1, column=column, sticky="ew", padx=(20 if column == 0 else 10, 20 if column else 10))
            setattr(self, attribute, menu)

        self.issue_button = ctk.CTkButton(card, text="Выдать билет", height=48, corner_radius=12,
                                          fg_color=look.ACCENT, hover_color=look.ACCENT_HOVER,
                                          font=look.font(16, bold=True), command=self.issue, state="disabled")
        self.issue_button.grid(row=2, column=0, columnspan=2, sticky="ew", padx=20, pady=20)

        footer = ctk.CTkFrame(body, fg_color="transparent")
        footer.pack(fill="x", pady=(16, 0))
        self.status = ctk.CTkLabel(footer, text="", text_color=look.MUTED, font=look.font(12), justify="left")
        self.status.pack(side="left")
        self.reload_button = ctk.CTkButton(footer, text="Обновить данные", width=150, height=32, corner_radius=10,
                                           fg_color="transparent", border_width=1,
                                           text_color=("gray15", "gray90"), command=self.reload)
        self.reload_button.pack(side="right")

    def _toggle_theme(self) -> None:
        ctk.set_appearance_mode("dark" if self.theme_switch.get() else "light")

    # --- данные ------------------------------------------------------------

    def reload(self) -> None:
        """Перечитывает students.xlsx и tickets.docx, создаёт results.xlsx.
        Проблемы с файлами — сообщение, а не падение."""
        problems = []
        self.roster = self._read(self.workspace.students, Roster.from_excel, Roster(), problems,
                                 "Нет файла со списком студентов")
        self.bank = self._read(self.workspace.tickets, QuestionBank.from_docx, QuestionBank(), problems,
                               "Нет файла с билетами")
        for problem in self.bank.problems:
            self.workspace.log(f"tickets.docx: {problem}")
        if self.workspace.tickets.exists() and not self.bank.tickets and not problems:
            problems.append(f"В {self.workspace.tickets} нет ни одного правильного билета.")
        try:
            self.book.ensure()
        except OSError as error:
            problems.append(f"Не удалось создать {self.book.path}: {error}")

        groups = self.roster.group_names
        self.group_menu.configure(values=groups or [PICK_GROUP])
        self.group_menu.set(PICK_GROUP)
        self._set_students([])

        status = f"Групп: {len(groups)}  ·  билетов: {len(self.bank.tickets)}"
        if self.bank.problems:
            status += f"\nПропущено при разборе билетов: {len(self.bank.problems)} (подробно в app.log)"
        self.status.configure(text=status)

        if problems:
            for problem in problems:
                self.workspace.log(problem)
            self.notify("warning", "Не удалось прочитать данные",
                        "\n\n".join(problems) + "\n\nИсправьте файлы и нажмите «Обновить данные».")

    @staticmethod
    def _read(path, reader, empty, problems, missing_text):
        if not path.exists():
            problems.append(f"{missing_text}: {path}")
            return empty
        try:
            return reader(path)
        except Exception as error:  # повреждённый файл не должен ронять окно
            problems.append(f"Не удалось прочитать {path}: {error}")
            return empty

    # --- выбор ---------------------------------------------------------------

    @property
    def group(self) -> str | None:
        value = self.group_menu.get()
        return value if value in self.roster.groups else None

    @property
    def person(self) -> Person | None:
        return self._people.get(self.student_menu.get())

    def on_group_picked(self, group: str) -> None:
        self._set_students(self.roster.people(group))

    def _set_students(self, people: list[Person]) -> None:
        self._people = {person.label: person for person in people}
        if people:
            self.student_menu.configure(values=list(self._people), state="normal")
            self.student_menu.set(PICK_STUDENT)
        else:
            self.student_menu.configure(values=[PICK_STUDENT], state="disabled")
            self.student_menu.set(NO_STUDENTS if self.group else PICK_STUDENT)
        self._refresh_button()

    def _refresh_button(self) -> None:
        ready = self.group is not None and self.person is not None and bool(self.bank.tickets)
        self.issue_button.configure(state="normal" if ready else "disabled")

    # --- выдача --------------------------------------------------------------

    def issue(self) -> None:
        group, person = self.group, self.person
        if group is None or person is None:
            return

        issue = self._issue_with_retry(group, person)
        if issue is None:
            return

        TicketDialog(self.master, group, person, issue, self.bank.tickets.get(issue.number)).show()

        # Готовы к следующему студенту.
        self.student_menu.set(PICK_STUDENT)
        self._refresh_button()

    def _issue_with_retry(self, group: str, person: Person) -> Issue | None:
        while True:
            try:
                return self.book.issue(group, person, self.bank.numbers, self.chance)
            except BookBusy:
                again = self.notify(
                    "retry", "Журнал занят",
                    f"Файл {self.book.path} открыт в Excel или другой программе.\n\n"
                    "Закройте его и нажмите «Повторить».")
                if not again:
                    return None
            except Exception as error:
                self.workspace.log(f"Ошибка записи в журнал: {error!r}")
                self.notify("error", "Ошибка", f"Не удалось записать результат: {error}")
                return None
