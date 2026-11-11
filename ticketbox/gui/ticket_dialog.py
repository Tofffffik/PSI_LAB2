"""Окно выданного билета: большой номер, студент, три вопроса.

Esc или Enter закрывают окно; главное окно остаётся открытым.
"""

import customtkinter as ctk

from ..core.result_book import Issue
from ..core.roster import Person
from . import look

TEXT_WIDTH = 440


class TicketDialog(ctk.CTkToplevel):
    def __init__(self, master, group: str, person: Person, issue: Issue,
                 questions: tuple[str, ...] | None) -> None:
        super().__init__(master)
        self.title("Билет")
        self.resizable(False, False)
        self.transient(master)

        frame = ctk.CTkFrame(self, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=28, pady=24)

        top = ctk.CTkFrame(frame, fg_color="transparent")
        top.pack(fill="x")
        # Круг — рамка фиксированного размера; подпись внутри по центру.
        circle = ctk.CTkFrame(top, width=88, height=88, corner_radius=44, fg_color=look.ACCENT)
        circle.pack(side="left")
        circle.pack_propagate(False)
        self.badge = ctk.CTkLabel(circle, text=str(issue.number), text_color=look.ON_ACCENT,
                                  font=look.font(36, bold=True))
        self.badge.place(relx=0.5, rely=0.5, anchor="center")
        caption = ctk.CTkFrame(top, fg_color="transparent")
        caption.pack(side="left", padx=(18, 0))
        self.heading = ctk.CTkLabel(caption, text=f"Билет № {issue.number}", font=look.font(22, bold=True))
        self.heading.pack(anchor="w")
        self.who = ctk.CTkLabel(caption, text=f"{person.label} · группа {group}", text_color=look.MUTED,
                                font=look.font(13))
        self.who.pack(anchor="w")

        self.repeat_chip: ctk.CTkLabel | None = None
        if issue.repeated:
            self.repeat_chip = ctk.CTkLabel(frame, text="Повторная выдача: номер взят из первой записи",
                                            fg_color=look.NOTE_BG, text_color=look.NOTE_TEXT, corner_radius=8,
                                            font=look.font(13))
            self.repeat_chip.pack(fill="x", pady=(16, 0), ipady=6)

        self.questions = list(questions or ())
        if questions is None:
            ctk.CTkLabel(frame, text=f"Вопросы билета № {issue.number} не найдены в tickets.docx.",
                         text_color=look.DANGER, font=look.font(13)).pack(anchor="w", pady=(16, 0))
        for index, question in enumerate(self.questions, start=1):
            box = ctk.CTkFrame(frame, corner_radius=12)
            box.pack(fill="x", pady=(14 if index == 1 else 8, 0))
            ctk.CTkLabel(box, text=f"Вопрос {index}", text_color=look.ACCENT,
                         font=look.font(12, bold=True)).pack(anchor="w", padx=16, pady=(10, 0))
            ctk.CTkLabel(box, text=question, wraplength=TEXT_WIDTH, justify="left",
                         font=look.font(14)).pack(anchor="w", padx=16, pady=(0, 12))

        self.close_button = ctk.CTkButton(frame, text="Закрыть  ·  Esc", width=160, height=38, corner_radius=10,
                                          fg_color="transparent", border_width=1, text_color=("gray15", "gray90"),
                                          command=self.destroy)
        self.close_button.pack(anchor="e", pady=(20, 0))

        self.bind("<Escape>", lambda _: self.destroy())
        self.bind("<Return>", lambda _: self.destroy())

    def show(self) -> None:
        """Показывает окно над главным и ждёт, пока его закроют."""
        owner = self.master
        self.update_idletasks()
        x = owner.winfo_rootx() + (owner.winfo_width() - self.winfo_reqwidth()) // 2
        y = owner.winfo_rooty() + (owner.winfo_height() - self.winfo_reqheight()) // 2
        self.geometry(f"+{x}+{y}")
        self.grab_set()
        self.focus_force()
        owner.wait_window(self)
