"""Ticketbox — выдача экзаменационных билетов.

Группы и студенты — из students.xlsx, билеты — из tickets.docx, журнал —
results.xlsx. Файлы ищутся в текущей папке, иначе берутся примеры из
lab2/data.

Запуск:
    python lab2/main.py
"""

from tkinter import messagebox

import customtkinter as ctk

from ticketbox.core.workspace import Workspace
from ticketbox.gui import look
from ticketbox.gui.main_window import TicketBoxWindow


def main() -> None:
    look.setup()
    workspace = Workspace.locate()
    app = ctk.CTk()
    app.title("Ticketbox — экзаменационные билеты")
    app.geometry("620x440")
    app.minsize(560, 420)

    def on_error(exc_type, error, traceback) -> None:
        workspace.log(f"Необработанная ошибка: {error!r}")
        messagebox.showerror("Ошибка", str(error), parent=app)

    app.report_callback_exception = on_error
    TicketBoxWindow(app, workspace)
    app.mainloop()


if __name__ == "__main__":
    main()
