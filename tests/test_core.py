import random
from datetime import datetime

import pytest
from docx import Document
from openpyxl import Workbook, load_workbook

from ticketbox.core.question_bank import QuestionBank
from ticketbox.core.result_book import HEADER, BookBusy, ResultBook
from ticketbox.core.roster import Person, Roster
from ticketbox.core.workspace import BUNDLED

IVANOV = Person("Иванов", "Пётр")
CLOCK = datetime(2026, 11, 11, 10, 15, 0)


def paragraphs(*texts):
    return [(text, False) for text in texts]


# --- список студентов ----------------------------------------------------------

@pytest.fixture
def roster_file(tmp_path):
    book = Workbook()
    book.remove(book.active)
    for group, people in [("ИС-21", [("Иванов", "Пётр"), ("Сидорова", "Анна")]), ("ИС-22", [("Петров", "Иван")]),
                          ("Пустая", [])]:
        sheet = book.create_sheet(group)
        sheet.append(["Фамилия", "Имя"])
        for person in people:
            sheet.append(person)
    path = tmp_path / "students.xlsx"
    book.save(path)
    return path


def test_groups_are_sheet_names_in_order(roster_file):
    assert Roster.from_excel(roster_file).group_names == ["ИС-21", "ИС-22", "Пустая"]


def test_people_of_a_group_with_labels(roster_file):
    people = Roster.from_excel(roster_file).people("ИС-21")

    assert people == [Person("Иванов", "Пётр"), Person("Сидорова", "Анна")]
    assert [person.label for person in people] == ["Иванов Пётр", "Сидорова Анна"]


@pytest.mark.parametrize("group", ["Пустая", "Нет такой"])
def test_empty_or_unknown_group_gives_no_people(roster_file, group):
    assert Roster.from_excel(roster_file).people(group) == []


# --- билеты --------------------------------------------------------------------

def test_tickets_from_word_document(tmp_path):
    document = Document()
    for text in ["Билет 1", "1. Первый", "2. Второй", "3. Третий", "Билет 2", "1) А", "2) Б", "3) В"]:
        document.add_paragraph(text)
    document.save(tmp_path / "tickets.docx")

    bank = QuestionBank.from_docx(tmp_path / "tickets.docx")

    assert bank.tickets == {1: ("Первый", "Второй", "Третий"), 2: ("А", "Б", "В")}
    assert bank.problems == []


def test_word_numbered_list_style_counts_as_questions(tmp_path):
    document = Document()
    document.add_paragraph("Билет 9")
    for text in ["Раз", "Два", "Три"]:
        document.add_paragraph(text, style="List Number")
    document.save(tmp_path / "tickets.docx")

    assert QuestionBank.from_docx(tmp_path / "tickets.docx").tickets == {9: ("Раз", "Два", "Три")}


def test_bundled_sample_has_ten_full_tickets():
    bank = QuestionBank.from_docx(BUNDLED / "tickets.docx")

    assert bank.numbers == list(range(1, 11))
    assert all(len(questions) == 3 for questions in bank.tickets.values())


def test_short_ticket_is_skipped_with_reason():
    bank = QuestionBank.from_paragraphs(paragraphs("Билет 1", "1. А", "2. Б", "Билет 2", "1. В", "2. Г", "3. Д"))

    assert bank.numbers == [2]
    assert "Билет 1" in bank.problems[0]


def test_unrecognised_title_skips_ticket_without_touching_previous():
    bank = QuestionBank.from_paragraphs(paragraphs(
        "Билет 1", "1. А", "2. Б", "3. В",
        "Билет второй", "1. Г", "2. Д", "3. Е",
        "Билет 3", "1. Ж", "2. З", "3. И"))

    assert bank.numbers == [1, 3]
    assert bank.tickets[1] == ("А", "Б", "В")
    assert len(bank.problems) == 1


def test_duplicate_number_keeps_first():
    bank = QuestionBank.from_paragraphs(paragraphs("Билет 1", "1. А", "2. Б", "3. В", "Билет 1", "1. Г", "2. Д",
                                                   "3. Е"))

    assert bank.tickets == {1: ("А", "Б", "В")}
    assert len(bank.problems) == 1


def test_lines_before_first_ticket_are_ignored():
    bank = QuestionBank.from_paragraphs(paragraphs("Экзаменационные билеты", "1. Не вопрос", "Билет 1",
                                                   "1. А", "2. Б", "3. В"))

    assert bank.tickets == {1: ("А", "Б", "В")} and bank.problems == []


# --- журнал результатов ----------------------------------------------------------

def rows(path):
    book = load_workbook(path)
    try:
        return [["" if value is None else str(value) for value in row]
                for row in book.worksheets[0].iter_rows(values_only=True)]
    finally:
        book.close()


@pytest.fixture
def book(tmp_path):
    return ResultBook(tmp_path / "results.xlsx", clock=lambda: CLOCK)


def test_ensure_creates_header(book):
    book.ensure()

    assert rows(book.path) == [list(HEADER)]


def test_first_issue_is_random_from_bank_and_not_repeat(book):
    issue = book.issue("ИС-21", IVANOV, [1, 2, 3], random.Random(1))

    assert issue.number in (1, 2, 3) and not issue.repeated
    assert rows(book.path)[1] == ["ИС-21", "Иванов", "Пётр", str(issue.number), "2026-11-11 10:15:00", "нет"]


def test_repeat_gets_first_number_and_new_row(book):
    chance = random.Random(9)
    first = book.issue("ИС-21", IVANOV, list(range(1, 21)), chance)
    book.issue("ИС-21", Person("Петров", "Иван"), list(range(1, 21)), chance)
    before = rows(book.path)

    again = book.issue("ИС-21", IVANOV, list(range(1, 21)), chance)

    after = rows(book.path)
    assert again.repeated and again.number == first.number
    assert after[:len(before)] == before
    assert after[-1][3] == str(first.number) and after[-1][5] == "да"


def test_same_name_in_another_group_is_another_student(book):
    book.issue("ИС-21", IVANOV, [5], random.Random(1))

    assert not book.issue("ИС-22", IVANOV, [5], random.Random(1)).repeated


def test_busy_book_raises_and_keeps_data(book, opened_in_excel):
    book.ensure()

    with opened_in_excel(book.path):
        with pytest.raises(BookBusy):
            book.issue("ИС-21", IVANOV, [1], random.Random(1))

    book.issue("ИС-21", IVANOV, [1], random.Random(1))
    assert len(rows(book.path)) == 2
