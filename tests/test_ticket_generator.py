from contextlib import redirect_stdout
import io
import tempfile
import unittest
from datetime import datetime
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

from ticket_generator import (
    HEADERS,
    ExcelJournal,
    JournalError,
    generate_ticket_number,
    read_console_line,
    run,
    validate_name,
)


class NameValidationTests(unittest.TestCase):
    def test_trims_names(self) -> None:
        self.assertEqual(validate_name("  Иванов  "), "Иванов")

    def test_rejects_blank_name(self) -> None:
        with self.assertRaises(ValueError):
            validate_name(" \t ")


class TicketGenerationTests(unittest.TestCase):
    def test_ticket_number_is_inclusive_range(self) -> None:
        with patch("ticket_generator.random.randint", return_value=20) as randint:
            self.assertEqual(generate_ticket_number(), 20)
        randint.assert_called_once_with(1, 20)


class ConsoleInputTests(unittest.TestCase):
    def test_escape_cancels_input(self) -> None:
        keys = iter(["И", "\x1b"])
        with patch("builtins.print"):
            self.assertIsNone(read_console_line("Name: ", lambda: next(keys)))

    def test_backspace_edits_input(self) -> None:
        keys = iter(["A", "B", "\b", "C", "\r"])
        with patch("builtins.print"):
            self.assertEqual(read_console_line("Name: ", lambda: next(keys)), "AC")

    def test_end_of_input_cancels_input(self) -> None:
        with patch("builtins.print"):
            self.assertIsNone(read_console_line("Name: ", lambda: None))


class ExcelJournalTests(unittest.TestCase):
    def test_creates_header_and_appends_entries(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "journal.xlsx"
            journal = ExcelJournal(path)
            assigned_at = datetime(2026, 10, 21, 12, 30, 0)

            journal.append("Иванов", "Иван", 7, assigned_at)
            journal.append("Петров", "Пётр", 12, assigned_at)

            workbook = load_workbook(path, data_only=True)
            worksheet = workbook.active
            self.assertEqual(
                tuple(cell.value for cell in worksheet[1]),
                HEADERS,
            )
            self.assertEqual(
                tuple(cell.value for cell in worksheet[2]),
                ("Иванов", "Иван", 7, assigned_at),
            )
            self.assertEqual(
                tuple(cell.value for cell in worksheet[3]),
                ("Петров", "Пётр", 12, assigned_at),
            )
            workbook.close()

    def test_rejects_invalid_ticket_number(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            journal = ExcelJournal(Path(directory) / "journal.xlsx")
            with self.assertRaises(ValueError):
                journal.append("Иванов", "Иван", 21)
            self.assertFalse(journal.path.exists())

    def test_rejects_existing_workbook_with_wrong_header(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "journal.xlsx"
            workbook = Workbook()
            workbook.active.append(["Unexpected", "columns"])
            workbook.save(path)
            workbook.close()

            with self.assertRaises(JournalError):
                ExcelJournal(path).append("Иванов", "Иван", 7)

            workbook = load_workbook(path, data_only=True)
            self.assertEqual(workbook.active["A1"].value, "Unexpected")
            self.assertEqual(workbook.active.max_row, 1)
            workbook.close()

    def test_failed_replace_keeps_previous_journal_intact(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "journal.xlsx"
            journal = ExcelJournal(path)
            journal.append("Иванов", "Иван", 7)

            with patch("ticket_generator.os.replace", side_effect=PermissionError):
                with self.assertRaises(PermissionError):
                    journal.append("Петров", "Пётр", 12)

            workbook = load_workbook(path, data_only=True)
            self.assertEqual(workbook.active.max_row, 2)
            self.assertEqual(workbook.active["A2"].value, "Иванов")
            workbook.close()
            self.assertEqual(list(Path(directory).glob("*.tmp.xlsx")), [])


class ApplicationTests(unittest.TestCase):
    def test_save_retry_reuses_the_same_student_and_ticket(self) -> None:
        class FailingOnceJournal:
            def __init__(self) -> None:
                self.records: list[tuple[str, str, int]] = []

            def append(self, last_name: str, first_name: str, ticket: int) -> None:
                self.records.append((last_name, first_name, ticket))
                if len(self.records) == 1:
                    raise PermissionError("file is locked")

        journal = FailingOnceJournal()
        output = io.StringIO()
        with (
            patch(
                "ticket_generator._prompt_name",
                side_effect=["Иванов", "Иван", None],
            ),
            patch("ticket_generator.generate_ticket_number", return_value=7),
            patch("ticket_generator.read_console_line", return_value=""),
            redirect_stdout(output),
        ):
            self.assertEqual(run(journal), 0)

        self.assertEqual(
            journal.records,
            [("Иванов", "Иван", 7), ("Иванов", "Иван", 7)],
        )
        self.assertIn("Закройте journal.xlsx в Excel.", output.getvalue())


if __name__ == "__main__":
    unittest.main()
