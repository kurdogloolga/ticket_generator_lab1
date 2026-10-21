"""Console application for generating exam tickets and recording them in Excel."""

from __future__ import annotations

import os
import random
import sys
import tempfile
from copy import copy
from datetime import datetime
from pathlib import Path
from typing import Callable

from openpyxl import Workbook, load_workbook

HEADERS = ("Last name", "First name", "Номер билета", "Дата и время")
MIN_TICKET_NUMBER = 1
MAX_TICKET_NUMBER = 20


class JournalError(Exception):
    """Raised when the journal file does not have the expected format."""


def validate_name(value: str) -> str:
    """Trim a name and reject values containing only whitespace."""
    name = value.strip()
    if not name:
        raise ValueError("Имя не может быть пустым.")
    return name


def generate_ticket_number() -> int:
    """Return a random ticket number from 1 through 20, inclusive."""
    return random.randint(MIN_TICKET_NUMBER, MAX_TICKET_NUMBER)


class ExcelJournal:
    """Append student ticket assignments to an Excel workbook."""

    def __init__(self, path: Path | str = "journal.xlsx") -> None:
        self.path = Path(path)

    def append(
        self,
        last_name: str,
        first_name: str,
        ticket_number: int,
        assigned_at: datetime | None = None,
    ) -> None:
        last_name = validate_name(last_name)
        first_name = validate_name(first_name)
        if not MIN_TICKET_NUMBER <= ticket_number <= MAX_TICKET_NUMBER:
            raise ValueError(
                f"Номер билета должен быть от {MIN_TICKET_NUMBER} "
                f"до {MAX_TICKET_NUMBER}."
            )

        if self.path.exists():
            workbook = load_workbook(self.path)
            worksheet = workbook.active
            current_headers = tuple(
                worksheet.cell(row=1, column=column).value
                for column in range(1, len(HEADERS) + 1)
            )
            if current_headers != HEADERS:
                workbook.close()
                raise JournalError(
                    "В файле журнала отсутствует ожидаемая шапка столбцов. "
                    "Существующий файл не изменён."
                )
        else:
            workbook = Workbook()
            worksheet = workbook.active
            worksheet.append(HEADERS)
            worksheet.freeze_panes = "A2"
            worksheet.auto_filter.ref = "A1:D1"
            for cell in worksheet[1]:
                font = copy(cell.font)
                font.bold = True
                cell.font = font

        worksheet.append(
            [
                last_name,
                first_name,
                ticket_number,
                assigned_at or datetime.now(),
            ]
        )
        worksheet.cell(row=worksheet.max_row, column=4).number_format = (
            "yyyy-mm-dd hh:mm:ss"
        )

        temporary_path: Path | None = None
        try:
            with tempfile.NamedTemporaryFile(
                prefix=f".{self.path.stem}-",
                suffix=".tmp.xlsx",
                dir=self.path.parent,
                delete=False,
            ) as temporary_file:
                temporary_path = Path(temporary_file.name)
            workbook.save(temporary_path)
            os.replace(temporary_path, self.path)
        finally:
            workbook.close()
            if temporary_path is not None and temporary_path.exists():
                temporary_path.unlink()


def _read_key() -> str | None:
    if os.name == "nt":
        import msvcrt

        character = msvcrt.getwch()
        if character in ("\x00", "\xe0"):
            msvcrt.getwch()
            return ""
        return character
    return sys.stdin.read(1) or None


def read_console_line(
    prompt: str,
    read_key: Callable[[], str | None] | None = None,
) -> str | None:
    """Read a line with basic editing; return None when Escape is pressed."""
    print(prompt, end="", flush=True)
    characters: list[str] = []
    key_reader = read_key or _read_key
    terminal_settings = None

    if os.name != "nt" and sys.stdin.isatty():
        import termios
        import tty

        terminal_settings = termios.tcgetattr(sys.stdin.fileno())
        tty.setraw(sys.stdin.fileno())

    try:
        while True:
            character = key_reader()
            if character is None:
                print()
                return None
            if character == "":
                continue
            if character == "\x1b":
                print()
                return None
            if character in ("\r", "\n"):
                print()
                return "".join(characters)
            if character in ("\b", "\x7f"):
                if characters:
                    characters.pop()
                    print("\b \b", end="", flush=True)
                continue
            if character.isprintable():
                characters.append(character)
                print(character, end="", flush=True)
    finally:
        if terminal_settings is not None:
            termios.tcsetattr(sys.stdin.fileno(), termios.TCSADRAIN, terminal_settings)


def _prompt_name(label: str) -> str | None:
    while True:
        value = read_console_line(f"{label}: ")
        if value is None:
            return None
        try:
            return validate_name(value)
        except ValueError as error:
            print(error)


def run(journal: ExcelJournal | None = None) -> int:
    """Run the interactive application. Return a process exit code."""
    journal = journal or ExcelJournal()
    print("Для выхода нажмите ESC.")

    while True:
        last_name = _prompt_name("Last name")
        if last_name is None:
            return 0
        first_name = _prompt_name("First name")
        if first_name is None:
            return 0

        ticket_number = generate_ticket_number()
        print(f"Билет № {ticket_number}.")

        while True:
            try:
                journal.append(last_name, first_name, ticket_number)
                print("Запись сохранена в journal.xlsx.")
                break
            except (PermissionError, OSError) as error:
                print(f"Не удалось сохранить запись: {error}")
                print("Закройте journal.xlsx в Excel.")
                if read_console_line(
                    "Нажмите Enter для повтора или ESC для выхода: "
                ) is None:
                    print("Запись не сохранена.")
                    return 1
            except JournalError as error:
                print(f"Ошибка формата журнала: {error}")
                return 1


if __name__ == "__main__":
    raise SystemExit(run())
