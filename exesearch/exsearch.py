"""
This searches through worksheets by Excel to find search terms.

:return: The cells, sheets and books a term was found in
:rtype: print()
"""

import datetime
from openpyxl import load_workbook, Workbook
from pathlib import Path
from tkinter import filedialog, Tk
from io import BytesIO
from msoffcrypto import OfficeFile
from collections import defaultdict
from warnings import filterwarnings


class Search:
    """
    This class holds properties related to a search.
    """

    def __init__(self) -> None:
        self.timestamp = datetime.datetime.now()

    # This defines type-hints for later variables to be added.
    term: str
    path: Path
    is_exclusive: bool
    is_case_sensitive: bool
    found: defaultdict


def ask_for_search() -> Search:
    """
    This function requests a path from the user to search through. This then exclusively or inclusively searches for a term.

    :return: The search information, along with a timestamp.
    :rtype: Search
    """
    new_search = Search()

    new_search.path = get_search_directory()
    new_search.term = input("What term are you looking for? ")

    new_search.is_exclusive = get_exclusivity()
    new_search.is_case_sensitive = True

    if not new_search.is_exclusive:
        # This is asked only when an inclusive search is used
        new_search.is_case_sensitive = get_case_sensitivity()

    return new_search


def find_workbooks(search: Search):
    """
    This function searches through each file in a path and if valid, searches that workbook for the search term.

    :param search: The search to perform.
    :type search: Search
    """
    for file in search.path.rglob("*.xlsm"):
        if file.is_file and file.suffix == ".xlsm" and "$" not in file.name:
            file_path = Path(search.path, file)
            search.found = search_for_term_in_workbook(
                file_path, search.term, search.is_exclusive, search.is_case_sensitive
            )


def get_case_sensitivity() -> bool:
    """
    This function asks the user if the search should be case sensitive.

    :return: If the search is case sensitive.
    :rtype: bool
    """
    return input("Is the term case sensitive? (y/n): ").lower().strip() == "y"


def get_exclusivity() -> bool:
    """
    This function asks the user if the search should be exclusive or not.

    :return: If the search should be exclusive
    :rtype: bool
    """
    return (
        input(
            "Would you like to exclusively search, finding only exactly matching cells? (y/n): "
        )
        .lower()
        .strip()
        == "y"
    )  # This adds the count of terms found to the count, after searching a workbook


def get_search_directory():
    """
    This function returns a searching directory chosen by the user.

    :return: The directory to search
    :rtype: Path
    """
    print("Please choose a directory.")

    root = Tk()
    root.withdraw()

    path = Path(filedialog.askdirectory(title="Please choose a directory"))

    return path


def prepare_workbook(book_path: Path) -> BytesIO:
    """
    This function prepares a workbook to be edited, by unlocking any encrypted files.

    :param book_path: The path of the workbook to prepare
    :type book_path: Path
    :return: The bytes stream of the decrypted workbook
    :rtype: BytesIO
    """
    with open(book_path, "rb") as workbook:
        office_file = OfficeFile(workbook)
        workbook.seek(0)  # This resets the read position for the file

        unencrypted_type = "plain"
        if (
            office_file.is_encrypted and office_file.type != unencrypted_type
        ):  # This checks if the file is not plain and encrypted
            default_key = input("What is the password to this file?")
            decrypted_workbook = (
                BytesIO()
            )  # This creates an in-memory BytesIO object to write the file to

            office_file.load_key(password=default_key)
            office_file.decrypt(decrypted_workbook)

            return decrypted_workbook

        else:
            return BytesIO(workbook.read())  # This return the in-memory file decrypted


def search_for_term_in_workbook(
    book_path: Path, term: str, exclusive: bool, case_sensitive: bool
) -> defaultdict:
    """
    This function finds terms inside a book, in all sheets.

    :param book: The path to the book
    :type book: Path
    :return: The count of how many times a term was found
    :rtype: int
    """
    workbook_byte_stream = prepare_workbook(book_path)
    workbook: Workbook = load_workbook(
        filename=workbook_byte_stream, data_only=True, read_only=True
    )

    print_workbook(book_path)  # This prints the name and path of the workbook

    count: int = 0
    found_cells = defaultdict(
        lambda: ""
    )  # This creates a dictionary to sort found cells by sheet

    for sheet in workbook:
        for column in sheet.iter_rows(max_col=11, max_row=99):
            for cell in column:
                if exclusive:
                    if cell.value == term:
                        count += 1
                        found_cells[sheet.title] += f"{cell.column_letter}{cell.row}, "
                else:
                    if case_sensitive:
                        if term in str(cell.value):
                            count += 1
                            found_cells[sheet.title] += (
                                f"{cell.column_letter}{cell.row}, "
                            )
                    else:
                        if term.lower().strip() in str(cell.value).lower().strip():
                            count += 1
                            found_cells[sheet.title] += (
                                f"{cell.column_letter}{cell.row}, "
                            )

    print_found_cells(found_cells)

    return found_cells


def print_workbook(book_path: Path) -> None:
    """
    This function prints the workbook name and path.

    :param book_path: The path to the workbook
    :type book_path: Path
    """
    print(f"\n_______________ Book: {book_path.name} _______________\n")
    print(f"Path: {book_path.absolute()}\n")


def print_found_cells(sheet_cells: defaultdict[str, str]) -> None:
    """
    This function prints the cells found, grouped by sheet.

    :param sheet_cells: The cells found, by sheet
    :type sheet_cells: defaultdict[str, str]
    """
    print("Cells Found:")

    if not len(sheet_cells):
        print("None")  # This prints if no cells were found in that book
        return

    for sheet in sheet_cells:
        print(
            f"Sheet '{sheet}': {sheet_cells[sheet]}"
        )  # This prints each sheet's found cells


def main():
    filterwarnings("ignore", category=UserWarning, module="openpyxl")

    search = ask_for_search()

    find_workbooks(search)

    input("\nPress enter to exit.")


if __name__ == "__main__":
    main()