"""
This searches through worksheets by Excel to find seach terms.

:return: The cells, sheets and books a term was found in
:rtype: print()
"""

from openpyxl import load_workbook, Workbook
from pathlib import Path
from tkinter import filedialog, Tk
from io import BytesIO
from msoffcrypto import OfficeFile
from collections import defaultdict
from warnings import filterwarnings


def search_workbooks():
    """
    This function requests a path from the user to search through. This then exclusively or inclusively searches for a term.
    """
    path = get_search_directory()

    term: str = input("What term are you looking for? ")
    exclusive: bool = (
        input(
            "Would you like to exclusively search, finding only exactly matching cells? (y/n): "
        )
        .lower()
        .strip()
        == "y"
    )
    case_sensitive: bool = True
    if not exclusive:
        case_sensitive = (
            input("Is the term case sensitive? (y/n): ").lower().strip() == "y"
        )  # This is asked only when a inclusive search is used

    count = 0  # This is the total count of terms found across workbooks initialised

    for file in path.rglob("*.xlsm"):
        if file.is_file and file.suffix == ".xlsm" and "$" not in file.name:
            file_path = Path(path, file)
            count += find_term_in_book(
                file_path, term, exclusive, case_sensitive
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

    path = Path(
        filedialog.askdirectory(title="Please choose a directory")
    )  # This creates an OS file dialogue, asking for a directory

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


def main():
    filterwarnings("ignore", category=UserWarning, module="openpyxl")
    search_workbooks()
    input("\nPress enter to exit.")


if __name__ == "__main__":
    main()
