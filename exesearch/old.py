def find_term_in_book(
    book_path: Path, term: str, exclusive: bool, case_sensitive: bool
) -> int:
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

    return count


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
