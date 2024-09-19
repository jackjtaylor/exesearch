"""
This module runs exsearch as a command-line tool.
"""

from warnings import filterwarnings


def main():
    filterwarnings("ignore", category=UserWarning, module="openpyxl")

    search = create_search()
    find_workbooks(search)

    input("\nPress enter to exit.")


if __name__ == "__main__":
    main()
