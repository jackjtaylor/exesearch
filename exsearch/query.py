"""
This module holds the Query class, which stores terms, matches and more for searches.
"""

from collections import defaultdict
from datetime import datetime
from pathlib import Path
from tkinter import Tk, filedialog


class Query:
    """
    This class holds properties related to a search.
    """

    def __init__(self) -> None:
        self.timestamp = datetime.now()
        self.matches = defaultdict(lambda: "")

    # This defines type-hints for later variables to be added.
    term: str
    path: Path
    is_exclusive: bool
    is_case_sensitive: bool
    matches: defaultdict

    def get_found_count(self) -> int:
        """
        This function returns the count of how many matches were found.

        :return: How many matches were found.
        :rtype: int
        """
        count = 0
        for value in self.matches.values():
            count += len(value.strip().split(","))
        return count

    def get_query_parameters(self):
        """
        This function requests a path from the user to search through. This then exclusively or
        inclusively searches for a term.

        :param query: The query to perform.
        :type query: Query
        :return: The search information, along with a timestamp.
        :rtype: Query
        """
        self.path = self.get_search_directory()
        self.term = input("What term are you looking for? ")

        self.is_exclusive = self.get_exclusivity()
        self.is_case_sensitive = True

        if not self.is_exclusive:
            # This is asked only when an inclusive search is used
            self.is_case_sensitive = self.get_case_sensitivity()

        return self

    def get_case_sensitivity(self) -> bool:
        """
        This function asks the user if the search should be case sensitive.

        :return: If the search is case sensitive.
        :rtype: bool
        """
        return input("Is the term case sensitive? (y/n): ").lower().strip() == "y"

    def get_exclusivity(self) -> bool:
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
        )  # This adds the count of terms found to the count, after searching a workbook.

    def get_search_directory(self):
        """
        This function returns a searching directory chosen by the user.

        :return: The directory to search
        :rtype: Path
        """
        print("Please choose a directory.")

        root = Tk()  # This creates a hidden tkinter root.
        root.withdraw()

        path = Path(
            filedialog.askdirectory(title="Please choose a directory.")
        )  # This shows a file
        # dialogue.

        return path
