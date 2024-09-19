"""
This module holds the Query class, which stores terms, matches and more for searches.
"""

from collections import defaultdict
from datetime import datetime
from pathlib import Path


class Query:
    """
    This class holds properties related to a search.
    """

    def __init__(self) -> None:
        self.timestamp = datetime.now()
        self.matches = defaultdict(lambda: "")
        self.matches["test"] = "A1, A2"

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
