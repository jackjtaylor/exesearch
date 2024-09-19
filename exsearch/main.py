"""
This module runs exsearch as a command-line tool.
"""

from warnings import filterwarnings
from query import Query
from searcher import Searcher


def main():
    filterwarnings("ignore", category=UserWarning, module="openpyxl")

    query = Query()
    query.get_query_parameters()

    searcher = Searcher()
    searcher.run_query(query)

    input("\nPress enter to exit.")


if __name__ == "__main__":
    main()
