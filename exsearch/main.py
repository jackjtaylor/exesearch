"""
This module runs exsearch as a command-line tool.
"""

from warnings import filterwarnings
from query import Query
from searcher import Searcher


def main():
    filterwarnings("ignore", category=UserWarning, module="openpyxl")  # This disables irrelevant
    # warnings.

    query = Query()
    query.get_query_parameters()  #  This uses the previously created query and initialises values.

    searcher = Searcher()
    searcher.run_query(query)  # This runs the passed-in query across a path.

    input("Press enter to exit.")  # This keeps the window open until the user is ready to exit.


if __name__ == "__main__":
    main()
