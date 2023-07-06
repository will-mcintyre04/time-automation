"""
Time Study Main GUI

Author: Will McIntyre
Date: 2023-07-05

Description:
This script provides a graphical user interface (GUI) for automating the Metalumen time
study process. It allows the user to perform two main actions: creating a new spreadsheet
and updating a database with a spreadsheet. The script integrates two separate functionalities
provided by the following two scripts: make_new_timestudy.py and database_update.py.
"""
from application import MainApplication

if __name__ == "__main__":
    app = MainApplication()
    app.start()