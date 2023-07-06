"""
Time Study Main GUI

Author: Will McIntyre
Date: 2023-07-05

Description:
This script provides a graphical user interface (GUI) for automating the Metalumen time
study process. It allows the user to perform two main actions: creating a new spreadsheet
and updating a database with a spreadsheet. The script integrates two separate functionalities
provided by the following two scripts: make_new_timestudy.py and database_update.py.

CLICK THE QUESTION MARK BUTTON FOR DOCUMENTATION

Usage (See Documentation for further details):

Launch the Time Study Main GUI by running the main.py script.
The main window titled "Time Study Tools" will appear.
Click the "Create Spreadsheet" button to generate a new time study spreadsheet using the make_new_timestudy.py script.
Click the "Update Database" button to update a database with a valid time study spreadsheet using the database_update.py script.
Optionally, click the question mark button at the bottom of the window to open the Metalumen time study documentation.

"""

import tkinter as tk
import subprocess
from PIL import Image, ImageTk
from database_update import DatabaseUpdater
from make_new_timestudy import TimeStudyCreator

class MainApplication:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Time Study Main GUI")
        self.root.geometry("400x250+760+400")

        self.create_widgets()

    def create_spreadsheet(self):
        """ 
        Calls the make_new_timestudy.py script to create a spreadsheet and folder.
        """
        creator = TimeStudyCreator()
        creator.main()

    def update_database(self):
        """ 
        Calls the update_database.py script to update a database.
        """
        updater = DatabaseUpdater()
        updater.main()

    def open_documentation(self):
        subprocess.Popen(["J:\\6.0 - Designer Folders\Will\Analysis\COBOT Research.docx"], shell=True)

    def create_widgets(self):
        """ 
        Populates the main GUI page with widgets.
        """
        icon = Image.open("images/open-book.png")
        resized_icon = icon.resize((32, 32))
        tk_icon = ImageTk.PhotoImage(resized_icon)
        self.root.iconphoto(True, tk_icon)

        main_title = tk.Label(self.root, text="Time Study Tools", font=("Arial", 16, "bold"))
        main_title.pack()
        sub_title = tk.Label(self.root, text = "Automate the Metalumen time study process. Click below to create a new spreadsheet or update a valid database with a spreadsheet.",
                            wraplength=375)
        sub_title.pack()

        create_button = tk.Button(self.root, text="Create Spreadsheet", bg = "#72B5E0",
                                command=self.create_spreadsheet, cursor="hand2")
        create_button.pack(pady=20)

        update_button = tk.Button(self.root, text="Update Database", bg = "#72B5E0",
                                command=self.update_database, cursor="hand2")
        update_button.pack(pady=20)

        # Insert button for documentation access
        image = Image.open("images\question_mark.png")
        resized_image = image.resize((20, 20))
        tk_image = ImageTk.PhotoImage(resized_image)
        question_button = tk.Button(self.root, image=tk_image, cursor="hand2", command=self.open_documentation)
        question_button.image = tk_image  # To prevent the image from being garbage collected
        question_button.pack(side=tk.BOTTOM, pady=10)
    def start (self):
        """ 
        Generates the GUI for the main application.
        """
        self.root.mainloop()
