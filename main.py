"""
Time Study Main GUI

Author: Will McIntyre
Date: 2023-07-05

Description:
This script provides a graphical user interface (GUI) for automating the Metalumen time
study process. It allows the user to perform two main actions: creating a new spreadsheet
and updating a database with a spreadsheet. The script integrates two separate functionalities
provided by the following two scripts: make_new_timestudy.py and database_update.py. Test 2123

CLICK THE QUESTION MARK BUTTON FOR DOCUMENTATION

Usage (See Documentation for further details):

Launch the Time Study Main GUI by running the main.py script.
The main window titled "Time Study Tools" will appear.
Click the "Create Spreadsheet" button to generate a new time study spreadsheet using the make_new_timestudy.py script.
Click the "Update Database" button to update a database with a valid time study spreadsheet using the database_update.py script.
Optionally, click the question mark button at the bottom of the window to open the Metalumen time study documentation.

"""

import tkinter as tk
from tkinter import messagebox
import subprocess
from PIL import Image, ImageTk
from database_update import DatabaseUpdater
from make_new_timestudy import TimeStudyCreator

def create_spreadsheet():
    """ 
    Calls the make_new_timestudy.py script from the same directory
    that main.py is in
    """
    creator = TimeStudyCreator()
    creator.main()

def update_database():
    """ 
    Calls the update_database.py script from the same directory
    that main.py is in
    """
    updater = DatabaseUpdater()
    updater.main()

def open_documentation():
    subprocess.Popen(["J:\\6.0 - Designer Folders\Will\Analysis\COBOT Research.docx"], shell=True)

def main():
    root = tk.Tk()
    root.title("Time Study Main GUI")
    root.geometry("400x250+760+400")

    main_title = tk.Label(root, text="Time Study Tools", font=("Arial", 16, "bold"))
    main_title.pack()
    sub_title = tk.Label(root, text = "Automate the Metalumen time study process. Click below to create a new spreadsheet or update a valid database with a spreadsheet.",
                         wraplength=375)
    sub_title.pack()

    create_button = tk.Button(root, text="Create Spreadsheet", bg = "#72B5E0",
                              command=create_spreadsheet, cursor="hand2")
    create_button.pack(pady=20)

    update_button = tk.Button(root, text="Update Database", bg = "#72B5E0",
                              command=update_database, cursor="hand2")
    update_button.pack(pady=20)

    # Insert button for documentation access
    image = Image.open("J:\\6.0 - Designer Folders\Will\Code Tools\Time Study Spreadsheet and Database Automation\question_mark.png")
    resized_image = image.resize((20, 20))
    tk_image = ImageTk.PhotoImage(resized_image)
    question_button = tk.Button(root, image=tk_image, cursor="hand2", command=open_documentation)
    question_button.image = tk_image  # To prevent the image from being garbage collected
    question_button.pack(side=tk.BOTTOM, pady=10)

    root.mainloop()

if __name__ == "__main__":
    main()
