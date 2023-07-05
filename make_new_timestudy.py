"""
Time Study Folder and Spreadsheet Creator

Author: Will McIntyre
Date: 2023-06-26

Description:
This script allows the user to create a new folder and copy a time tracking spreadhseet with the same name into it. 
The script provides a graphical user interface (GUI) for the user to input the file name and destination folder path, or
browse for the destination folder.

Usage (See Documentation for further details):

Enter the file name (without a date) in the input field.
Enter the destination folder path or leave it blank to use the default (J:/6.0 - Designer Folders/Will/Test).
Click the 'Browse' button to select the destination folder interactively.
Click the 'Create Folder and Spreadsheet' button to execute the script.

Note: The default source file path and other constants used in this script are specific to the author's directory
structure and can be modified according to your own requirements.
"""

import tkinter as tk
from tkinter import messagebox, filedialog
import shutil
import os
from datetime import date

class EmptyInputError(Exception):
    pass

class TimeStudyCreator:
    DEFAULT_FOLDER_DEST = "J:/6.0 - Designer Folders/Will/Time Study Forms/New Time Studies"
    DEFAULT_TIME_TEMPLATE = "J:/6.0 - Designer Folders/Will/Time Study Forms/Time Tracking Template v3.xlsm"

    def __init__(self):
        self.window = None
        self.entry = None
        self.destination_entry = None


    def create_folder_and_copy_spreadsheet(self):
        """
        Creates the new folder given an inputted name and copies a spreadsheet (Time Tracking Template v3) with the same 
        name into it. Additionally, creates a 'Videos' folder within this new folder.
        """
        file_name = self.entry.get()

        today = date.today()  # formatted to "YYYY-MM-DD"
        formatted_date = today.strftime("%Y-%m-%d")
        file_name_with_date = f"{file_name} {formatted_date}"

        # Source file path and destination folder path input or default
        source_file_path = self.DEFAULT_TIME_TEMPLATE
        destination_folder_path = self.destination_entry.get() or self.DEFAULT_FOLDER_DEST

        folder_path = os.path.join(destination_folder_path, file_name_with_date)

        # Make the directory, if it already exists or is invalid, display an error message and break from the function
        try:
            if not os.path.exists(destination_folder_path):
                raise FileNotFoundError()
            if file_name == "":
                raise EmptyInputError()
            os.makedirs(folder_path)
        except FileNotFoundError:
            messagebox.showerror("Error", f"the directory '{destination_folder_path}' does not exist.")
            return
        except FileExistsError:
            messagebox.showerror("Error", f"The folder '{folder_path}' already exists.")
            return
        except EmptyInputError:
            messagebox.showerror("Error", "Please enter a file name.")
            return

        # Create "Videos" folder within the newly created folder
        videos_path = os.path.join(folder_path, "Videos")
        os.makedirs(videos_path)

        # Copy the source file to the destination folder with the user-provided name
        destination_file_path = os.path.join(folder_path, f"{file_name_with_date}.xlsm")
        shutil.copy2(source_file_path, destination_file_path)

        # Display a message box with the success message
        message = f"Folder '{folder_path}' and the renamed spreadsheet '{file_name}' have been created."
        messagebox.showinfo("Success", message)

        # Close the main window
        self.window.destroy()
        return destination_file_path

    def create_and_open_spreadsheet(self):
        file_path = self.create_folder_and_copy_spreadsheet()
        if file_path:
            os.startfile(file_path)

    def browse_folder(self):
        """
        Gets a file path selected from the user and fills in the destination_entry widget
        """
        folder_path = filedialog.askdirectory()
        self.destination_entry.delete(0, tk.END)
        self.destination_entry.insert(tk.END, folder_path)
    
    def destroy_window(self):
        """ Exits the current window. """
        self.window.destroy()

    def main(self):
        # Create the main (root) window and set up attributes
        self.window = tk.Tk()
        self.window.geometry("400x250+760+400")  # 400 x 250 pixels, 760 pixels across and 400 down
        self.window.title("Time Study Folder and Spreadsheet Creator")

        # Create a label and an entry field for the file name
        label = tk.Label(self.window, text="Enter the Name of the File (With No Date):")
        label.pack()
        self.entry = tk.Entry(self.window)
        self.entry.pack()

        # Create a label and an entry field for the destination folder path
        destination_label = tk.Label(self.window, text="Enter the Destination Folder Path:")
        destination_label.pack()
        self.destination_entry = tk.Entry(self.window, width=60)
        self.destination_entry.insert(tk.END, self.DEFAULT_FOLDER_DEST)
        self.destination_entry.pack()

        # Create a button to browse for the destination folder
        browse_button = tk.Button(self.window, text="Browse", command=self.browse_folder, cursor="hand2")
        browse_button.pack()

        # Create buttons to execute the script
        create_button = tk.Button(self.window, text="Create Folder and Spreadsheet",
                           command=self.create_folder_and_copy_spreadsheet, cursor="hand2")
        create_button.pack(pady=(30, 0))

        open_button = tk.Button(self.window, text="Create and Open Spreadsheet",
                           command=self.create_and_open_spreadsheet, bg="#72B5E0", cursor="hand2")
        open_button.pack(pady=(10, 0))
        exit_button = tk.Button(self.window, text = "Exit", bg = "#FB5858", command=self.destroy_window, cursor="hand2")
        exit_button.pack(pady=(5, 0))

        # Start the GUI event loop
        self.window.mainloop()

if __name__ == "__main__":
    creator = TimeStudyCreator()
    creator.main()