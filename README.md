# Time Study Automation

**Author:** Will McIntyre  
**Date:** 2023-07-10

---

## Description

This script provides a graphical user interface (GUI) for automating the Metalumen time study process. It allows the user to perform two main actions:

1. Creating a new time tracking template folder in a specified directory along with a spreadsheet (see `references/Time Tracking Template v3`) and `Videos` folder.
2. Updating a specified time-tracking database with a valid spreadsheet.

Note: Please see the [wiki](https://github.com/will-mcintyre04/time-automation/wiki) for more in-depth instruction.

---

## How to Use

1. Download the "Time Study Spreadsheet and Database Automation.zip" [`here`](https://drive.google.com/file/d/1ryz9c-6N8hgT8R8zdiBeNg5aCCKWsj1G/view?usp=sharing). Note: Windows might prompt to trust the download before accesing it.
2. Export all of the folders onto the local disc.
3. Run the application through the "main.exe" executable. Note: Keep this application within the original folder, or in a new one along with the `build` `images` and `references` folders.

## BTools:

* <img alt="Python" width="30px" style="padding-right:10px;" src="https://cdn.jsdelivr.net/gh/devicons/devicon/icons/python/python-original.svg" /> Python
* <img alt="vsc" width="30px" style="padding-right:10px;" src="https://cdn.jsdelivr.net/gh/devicons/devicon/icons/vscode/vscode-original.svg"/> VSC
* <img alt="Git" width="30px" style="padding-right:10px;" src="https://cdn.jsdelivr.net/gh/devicons/devicon/icons/git/git-original.svg" /> git
* <img alt = "sqlite" width="30px" style="padding-right:10px;" src="https://cdn.jsdelivr.net/gh/devicons/devicon/icons/sqlite/sqlite-original.svg" /> sqlite

## Dependencies:
- `tkinter`: Used to create the graphical user interface (GUI).
- `subprocess`: Used to call external scripts.
- `os`: Used for file and directory operations.
- `PIL` (Python Imaging Library): Used to display the question mark button image in the GUI.
- `sqlite`: Used to connect to the SQLite database.
- `openpyxl`: Used to load and manipulate Excel spreadsheets.
## Support

For any issues or inquiries, please contact Will McIntyre at william.d.j.mcintyre@gmail.com.
