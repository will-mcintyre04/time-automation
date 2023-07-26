# Time Study Automation

**Author:** Will McIntyre  
**Date:** 2023-07-10

---

## Description

This script provides a graphical user interface (GUI) for automating the Metalumen time study process. It allows the user to perform two main actions:

1. Creating a new time tracking template folder in a specified directory along with an automated time-tracking spreadsheet (see `references/Time Tracking Template v3`) and `Videos` folder. See an example directory tree below:
```
|-- RM2DU-4ft 2023-07-24
|   |-- RM2DU-4ft 2023-07-24.xlsm
|   `-- Videos
|       `-- IMG_0966.mp4
|-- RM4D-11ft 2023-07-24
|   |-- RM4D-11ft 2023-07-24.xlsm
|   `-- Videos
|-- RM4DOD-6ft 2023-07-24
|   |-- RM4DOD-6ft 2023-07-24.xlsm
|   `-- Videos
|       `-- IMG_1091.mp4
|-- RML-11ft 2023-07-24
|   |-- RML-11ft 2023-07-24.xlsm
|   `-- Videos
|       `-- IMG_1099.mp4
`-- S6W-5ft 2023-07-24
    |-- S6W-5ft 2023-07-24.xlsm
    `-- Videos
        `-- IMG_1103.mp4
```
2. Updating a specified time-tracking database with a valid spreadsheet.

Note: Please see the [wiki](https://github.com/will-mcintyre04/time-automation/wiki) for more in-depth instruction.

---

## Usage

1. Download the `Time Study Spreadsheet and Database Automation.zip` [here](https://drive.google.com/file/d/1Fy28Ba5x0wQScx8R9H29Ba8LpitO5GhB/view?usp=sharing). Note: Windows might prompt to trust the download before accesing it.
2. Export all .zip contents onto the local disc.
3. Run the application through the `Time Study Automation.exe` executable. Note: Keep this application within the original folder, or in a new one along with the `images` and `references` folders. Consider creating a shortcut to the application path that can be stored anywhere across the machine.

## Tools:

* <img alt="Python" width="30px" style="padding-right:10px;" src="https://cdn.jsdelivr.net/gh/devicons/devicon/icons/python/python-original.svg" /> Python
* <img alt="vsc" width="30px" style="padding-right:10px;" src="https://cdn.jsdelivr.net/gh/devicons/devicon/icons/vscode/vscode-original.svg"/> VSC
* <img alt="Git" width="30px" style="padding-right:10px;" src="https://cdn.jsdelivr.net/gh/devicons/devicon/icons/git/git-original.svg" /> git
* <img alt = "sqlite" width="30px" style="padding-right:10px;" src="https://cdn.jsdelivr.net/gh/devicons/devicon/icons/sqlite/sqlite-original.svg" /> sqlite

## Dependencies:
- `tkinter`: Used to create the graphical user interface (GUI).
- `subprocess`: Used to call external scripts.
- `os`: Used for file and directory operations.
- `PIL` (Python Imaging Library): Adds Tkinter GUI support for opening and manipulating image file formats.
- `sqlite`: Used to connect to the SQLite database.
- `openpyxl`: Used to load and manipulate Excel spreadsheets.
- `webbrowser`: Provides a high-level interface to display web-based documents to users.
## Support

For any issues or inquiries, please contact Will McIntyre at william.d.j.mcintyre@gmail.com.
