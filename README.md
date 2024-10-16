# STOP GIVING CLASS NOTES AS POWERPOINTS

# PPX_2_PDF.py
Automatically converts pptx or ppt files into pdf files. This project is only intended for personal use but is open to improvements and could be updated. 

The PowerPoint and PDF directories are respectively set as source and target directories as script arguments or can be hardcoded into constant variables `DEFAULT_PDF_DIR` and `DEFAULT_PPX_DIR`, if no arguments are provided, default directories are used but if the directories aren't correct, asks to re-enter each directory path until they are found on the system and then or the user can enter `exit` to exit the program. 

The program will not convert the same filename twice. If the resulting PDF name is the same as the original name, it doesn't try to convert it into a PDF. 

# To Run
In project root directory with arguments:
- Python must be installed.
`python --version`
- comtypes must be installed.
`python -m pip install -r requirements.txt`
- Run the script PPX_2_PDF.py PPX_DIR: str, PDF_DIR: str
`python .\PPX_2_PDF\PPX_2_PDF.py "C:\PowerPoint\Directory" "C:\PDF\Directory"`

# This can be done with the help of the Windows file explorer 
Right click both soure and target folders to access the "Copy as access path" button and use it to copy the path into as command arguments.
There is no need to put in `""` in advance, on Windows 11 they will be copied already. [More information](https://www.howtogeek.com/670447/how-to-copy-the-full-path-of-a-file-on-windows-10/)

## Compatibility
Only compatible with Windows, tested on Windows 11.
Utilizes the [comtypes](https://pythonhosted.org/comtypes/#the-comtypes-package) Python library.
## Todo:
- TODO Sorting/Search algorithms implementation for optimization
- TODO Logger and styling
- TODO Additional file extension handling
