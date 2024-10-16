# STOP GIVING CLASS NOTES AS POWERPOINTS

# PPX-2-PDF
Automatically converts pptx or ppt files into pdf files. This project is only intended for personal use but is open to improvements and could be updated. 

The PowerPoint and PDF directories are respectively set as script arguments or can be hardcoded into constant variables DEFAULT_PDF_DIR and DEFAULT_PPX_DIR, if no arguments are provided, default directories are used but if the directories aren't correct, the program fails. 

The program will not convert the same file name twice. If the resulting PDF name is the same as the original name, it doesn't try to convert it into a PDF. 

# Run
- Python must be installed.
- comtypes must be installed.
In root directory:
`python .\PPX_2_PDF\PPX_2_PDF.py "C:\PowerPoint\Directory" "C:\PDF\Directory"`
## Compatibility
Only compatible with Windows, tested on Windows 11.
Utilizes the [comtypes](https://pythonhosted.org/comtypes/#the-comtypes-package) Python library.
## Todo:
- TODO ppx and pdf dir validation
- TODO Sorting/Search algorithms implementation for optimization
- TODO Script arg handling
- TODO Error Handling
- TODO Logger and styling
- TODO Additional file extension handling
