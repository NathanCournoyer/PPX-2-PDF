import sys, os # System stuff
from typing import final # limited const implementation 
from comtypes import client # type: ignore VS Code Type Warnings 

    # TODO ppx and pdf dir validation
    # TODO Sorting/Search algorithms implementation for optimization
    # TODO Script arg handling
    # TODO Error Handling
    # TODO Logger and styling

# Main
def main() -> int:
    # CONST
    DEFAULT_PDF_DIR: final = "G:\Mon disque\shared\\2024\AUT\COM120\PDF"
    DEFAULT_PPX_DIR: final = "G:\Mon disque\shared\\2024\AUT\COM120\PPX"
    PPX_2_PDF = sys.argv[0]

    # Variables
    pdf_str_dir: str = ""
    ppx_str_dir: str = ""
    exit_code: int = 0
    valid_file_path: bool = False    

    # No real arg handling for now, one teacher sole reason this exists, const dir
    print(f"Script {PPX_2_PDF} running: {sys.argv}")
    
    if len(sys.argv) > 1:
        # The PowerPoint folder provided as 2nd arg.
        print('Found PowerPoint dir arg.')
        ppx_str_dir = sys.argv[1]
    else : 
        # Default location
        ppx_str_dir= DEFAULT_PPX_DIR

    if len(sys.argv) > 2:
        # The PDF folder provided as 3rd arg.
        print('Found PDF dir arg.')
        pdf_str_dir = sys.argv[2]
    else:
        # Default PDF location
        pdf_str_dir = DEFAULT_PDF_DIR

    while (not valid_file_path):
        try :
            # Iterables file lists as bytes lists
            ppx_bytes_dir: bytes  = os.fsencode(ppx_str_dir) 
            pdf_bytes_dir: bytes = os.fsencode(pdf_str_dir)

            # Valid File Path Loop Exit
            valid_file_path = True
        except FileNotFoundError as file_not_found_error:
            print(f"File not found: {file_not_found_error}")
            ppx_bytes_dir = input("Enter PowerPoint source Directory: ")
            pdf_bytes_dir = input("Enter PDF target Directory: ")



    # https://stackoverflow.com/a/10378012 CC BY-SA 4.0
    for file in os.listdir(ppx_bytes_dir):
        filename = os.fsdecode(file)
        input_filename = os.path.join(ppx_str_dir, filename)
        output_filename = os.path.join(pdf_str_dir, filename)
        if (filename.endswith(".pptx") or filename.endswith(".ppt")) and is_ppx_file_in_pdf_dir(filename, pdf_bytes_dir) == False:
            try: 
                print(f"{filename}.pdf will be placed in {output_filename}")
                PPTtoPDF(input_filename, output_filename)
            except Exception as file_error:
                print(f"Oopsie daisies: {file_error}. It seems like the path is wrong or inaccessible.")
                exit_code = -1
        else:
            print(f"{filename} is not a valid candidate for conversion.")
    
    print(f"Script {PPX_2_PDF} exiting: {exit_code}")
    return exit_code

# Sequential search through pdf dir, check if the ppx is converted already
def is_ppx_file_in_pdf_dir(ppx_filename: str, pdf_bytes_dir: bytes) -> bool:
    # Iterable list[bytes] representing PDF files.
    pdf_files = os.listdir(pdf_bytes_dir)
    for file in pdf_files:
        filename = os.fsdecode(file)
        if filename == ppx_filename + ".pdf":
            return True
    return False

# https://stackoverflow.com/a/31624001 CC BY-SA 4.0
def PPTtoPDF(inputFileName, outputFileName):
    powerpoint = client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1

    if outputFileName[-3:] != 'pdf':
        outputFileName = outputFileName + ".pdf"
    deck = powerpoint.Presentations.Open(inputFileName)
    deck.SaveAs(outputFileName, 32) # formatType = 32 for ppt to pdf 
    # https://learn.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype 
    # deck.Close() Unrequired, causes errors
    powerpoint.Quit() # Does the job ^

if __name__ == '__main__':
    sys.exit(main())  