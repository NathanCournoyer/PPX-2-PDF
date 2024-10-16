import sys, os
from typing import final
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

    # Variables
    pdf_str_dir: str = ""
    ppx_str_dir: str = ""
    exit_code = 0
    
    # Defining script arguments;
    PPX_2_PDF = sys.argv[0]

    # No real arg handling for now, one teacher sole reason this exists, const dir
    print(f"Script {PPX_2_PDF} running: {sys.argv}")
    
    if len(sys.argv) > 1:
        # The PowerPoint folder provided.
        print('Found PowerPoint dir arg.')
        ppx_str_dir = sys.argv[1]
    else : 
        # Default location
        ppx_str_dir= DEFAULT_PPX_DIR

    if len(sys.argv) > 2:
        # The PDF folder provided.
        print('Found PDF dir arg.')
        pdf_str_dir = sys.argv[2]
    else:
        # Default PDF location
        pdf_str_dir = DEFAULT_PDF_DIR

    
    # Iterables
    ppx_bytes_dir: bytes  = os.fsencode(ppx_str_dir) 
    pdf_bytes_dir: bytes = os.fsencode(pdf_str_dir)

    # https://stackoverflow.com/a/10378012 CC BY-SA 4.0
    for file in os.listdir(ppx_bytes_dir):
        filename = os.fsdecode(file)
        input_filename = os.path.join(ppx_str_dir, filename)
        output_filename = os.path.join(pdf_str_dir, filename)
        if (filename.endswith(".pptx") or filename.endswith(".ppt")) and is_ppx_file_in_pdf_dir(filename, pdf_bytes_dir) == False:
            try: 
                print(f"{filename}.pdf will be placed in {output_filename}")
                PPTtoPDF(input_filename, output_filename)
            except Exception as e:
                print(f"Oopsie daisies: {e}")
                exit_code = -1
        else:
            print(f"{filename} is not a valid candidate for conversion.")
    
    print(f"Script {PPX_2_PDF} exiting: {exit_code}")
    return exit_code

# Sequential unsorted search pdf dir, check if the ppx is converted already
def is_ppx_file_in_pdf_dir(ppx_filename: str, pdf_bytes_dir: bytes) -> bool:
    file_is_present = False
    pdf_files = os.listdir(pdf_bytes_dir)
    for file in pdf_files:
        filename = os.fsdecode(file)
        if filename == ppx_filename + ".pdf":
            return True
    return file_is_present

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