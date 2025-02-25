# This is the base code, I do have another version that will utilise Flask (open-source Python web framework) to create UI friendly input.

import os
from time import sleep
import win32com.client

# Path to the parent directory containing all dwg files.
user_input = input("Please enter the drawings directory.")
project_root = r'%s' % user_input
# Path to the LISP file
lisp_file = input("Input Lisp File : ") # This can be changed to a dedicated lisp, so there is more security.
# LISP command name
lisp_command = "SETUPSITELAYERANDBINDXREF" # This can all be customised to use whatever lisp needs to be used. for this example i'm using my SITE BINDING SCRIPT & LAYER MAKING.
# Tags to filter for in the file names, this looks out for anything containing the below, also can be customised.
tags = ["DWG300", "DWG350", "DWG380"]

# This will open the directory that had been provided & for each file ending with .dwg & the tags that are provided below will be selected.
def find_dwg_files_with_tags(root_folder, tags):
    """find all DWG files with specific tags in their names."""
    dwg_files = []
    for dirpath, _, filenames in os.walk(root_folder):
        for file in filenames:
            if file.endswith(".dwg") and any(tag in file for tag in tags):
                dwg_files.append(os.path.join(dirpath, file))
    return dwg_files
# This definition will split the path + file-name to make it into a variable, allowing us to create the new anme
def create_bound_filename(original_path):
    """Generate a new file name with '_BOUND' before the extension."""
    dir_name, file_name = os.path.split(original_path)
    name, ext = os.path.splitext(file_name)
    new_name = f"{name}_BOUND{ext}"
    return os.path.join(dir_name, new_name)
# this creates a instance of autocad, this will need to be changed to represent the current version of AutoCad, such as AutoCad.Application.24(2024). Within this script, it's not ideal to use AutoCAD LT as they do not have ActiveX Automation support.
def run_lisp_on_drawings(dwg_files, lisp_file, lisp_command):
    acad = win32com.client.Dispatch("AutoCAD.Application.25")
    acad.Visible = False  # This would be only used for debugging purposes, leave false for better performance. (having this set to true I find that AutoCAD can be slow & tend to lag the script behind.)
    
    for dwg_file in dwg_files: #For every dwg file found in directory provided, it will open the document, run the lisp and then save the file 

        print(f"Processing: {dwg_file}")
        try:
            # Open the DWG file
            doc = acad.Documents.Open(dwg_file)
            sleep(5)
            # Load and run the LISP command
            acad.ActiveDocument.SendCommand(f"(load \"{lisp_file}\")\n")
            acad.ActiveDocument.SendCommand(f"{lisp_command}\n")

            # Save the drawing with _BOUND suffix
            bound_file = dwg_file.replace(".dwg", "_BOUND.dwg")
            doc.SaveAs(bound_file)

        except Exception as e:
            print(f"Error processing {dwg_file}: {e}")
        finally:
            try:
                doc.Close(False)  # False = do not save changes, we want it to create a new file with a suffix of _BOUND at the end.
            except Exception:
                pass

# Main script execution
if __name__ == "__main__":
    dwg_files = find_dwg_files_with_tags(project_root, tags)
    print(f"Found {len(dwg_files)} DWG files with specified tags.")
    run_lisp_on_drawings(dwg_files, lisp_file, lisp_command)
