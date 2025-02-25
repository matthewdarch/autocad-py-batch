import os
from time import sleep
import win32com.client

# Path to the parent directory containing all projects
user_input = input("Please enter the drawings directory.")
project_root = r'%s' % user_input
# Path to the LISP file
lisp_file = input("Input Lisp File : ")
# LISP command name
lisp_command = "SETUPSITELAYERANDBINDXREF"
# Tags to filter for in the file names
tags = ["DWG300", "DWG350", "DWG380"]

def find_dwg_files_with_tags(root_folder, tags):
    """find all DWG files with specific tags in their names."""
    dwg_files = []
    for dirpath, _, filenames in os.walk(root_folder):
        for file in filenames:
            if file.endswith(".dwg") and any(tag in file for tag in tags):
                dwg_files.append(os.path.join(dirpath, file))
    return dwg_files

def create_bound_filename(original_path):
    """Generate a new file name with '_BOUND' before the extension."""
    dir_name, file_name = os.path.split(original_path)
    name, ext = os.path.splitext(file_name)
    new_name = f"{name}_BOUND{ext}"
    return os.path.join(dir_name, new_name)

def run_lisp_on_drawings(dwg_files, lisp_file, lisp_command):
    acad = win32com.client.Dispatch("AutoCAD.Application.25")
    acad.Visible = False  # This would be only used for debugging purposes, leave false for better performance. (having this set to true I find that AutoCAD can be slow & tend to lag the script behind.)

    for dwg_file in dwg_files:

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
