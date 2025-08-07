import sys
import zipfile
import os
import tempfile
from oletools import olevba

def extract_macros_from_xlsm(file_path):
    """
    Extracts VBA macros from an .xlsm file using oletools as a library.
    This function is designed to be portable and can be used in environments
    like Databricks notebooks.

    :param file_path: Path to the .xlsm file.
    :return: A list of dictionaries, where each dictionary contains
             'stream_path', 'vba_filename', and 'vba_code'.
             Returns an empty list if no macros are found or in case of an error.
    """
    if not os.path.exists(file_path):
        print(f"Error: File not found at {file_path}")
        return []

    macros = []

    # Create a temporary directory to extract vbaProject.bin
    with tempfile.TemporaryDirectory() as temp_dir:
        try:
            with zipfile.ZipFile(file_path, 'r') as zip_ref:
                vba_project_path_in_zip = 'xl/vbaProject.bin'
                if vba_project_path_in_zip not in zip_ref.namelist():
                    print("No vbaProject.bin found in the .xlsm file. It may be macro-free.")
                    return []

                # Extract vbaProject.bin to the temporary directory
                zip_ref.extract(vba_project_path_in_zip, path=temp_dir)
                vba_project_path = os.path.join(temp_dir, vba_project_path_in_zip)

                # Use olevba's VBA_Parser to analyze the file
                vba_parser = olevba.VBA_Parser(vba_project_path)
                if vba_parser.detect_vba_macros():
                    print("VBA Macros found, extracting...")
                    for (filename, stream_path, vba_filename, vba_code) in vba_parser.extract_macros():
                        macros.append({
                            "stream_path": stream_path,
                            "vba_filename": vba_filename,
                            "vba_code": vba_code
                        })
                else:
                    print("No VBA macros were detected by the parser.")
                vba_parser.close()

        except zipfile.BadZipFile:
            print(f"Error: '{file_path}' is not a valid zip file (or .xlsm file).")
            return []
        except Exception as e:
            print(f"An unexpected error occurred: {e}")
            return []

    return macros

def main():
    """
    Main function to run the macro extractor from the command line.
    """
    if len(sys.argv) != 2:
        print("Usage: python macro_parser.py <path_to_xlsm_file>")
        sys.exit(1)

    file_to_parse = sys.argv[1]

    print(f"Analyzing {file_to_parse}...")
    extracted_macros = extract_macros_from_xlsm(file_to_parse)

    if not extracted_macros:
        print("No macros were extracted.")
        return

    print("\n--- Extracted Macros ---")
    for i, macro_info in enumerate(extracted_macros, 1):
        print(f"\n--- Macro #{i} ---")
        print(f"Stream Path: {macro_info['stream_path']}")
        print(f"VBA Filename: {macro_info['vba_filename']}")
        print("--- Code ---")
        print(macro_info['vba_code'])
        print("--- End Code ---")

if __name__ == "__main__":
    main()
