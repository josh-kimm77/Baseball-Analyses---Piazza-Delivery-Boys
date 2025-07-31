import os
import argparse
from docx2pdf import convert

def convert_docx_in_directory(input_directory_path, output_directory_path):
    """
    Converts all .docx files from an input directory to .pdf format
    and saves them in a specified output directory.

    Args:
        input_directory_path (str): The path to the directory containing DOCX files.
        output_directory_path (str): The path to the directory where PDF files will be saved.
    """
    if not os.path.isdir(input_directory_path):
        print(f"Error: Input directory '{input_directory_path}' does not exist.")
        return

    # Create the output directory if it doesn't exist
    if not os.path.exists(output_directory_path):
        os.makedirs(output_directory_path)
        print(f"Created output directory: '{output_directory_path}'")

    print(f"Scanning input directory: {input_directory_path} for .docx files...")
    docx_files_found = 0
    converted_files_count = 0
    errors_count = 0

    for filename in os.listdir(input_directory_path):
        if filename.endswith(".docx"):
            docx_files_found += 1
            docx_filepath = os.path.join(input_directory_path, filename)
            
            # Create the output PDF filename by replacing .docx with .pdf
            pdf_filename = filename.replace(".docx", ".pdf")
            pdf_filepath = os.path.join(output_directory_path, pdf_filename)

            print(f"Attempting to convert: '{filename}' to '{pdf_filename}'...")
            try:
                convert(docx_filepath, pdf_filepath)
                print(f"Successfully converted: '{filename}' and saved to '{pdf_filepath}'")
                converted_files_count += 1
            except Exception as e:
                print(f"Error converting '{filename}': {e}")
                errors_count += 1

    print("\n--- Conversion Summary ---")
    print(f"Total .docx files found: {docx_files_found}")
    print(f"Files successfully converted: {converted_files_count}")
    print(f"Files with errors during conversion: {errors_count}")

# --- How to use the script with command-line arguments ---
if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Convert all .docx files in a specified input directory to .pdf format "
                    "and save them in an output directory."
    )

    parser.add_argument(
        "input_dir",
        type=str,
        help="The full path to the INPUT directory containing .docx files."
    )
    parser.add_argument(
        "output_dir",
        type=str,
        help="The full path to the OUTPUT directory where .pdf files will be saved."
    )

    args = parser.parse_args()

    # Call the conversion function with the parsed arguments
    convert_docx_in_directory(args.input_dir, args.output_dir)