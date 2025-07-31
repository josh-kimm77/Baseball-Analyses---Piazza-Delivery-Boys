import os
import argparse
from docx2pdf import convert
import time # Import the time module for human-readable timestamps (optional, for printing)

def convert_docx_in_directory_if_newer(input_directory_path, output_directory_path):
    """
    Converts .docx files from an input directory to .pdf format in an output directory,
    only if the DOCX file has been modified more recently than its corresponding PDF,
    or if the PDF does not exist.

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
    skipped_files_count = 0
    errors_count = 0

    for filename in os.listdir(input_directory_path):
        if filename.endswith(".docx"):
            docx_files_found += 1
            docx_filepath = os.path.join(input_directory_path, filename)
            
            # Construct the potential PDF filename and path
            pdf_filename = filename.replace(".docx", ".pdf")
            pdf_filepath = os.path.join(output_directory_path, pdf_filename)

            # Get modification times
            docx_mtime = os.path.getmtime(docx_filepath)
            
            # Check if PDF exists and its modification time
            pdf_exists = os.path.exists(pdf_filepath)
            pdf_mtime = os.path.getmtime(pdf_filepath) if pdf_exists else 0 # Use 0 if PDF doesn't exist

            if not pdf_exists or docx_mtime > pdf_mtime:
                # DOCX is newer or PDF doesn't exist, so convert
                print(f"Converting: '{filename}' (DOCX last modified: {time.ctime(docx_mtime)})")
                if pdf_exists:
                    print(f"    (Existing PDF last modified: {time.ctime(pdf_mtime)})")
                
                try:
                    convert(docx_filepath, pdf_filepath)
                    print(f"Successfully converted: '{filename}' and saved to '{pdf_filepath}'")
                    converted_files_count += 1
                except Exception as e:
                    print(f"Error converting '{filename}': {e}")
                    errors_count += 1
            else:
                # PDF is up-to-date
                print(f"Skipping: '{filename}' (PDF is up-to-date)")
                skipped_files_count += 1

    print("\n--- Conversion Summary ---")
    print(f"Total .docx files found: {docx_files_found}")
    print(f"Files successfully converted: {converted_files_count}")
    print(f"Files skipped (PDF up-to-date): {skipped_files_count}")
    print(f"Files with errors during conversion: {errors_count}")

# --- How to use the script with command-line arguments ---
if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Convert all .docx files in a specified input directory to .pdf format "
                    "in an output directory, converting only if the DOCX is newer or PDF is missing."
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
    convert_docx_in_directory_if_newer(args.input_dir, args.output_dir)