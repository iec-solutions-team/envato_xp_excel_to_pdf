from datetime import datetime
import os
import zipfile
import re

def handle_folder_creation(folder_path, folder_name):
    has_error = False
    error_msg = ''
    try:
        full_path = os.path.join(folder_path, folder_name)


        # Check if the folder exists
        if os.path.exists(full_path) and os.path.isdir(full_path):
            # If it exists, delete it and its contents
            os.rmdir(full_path)
            print(f"Deleted folder: {folder_name}")

        # Create the folder
        os.mkdir(full_path)
    except Exception as e:
        has_error = True
        error_msg = str(e)
    return has_error, error_msg


def create_zip(source_folder, output_path):
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, _, files in os.walk(source_folder):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, source_folder)
                zipf.write(file_path, arcname)


def process_pdf_filename_format(pdf_folder_path, pdf_folder_name, pdf_filename_format, all_excel_columns, excel_row, excel_row_index):

    print(f"\n\n\n")
    print(f"pdf_folder_path : {pdf_folder_path}")
    print(f"pdf_filename_format : {pdf_filename_format}")
    print(f"all_excel_columns : {all_excel_columns}")
    print(f"excel_row : {excel_row}")
    print(f"excel_row_index : {excel_row_index}")

    # Get the current date
    current_date = datetime.now()
    today_date = current_date.strftime('%Y-%m-%d') # Format as "YYYY/MM/DD"
    timestamp = datetime.now().strftime('%Y-%m-%d %H.%M.%S')

    built_in_formats = {
        "today": today_date,
        # "DD-MM-YYYY": format2,
        "timestamp": timestamp,
        "increment": excel_row_index + 1,
    }


    excel_columns_with_value = {}

    for excel_column in all_excel_columns:
        excel_columns_with_value[excel_column] = excel_row[excel_column]


    all_formats = {**built_in_formats, **excel_columns_with_value}
    print(f"\nall_formats: {all_formats}\n")

    filename = pdf_filename_format.format(**all_formats)
    filename = f"{pdf_folder_path}/{pdf_folder_name}/{filename}.pdf"

    print(f"\n\n>>>>>filename  :{filename}>>>>>>>\n\n")

    return filename


def is_validate_filename_input(text):
    # This regex pattern allows letters, numbers, spaces, underscores, and periods.
    # You can adjust it to suit your specific requirements.
    # pattern = r'^[a-zA-Z0-9\s_\.]*$'
    pattern = r'^[a-zA-Z0-9\s_\.{}]*$'
    return re.match(pattern, text) is not None