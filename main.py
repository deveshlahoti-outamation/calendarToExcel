import re
import string
import os
import glob

import PyPDF2
import fitz

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


days_of_the_week = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]

excel_data = {
    'Date of Service': [],
    'Location': [],
    'Patient Name': [],
    'Medical Number': [],
    'Procedure': [],
    'CPT Codes': [],
    'Comments': [],
}


def extract_text(pdf_path):
    doc = fitz.open(pdf_path)
    page = doc[0]
    text = page.get_text("text")

    return text


def combine_pdf_pages(input_path):
    src = fitz.open(input_path)
    doc = fitz.open()  # empty output PDF

    width = src[0].rect.width  # width of the pages
    total_height = sum([page.rect.height for page in src])  # total height of all pages

    # Create a new page with the total width and height
    page = doc.new_page(-1, width=width, height=total_height)

    y_offset = 0  # offset where to insert the next page

    # Loop over each page in the original PDF
    for spage in src:
        # Calculate the rectangle where to insert the page
        r = fitz.Rect(0, y_offset, width, y_offset + spage.rect.height)
        # Insert the page
        page.show_pdf_page(r, src, spage.number)
        # Update the offset
        y_offset += spage.rect.height

    # Save the new PDF
    doc.save(input_path, garbage=3, deflate=True)


def initialize_data():
    pdf_files = glob.glob(os.path.join('files', '*.pdf'))

    for pdf in pdf_files:
        combine_pdf_pages(pdf)

    return pdf_files


def format_lines(lines):
    lines = [''.join(ch for ch in line if ch in string.printable) for line in lines]
    lines = [line for line in lines if line.strip()]
    return lines


def check_line(line):
    pattern = r"^(?!Location:)[\W]*.*//.*//.*//?.*"
    match = re.search(pattern, line.strip())
    return match


def format_text(file_name):
    text = extract_text(f"{file_name}.pdf")
    lines = text.strip().split("\n")
    lines = format_lines(lines)

    events = []

    for index in range(len(lines)):
        line = lines[index]
        event = ['', '', '', '']

        if check_line(line):
            event[0] = line
            for x in range(lines.index(line), len(lines)):
                if "Location: " in lines[x]:
                    event[1] = lines[x - 1].strip()
                    event[2] = lines[x].strip()
                    try:
                        if not check_line(lines[x + 1]):
                            if not any(day in lines[x + 1] for day in days_of_the_week):
                                event[3] = lines[x + 1].strip()
                        break
                    except Exception as e:
                        print("file finished")
            events.append(event)

    return events


def format_events(events, df):
    for event in events:
        # print (event)
        comments = ''.join(char for char in event[3] if char in string.printable)

        # First section of array
        current = event[0].split("//")
        location = current[0].strip()
        medical_number = current[1].strip()
        patient_name = current[2].strip()

        # Second section of array
        current = event[1]
        pattern = r'\d+/\d+/\d+'
        match = re.search(pattern, current)
        date = (match.group()).strip()

        # Third section of array
        current = event[2].split("//")
        procedure = (current[0][len("Location: "):]).strip()

        new_row = pd.DataFrame({
            'Date of Service': [date],
            'Location': [location],
            'Patient Name': [patient_name],
            'Medical Number': [medical_number],
            'Procedure': [procedure],
            'CPT Codes': [''],
            'Comments': [comments]
        })
        df = pd.concat([df, new_row], ignore_index=True)
    return df


def create_excel(df, file_name):
    df.to_excel(f"output/{file_name}.xlsx", index=False, engine='openpyxl')
    book = load_workbook(f"output/{file_name}.xlsx")
    sheet = book.active
    for column in sheet.columns:
        max_length = 0
        column = [cell for cell in column]
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = max_length
        sheet.column_dimensions[get_column_letter(column[0].column)].width = adjusted_width

    # Save the workbook
    book.save(f"output/{file_name}.xlsx")


def delete_files_in_folder(folder_path):
    # Get a list of all files in the folder
    file_list = os.listdir(folder_path)

    # Iterate through the list and delete each file
    for file_name in file_list:
        file_path = os.path.join(folder_path, file_name)
        if os.path.isfile(file_path):
            os.remove(file_path)
            print(f"Deleted file: {file_name}")
        else:
            print(f"Skipping non-file item: {file_name}")


def clean_files():
    delete_files_in_folder("files")
    delete_files_in_folder("output/files")


def main():
    files_dir = 'files'
    if not os.path.exists(files_dir):
        os.makedirs(files_dir)

    pdf_files = initialize_data()

    for pdf in pdf_files:
        df = pd.DataFrame(excel_data)

        file_name = os.path.splitext(pdf)[0]
        print(file_name)
        events = format_text(file_name)
        df = format_events(events, df)
        create_excel(df, file_name)
