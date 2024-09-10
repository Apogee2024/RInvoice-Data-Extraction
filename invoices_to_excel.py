import os
import pandas
import xlsxwriter
from pdf_extractor import invoice_text_extractor, get_po
import datetime
from dotenv import load_dotenv  

# Load environment variables from the .env file
load_dotenv()

def main():
    # Get the current date
    current_date = str(datetime.date.today()).strip()

    # Get the folder path from the environment variable
    folder_path = os.getenv('PDF_FOLDER_PATH')

    # Check if the environment variable is set
    if not folder_path:
        print("Error: The 'PDF_FOLDER_PATH' environment variable is not set.")
        return

    # Check if the folder exists
    if not os.path.isdir(folder_path):
        print(f"Error: The folder '{folder_path}' does not exist.")
        return

    # List files in the folder
    file_names = os.listdir(folder_path)
    new_df = pandas.DataFrame()

    # Iterate over each file in the folder
    for file_name in file_names:
        # Construct the full file path
        file_path = os.path.join(folder_path, file_name)
        print(file_path)

        # Check if the item in the folder is a file (not a subfolder)
        if os.path.isfile(file_path):
            # Check if the file is a PDF (by checking the file extension)
            if file_path.lower().endswith('.pdf'):
                po_number, via = get_po(file_path)
                # Process the PDF file using your script logic
                invoice_number, tracking_number, serial_numbers, so_number = invoice_text_extractor(file_path)
                series = pandas.Series({
                    'Invoice #': invoice_number, 'ship via': via, 'Tracking #': tracking_number, 
                    'S.O. #': so_number, 'P.O. #': po_number, 'serial #': serial_numbers, 
                    'folderpath': folder_path, 'file_name': file_name
                })
                new_df = pandas.concat([new_df, series.to_frame().T])

    # Save the data to Excel
    new_df.to_excel(f'data/invoice_data_{current_date}.xlsx', engine='xlsxwriter', index=False)

if __name__ == "__main__":
    main()

