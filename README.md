
# RInvoice-Data-Extraction

This project is designed to extract specific data from PDF invoices in a designated folder and output that data into an Excel file. The extracted data can be further used for various purposes such as feeding into invoice entry tool. Usecases include using this file for entry into quicks and other ERP systems of chocie.

## Features
- Extracts essential data from PDF invoices (e.g., invoice number, tracking number, PO number, serial numbers, etc.).
- Uses template matching via OpenCV to find and click on grayscale elements (optional).
- Uses regex patterns to extract specific text from the PDF files.
- Generates an Excel file with the extracted data.
- Automatically handles files within a specified folder.
- Supports dynamic folder paths using environment variables for flexibility.

## Prerequisites

Before you begin, ensure you have the following installed on your system:

- **Python 3.x**
- **Pandas** (For data manipulation)
- **XlsxWriter** (For writing data to Excel files)
- **python-dotenv** (For handling environment variables)
- **OpenCV (cv2)** (For image processing)
- **Tabula** (For extracting tables from PDFs)
  
You can install these packages by running:
```bash
pip install pandas xlsxwriter python-dotenv opencv-python-headless tabula-py
```

## Setup

1. **Clone the repository**:
   ```bash
   git clone https://github.com/yourusername/RInvoice-Data-Extraction.git
   cd RInvoice-Data-Extraction
   ```

2. **Create a `.env` file**:
   In the root of the project, create a `.env` file and add the path to the folder containing your PDF invoices. This file should contain:
   ```env
   PDF_FOLDER_PATH="C:/path/to/your/invoice/folder"
   ```

   Make sure to replace `"C:/path/to/your/invoice/folder"` with the actual folder path where your PDF files are stored.

3. **Add your PDF extraction logic**:
   The script uses custom functions `invoice_text_extractor` and `get_po`. Ensure that these functions are implemented and properly extract data from your PDFs. These functions are expected to return values like `invoice_number`, `tracking_number`, `po_number`, etc.

   Update the po_so.json file for your tables on your invoice pages, you can do this using the tabula java app. I'd reccomend using the chocolatey package for Tabula.

## Functions Overview

### `invoice_text_extractor(filepath)`

This function processes the PDF file to extract textual information using regex patterns:
- **Invoice Number**: Extracted using `Invoice #\s*:\s*(\d+)`.
- **Tracking Number**: Extracted using `Tracking #\s*([A-Za-z0-9]+)`.
- **Serial Numbers**: Extracts serial numbers using various regex patterns.
- **Sales Order (S.O.) Number**: Extracted using `S\.O\. #\s*:\s*([\d\/]+)`.

It returns the extracted `invoice_number`, `tracking_number`, `serial_numbers`, and `so_number`.

### `get_po(filepath)`

This function uses the `tabula` library to extract Purchase Order (P.O.) numbers and shipment method from the PDF, using a pre-defined template (`po_so.json`). It attempts to find and extract the following:
- **P.O. Number**
- **Ship Via**

Returns both the extracted `po_number` and `ship_via`.

### `find_and_click_grayscale(filename)`

This function uses OpenCV to capture the screen, convert it to grayscale, and perform template matching to find elements based on a grayscale image. If the element is found, it automatically clicks on the center of the matched element.

## Running the Script

Once the environment is set up, you can run the script to process the invoices and generate the Excel file.

1. Run the script:
   ```bash
   python main.py
   ```

2. The script will:
   - Read all the PDF files in the folder specified by the `PDF_FOLDER_PATH`.
   - Extract relevant data from the PDFs using regex and `tabula`.
   - Save the extracted data into an Excel file named `invoice_data_<date>.xlsx` in the `data/` directory.

### Example Output

The output Excel file will contain the following columns for each processed PDF:

- **Invoice #**
- **ship via**
- **Tracking #**
- **S.O. #**
- **P.O. #**
- **serial #**
- **folderpath**
- **file_name**

## Error Handling

The script includes basic error handling:
- It checks if the `PDF_FOLDER_PATH` environment variable is set correctly.
- It verifies if the folder exists before proceeding with file extraction.
- If no PDFs are found in the folder, the script will print the file paths and continue processing.

## Customization

- If you have specific extraction logic for your PDFs, ensure that your `invoice_text_extractor` and `get_po` functions are customized to meet your needs.
- You can modify the columns or output format of the Excel file by editing the `pandas.Series()` in the `main()` function.
