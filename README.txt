DocxToSmartsheet
================

This Python script processes DOCX files, extracts data, and updates a Smartsheet using the Smartsheet API with a Tkinter GUI.

Dependencies:
-------------
- Python 3.x
- pyautogui
- smartsheet
- tkinter (pre-installed)
- python-docx

Install with:
pip install pyautogui smartsheet-python-sdk python-docx

Configuration:
--------------
1. Set your Smartsheet API key in `config.py` (`API_KEY`).
API_URL = "https://api.smartsheet.com/2.0/sheets"
API_KEY = "api key"
2. Create a `smartsheet_data.py` file with columns names and URLs:
column_name = 
[
    "column_name_1",
    "column_name_2"
]
3. Create a `vendor_data.json` file with vendor names and URLs:
{
  "Vendor A": "unical part of Smartsheet URL for Vendor A",
  "Vendor B": "unical part of Smartsheet URL for Vendor B"
}


Logging:
--------
Logs are saved in `rwsheet.log`.

Functions:
----------
- `get_d_code_en_num_from_docx(file_path)`: Extracts D code and Enhancement number from DOCX.
- `get_count_import_docx_data_to_smartsheet(file_path)`: Processes DOCX data and updates Smartsheet.
- `searching_row_id(sheet_id, d_code_column_name, d_code, en_num_column_name, e_number)`: Finds a row ID.
- `update_single_cell_in_smartsheet(sheet_id, row_id, column_id, new_value)`: Updates a cell in Smartsheet.

GUI:
----
1. **Select Vendor**: Choose a vendor.
2. **Upload File**: Select a DOCX file.
3. **Start Process**: Process the DOCX and update Smartsheet.
4. **Connection Test**: Test the connection to Smartsheet.

Usage:
------
1. Run the script to open the GUI.
2. Select a vendor and upload a DOCX file.
3. Click "Start" to process and update Smartsheet.
4. Use "Connection Test" to verify the connection.

License:
--------
MIT License