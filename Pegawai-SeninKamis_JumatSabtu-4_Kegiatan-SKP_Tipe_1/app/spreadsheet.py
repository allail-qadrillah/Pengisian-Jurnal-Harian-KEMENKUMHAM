from google.oauth2.service_account import Credentials
from datetime import datetime
from dotenv import load_dotenv
import gspread
import os

load_dotenv()

# The scope of the Google Sheets API that allows us to read and write data.
scopes = [
    "https://www.googleapis.com/auth/spreadsheets"
]
# The credentials to access the Google Sheets API.
creds = Credentials.from_service_account_file(
    "sheet_cred.json", scopes=scopes)
client = gspread.authorize(creds)
worksheet = client.open_by_key(os.getenv('sheet_id')).sheet1

def get_sheet_row_col(row=int, col=int):
  """
  Get the value of a cell in a Google Sheet.

  Parameters:
  - row (int): The row number of the cell.
  - col (int): The column number of the cell.

  Returns:
  - str: The value of the cell.

  Example:
  get_sheet_row_col(1, 1)  # Returns the value of the cell in the first row and first column.
  """
  return worksheet.cell(row, col).value

def get_sheet_table_values() -> list:
    """medapatkan semua data dari tabel dalam bantuk list"""
    return worksheet.get_all_values()[:9]

def get_sheet_time(value) -> str:
    """
    Converts a value to a string representation of time.

    Parameters:
    - value (int or str): The value to be converted. If an integer between 0 and 9 (inclusive), it will be padded with a leading zero.

    Returns:
    - str: The string representation of the value.

    Example:
    >>> get_sheet_time(5)
    '05'
    >>> get_sheet_time('10')
    '10'
    """
    try:
      value = int(value)  # jika value adalah integer
      if isinstance(value, int) and 0 <= value <= 9:
        value = f"0{value}"
    except ValueError:  # jika value adalah string
      pass

    return str(value)


def get_skp_value(skp):
    """
    This function takes a string parameter 'skp' and returns a tuple containing an integer value and a string value.

    Parameters:
    - skp (str): The input string representing the SKP value.

    Returns:
    - tuple: A tuple containing an integer value and a string value. The integer value represents the SKP category, while the string value represents the SKP code.

    SKP Categories:
    0: Membantu pelaksanaan tugas dan pendalaman materi bidang pengamanan
    1: Membantu pelaksanaan tugas dalam pendalaman materi bidang pelayanan tahanan
    2: Membantu pelaksanaan tugas dalam pendalaman materi bidang pengelolaan
    3: Melaksanakan tugas lainnya yang diperintahkan oleh pimpinan
    4: Lain-Lain
    5: Tugas Tambahan
    6: Kreatifitas

    SKP Codes:
    - For categories 0 to 3: A string value in the format '{current_year}999967720178403_{category_number}'.
    - For category 4: The string value is 'lainlain'.
    - For category 5: The string value is '2'.
    - For category 6: The string value is '3'.
    """

    if skp == "Membantu pelaksanaan tugas dan pendalaman materi bidang pengamanan":
      return 0, f"{datetime.now().year}999967720178403_01"
    elif skp == "Membantu pelaksanaan tugas dalam pendalaman materi bidang pelayanan tahanan":
      return 1, f"{datetime.now().year}999967720178403_02"
    elif skp == "Membantu pelaksanaan tugas dalam pendalaman materi bidang pengelolaan":
      return 2, f"{datetime.now().year}999967720178403_03"
    elif skp == "Melaksanakan tugas lainnya yang diperintahkan oleh pimpinan":
      return 3, f"{datetime.now().year}999967720178403_04"
    elif skp == "Lain-Lain":
      return 4, "lainlain"
    elif skp == "Tugas Tambahan":
      return 5, "2"
    elif skp == "Kreatifitas":
      return 6, "3"
