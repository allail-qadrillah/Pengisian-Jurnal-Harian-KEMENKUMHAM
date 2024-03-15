from google.oauth2.service_account import Credentials
from datetime import datetime
from dotenv import load_dotenv
import gspread
import os

load_dotenv()

scopes = [
    "https://www.googleapis.com/auth/spreadsheets"
]
creds = Credentials.from_service_account_file(
    "sheet_cred.json", scopes=scopes)
client = gspread.authorize(creds)
worksheet = client.open_by_key(os.getenv('sheet_id')).sheet1

def get_sheet_row_col(row=int, col=int):
    """mendapatkan data dari spredsheet dengan memperolehnya dari nilai baris dan kolom
    """
    return worksheet.cell(row, col).value

def get_sheet_table_values() -> list:
    """medapatkan semua data dari tabel dalam bantuk list"""
    return worksheet.get_all_values()[:9]

def get_sheet_time(value) -> str:
    """mendapatkan nilai waktu jam dan menit
    - jika nilainya adalah integer dan dibawah 10, tambahkan 0 dibelakangnya. 
    karena nilai yang diperlukan oleh jurnal adalah seperti 01-09"""

    try:
      value = int(value)  # jika value adalah integer
      if isinstance(value, int) and 0 <= value <= 9:
        value = f"0{value}"
    except ValueError:  # jika value adalah string
      pass

    return str(value)


def get_skp_value(skp):
    """mendapatkan nilai value SKP dari spreadsheet
    return: index, value
    ex: get_skp_value(2, 7)[0] or get_skp_value(2, 7)[1]
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
