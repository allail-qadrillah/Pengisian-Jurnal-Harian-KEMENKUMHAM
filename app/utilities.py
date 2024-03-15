from datetime import date, timedelta, datetime
from random import randint
from dotenv import load_dotenv
import logging
import smtplib
import holidays
import pytz
import os

from .spreadsheet import get_sheet_row_col, get_sheet_time, get_skp_value, get_sheet_table_values

load_dotenv()
class Util:
    def __init__(self) -> None:
        self.now = datetime.now()
        self.tz = pytz.timezone('Asia/Jakarta')
        self.date = datetime.now(self.tz).date()
        self.time = datetime.now(self.tz).time().strftime("%H:%M:%S")
        self.weekday = datetime.today().weekday()

    def write_log(self, write):
        """
        menulis pada db/log.txt
        """
        file = open('./app/db/log.txt', 'a')
        file.write(
            f"{datetime.now(self.tz).strftime('%Y-%m-%d %H:%M:%S')} | {write} \n")
        file.close()

    def time_jakarta_to_utc(self, time: str):
        """ 
        Convert waktu jakarta time to UTC
        params time: %H:%M
        """
        try:
            # Parsing waktu dalam format jam:menit
            jam, menit = map(int, time.split(':'))

            # Waktu dalam zona waktu Jakarta
            jakarta_time = self.tz.localize(
                datetime(self.date.year, self.date.month, self.date.day, jam, menit, 0))
            # Konversi ke UTC
            utc_time = jakarta_time.astimezone(pytz.timezone('UTC'))
            return utc_time.strftime('%H:%M')
        except ValueError:
            return "Format waktu tidak valid"

    def cek_rentang_tanggal(self, mulai: str, selesai: str) -> bool:
        """
        mengembalikan true ketika hari ini berada pada rentang yang ditentukan
        params:  "%Y-%m-%d"
        """
        mulai = datetime.strptime(mulai, "%Y-%m-%d").date()
        selesai = datetime.strptime(selesai, "%Y-%m-%d").date()
        if mulai <= self.date <= selesai:
            return True
        return False

    def format_duration(self, seconds):
        """get dengan format jam:menit:detik dari nilai detik"""
        hours, remainder = divmod(seconds, 3600)
        minutes, seconds = divmod(remainder, 60)
        return f"{int(hours):02d}:{int(minutes):02d}:{int(seconds):02d}"

    def list_holiday(self) -> list:
        id_holidays = holidays.CountryHoliday('ID')
        # Mendapatkan daftar hari libur akhir pekan
        weekend_holidays = []

        # Iterasi melalui setiap tanggal dalam rentang tahun
        start_date = date(self.date.year, 1, 1)
        end_date = date(self.date.year, 12, 31)

        while start_date <= end_date:
            # Periksa apakah tanggal adalah hari Minggu
            if start_date.weekday() == 6:  # 6 mewakili hari Minggu
                weekend_holidays.append(start_date)

            start_date += timedelta(days=1)

        # Tambahkan hari libur akhir pekan ke dalam objek holiday
        for holiday in weekend_holidays:
            id_holidays[holiday] = "Weekend Holiday"
        return id_holidays

    def is_holiday(self) -> bool:
        """apakah hari ini libur?"""
        if datetime.now(self.tz).date() in self.list_holiday():
            return True
        return False

    def is_senin_kamis(self) -> bool:
        """apakah hari ini rentang senin - kamis"""
        if 0 <= datetime.now(self.tz).weekday() <= 3:
            return True
        return False

    def is_jumat_sabtu(self) -> bool:
        """apakah hari ini rentang jumat - sabtu"""
        if 4 <= datetime.now(self.tz).weekday() <= 5:
            return True
        return False

    def random_time(self, time, range: int = 5) -> int:
        """menjumlahkan waktu secara random"""
        return str("{0:02d}".format(int(time) + randint(1, range)))

    def get_jurnal(self) -> dict:
        """mendapatkan data json dari jurnal"""

        value = get_sheet_table_values()
        return {
            "senin-kamis": [
                {
                    "kegiatan": value[1][1],
                    "jam_mulai": get_sheet_time(value[1][2]),
                    "menit_mulai": get_sheet_time(value[1][3]),
                    "jam_selesai": get_sheet_time(value[1][4]),
                    "menit_selesai": get_sheet_time(value[1][5]),
                    "skp": get_skp_value(value[1][6])[0],
                    "skp_value": get_skp_value(value[1][6])[1],
                    "jumlah_diselesaikan": int(value[1][7])
                },
                {
                    "kegiatan": value[2][1],
                    "jam_mulai": get_sheet_time(value[2][2]),
                    "menit_mulai": get_sheet_time(value[2][3]),
                    "jam_selesai": get_sheet_time(value[2][4]),
                    "menit_selesai": get_sheet_time(value[2][5]),
                    "skp": get_skp_value(value[2][6])[0],
                    "skp_value": get_skp_value(value[2][6])[1],
                    "jumlah_diselesaikan": int(value[2][7])
                },
                {
                    "kegiatan": value[3][1],
                    "jam_mulai": get_sheet_time(value[3][2]),
                    "menit_mulai": get_sheet_time(value[3][3]),
                    "jam_selesai": get_sheet_time(value[3][4]),
                    "menit_selesai": get_sheet_time(value[3][5]),
                    "skp": get_skp_value(value[3][6])[0],
                    "skp_value": get_skp_value(value[3][6])[1],
                    "jumlah_diselesaikan": int(value[3][7])
                },
                {
                    "kegiatan": value[4][1],
                    "jam_mulai": get_sheet_time(value[4][2]),
                    "menit_mulai": get_sheet_time(value[4][3]),
                    "jam_selesai": get_sheet_time(value[4][4]),
                    "menit_selesai": get_sheet_time(value[4][5]),
                    "skp": get_skp_value(value[4][6])[0],
                    "skp_value": get_skp_value(value[4][6])[1],
                    "jumlah_diselesaikan": int(value[4][7])
                }
            ],
            "jumat-sabtu": [
                {
                    "kegiatan": value[5][1],
                    "jam_mulai": get_sheet_time(value[5][2]),
                    "menit_mulai": get_sheet_time(value[5][3]),
                    "jam_selesai": get_sheet_time(value[5][4]),
                    "menit_selesai": get_sheet_time(value[5][5]),
                    "skp": get_skp_value(value[5][6])[0],
                    "skp_value": get_skp_value(value[5][6])[1],
                    "jumlah_diselesaikan": int(value[5][7])
                },
                {
                    "kegiatan": value[6][1],
                    "jam_mulai": get_sheet_time(value[6][2]),
                    "menit_mulai": get_sheet_time(value[6][3]),
                    "jam_selesai": get_sheet_time(value[6][4]),
                    "menit_selesai": get_sheet_time(value[6][5]),
                    "skp": get_skp_value(value[6][6])[0],
                    "skp_value": get_skp_value(value[6][6])[1],
                    "jumlah_diselesaikan": int(value[6][7])
                },
                {
                    "kegiatan": value[7][1],
                    "jam_mulai": get_sheet_time(value[7][2]),
                    "menit_mulai": get_sheet_time(value[7][3]),
                    "jam_selesai": get_sheet_time(value[7][4]),
                    "menit_selesai": get_sheet_time(value[7][5]),
                    "skp": get_skp_value(value[7][6])[0],
                    "skp_value": get_skp_value(value[7][6])[1],
                    "jumlah_diselesaikan": int(value[7][7])
                },
                {
                    "kegiatan": value[8][1],
                    "jam_mulai": get_sheet_time(value[8][2]),
                    "menit_mulai": get_sheet_time(value[8][3]),
                    "jam_selesai": get_sheet_time(value[8][4]),
                    "menit_selesai": get_sheet_time(value[8][5]),
                    "skp": get_skp_value(value[8][6])[0],
                    "skp_value": get_skp_value(value[8][6])[1],
                    "jumlah_diselesaikan": int(value[8][7])
                }
            ]
        }

    def send_email(self, subject:str, body:str):
        """Sending email to receiver"""
        try:
            RECEIVER_EMAIL = get_sheet_row_col(5, 11)

            if RECEIVER_EMAIL is not None:
                logging.info(f"Sending email to {RECEIVER_EMAIL}")
                server_ssl = smtplib.SMTP_SSL('smtp.gmail.com', 465)
                server_ssl.ehlo()
                server_ssl.login(os.getenv('email_sender'), os.getenv('email_password'))
                server_ssl.sendmail(os.getenv('email_sender'), RECEIVER_EMAIL, 
                                    f"Subject: {subject}\n\n{body}")
                server_ssl.close()

        except Exception as e:
            print('Error:', e)

    def parse_data_to_pretty_output(self, json_data:dict, value:str):
        """merapikan data json dari jurnal menjadi deskripsi yang rapi
        ex parse_data_to_pretty_output(get_jurnal(), "senin-kamis")
        """

        # Inisialisasi list untuk penjelasan
        descriptions = []

        # Iterasi melalui setiap objek dan tambahkan penjelasan ke list
        for i, obj in enumerate(json_data[value], start=1):
            kegiatan = obj["kegiatan"]
            jam_mulai = obj["jam_mulai"]
            menit_mulai = obj["menit_mulai"]
            jam_selesai = obj["jam_selesai"]
            menit_selesai = obj["menit_selesai"]
            description = f"{i}. {kegiatan} dimulai dari jam {jam_mulai}:{menit_mulai} hingga {jam_selesai}:{menit_selesai}"
            descriptions.append(description)

        # Menggabungkan semua penjelasan dalam satu string
        formatted_descriptions = "\n".join(descriptions)

        return formatted_descriptions