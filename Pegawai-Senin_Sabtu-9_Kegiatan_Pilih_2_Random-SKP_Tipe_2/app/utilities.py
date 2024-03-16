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
    """
    Utility class for various helper functions.

    This class provides various helper functions for common tasks such as time conversion, checking date ranges, formatting durations, checking holidays, generating random times, retrieving data from a spreadsheet, sending emails, and parsing data into a readable format.

    Attributes:
        now (datetime): The current date and time.
        tz (pytz.timezone): The timezone used for date and time calculations.
        date (date): The current date.
        time (str): The current time in the format 'HH:MM:SS'.
        weekday (int): The current weekday (0-6, where 0 is Monday and 6 is Sunday).

    Methods:
        time_jakarta_to_utc(time: str) -> str:
            Converts a given time in Jakarta timezone to UTC timezone.

        cek_rentang_tanggal(mulai: str, selesai: str) -> bool:
            Checks if the current date is within the specified date range.

        format_duration(seconds: int) -> str:
            Formats a duration in seconds into the format 'HH:MM:SS'.

        list_holiday() -> list:
            Returns a list of holidays in Indonesia.

        is_holiday() -> bool:
            Checks if the current date is a holiday.

        random_time(time: str, range: int = 5) -> int:
            Generates a random time by adding a random number of minutes to the given time.

        get_jurnal() -> dict:
            Retrieves data from a spreadsheet and returns it as a dictionary.

        send_email(subject: str, body: str):
            Sends an email to a specified receiver.

        parse_data_to_pretty_output(json_data: dict, value: str) -> str:
            Parses data from the journal into a formatted description.

    Example:
        util = Util()
        utc_time = util.time_jakarta_to_utc('09:30')
        is_within_range = util.cek_rentang_tanggal('2022-01-01', '2022-12-31')
        duration = util.format_duration(3600)
        holidays = util.list_holiday()
        is_holiday = util.is_holiday()
        random_time = util.random_time('09:30')
        jurnal_data = util.get_jurnal()
        util.send_email('Subject', 'Body')
        pretty_output = util.parse_data_to_pretty_output(jurnal_data, 'senin-kamis')
    """
    def __init__(self) -> None:
        self.now = datetime.now()
        self.tz = pytz.timezone('Asia/Jakarta')
        self.date = datetime.now(self.tz).date()
        self.time = datetime.now(self.tz).time().strftime("%H:%M:%S")
        self.weekday = datetime.today().weekday()

    def time_jakarta_to_utc(self, time: str):
        """
        Converts a given time in Jakarta timezone to UTC timezone.

        Parameters:
            time (str): The time to be converted in the format 'HH:MM'.

        Returns:
            str: The converted time in UTC timezone in the format 'HH:MM'.

        Raises:
            ValueError: If the input time is not in the correct format.

        Example:
            >>> util = Util()
            >>> util.time_jakarta_to_utc('09:30')
            '02:30'
            >>> util.time_jakarta_to_utc('24:00')
            '19:00'
            >>> util.time_jakarta_to_utc('abc')
            'Format waktu tidak valid'
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
        """
        Retrieves a list of holidays in Indonesia.

        This method iterates through each date in the specified year and checks if it is a Sunday. If a date is a Sunday, it is considered a weekend holiday. The method then adds these weekend holidays to a dictionary object, where the keys are the holiday dates and the values are set to "Weekend Holiday". The dictionary is created using the 'holidays' module, specifically the 'CountryHoliday' class with the country code 'ID' for Indonesia.

        Returns:
            list: A list of holidays in Indonesia, including both official holidays and weekend holidays.

        Example:
            >>> util = Util()
            >>> holidays = util.list_holiday()
            >>> print(holidays)
            {
                date(2022, 1, 2): "Weekend Holiday",
                date(2022, 1, 9): "Weekend Holiday",
                date(2022, 1, 16): "Weekend Holiday",
                ...
            }
        """
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
        """
        Checks if the current date is a holiday.

        Returns:
            bool: True if the current date is a holiday, False otherwise.

        Example:
            >>> util = Util()
            >>> util.is_holiday()
            True
            >>> util.is_holiday()
            False
        """
        if datetime.now(self.tz).date() in self.list_holiday():
            return True
        return False

    def random_time(self, time, range: int = 5) -> int:
        """
        Generates a random time by adding a random number of minutes to the given time.

        Parameters:
            time (str): The time in the format 'HH:MM' to which the random number of minutes will be added.
            range (int, optional): The range of random minutes to be added. Defaults to 5.

        Returns:
            str: The generated random time in the format 'HH:MM'.

        Example:
            >>> util = Util()
            >>> util.random_time('09:30')
            '09:35'
            >>> util.random_time('12:00', range=10)
            '12:07'
            >>> util.random_time('24:00')
            '24:05'
        """
        return str("{0:02d}".format(int(time) + randint(1, range)))

    def get_jurnal(self) -> dict:
        """
        Retrieves data from a spreadsheet and returns it as a dictionary.

        Returns:
            dict: A dictionary containing the retrieved data from the spreadsheet.

        Example:
            >>> util = Util()
            >>> jurnal_data = util.get_jurnal()
            >>> print(jurnal_data)
            {
                "senin-kamis": [
                    {
                        "kegiatan": "Kegiatan 1",
                        "jam_mulai": "09:00",
                        "menit_mulai": "00",
                        "jam_selesai": "17:00",
                        "menit_selesai": "30",
                        "skp": "SKP 1",
                        "skp_value": "SKP Value 1",
                        "jumlah_diselesaikan": 5
                    },
                    {
                        "kegiatan": "Kegiatan 2",
                        "jam_mulai": "10:00",
                        "menit_mulai": "30",
                        "jam_selesai": "18:00",
                        "menit_selesai": "00",
                        "skp": "SKP 2",
                        "skp_value": "SKP Value 2",
                        "jumlah_diselesaikan": 3
                    }
                ],
                "jumat-sabtu": [
                    {
                        "kegiatan": "Kegiatan 3",
                        "jam_mulai": "08:00",
                        "menit_mulai": "15",
                        "jam_selesai": "16:30",
                        "menit_selesai": "45",
                        "skp": "SKP 3",
                        "skp_value": "SKP Value 3",
                        "jumlah_diselesaikan": 2
                    },
                    {
                        "kegiatan": "Kegiatan 4",
                        "jam_mulai": "09:30",
                        "menit_mulai": "45",
                        "jam_selesai": "17:45",
                        "menit_selesai": "15",
                        "skp": "SKP 4",
                        "skp_value": "SKP Value 4",
                        "jumlah_diselesaikan": 4
                    }
                ]
            }
        """

        value = get_sheet_table_values()
        return {
            "senin-sabtu": [
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
                },
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
                },
                {
                    "kegiatan": value[9][1],
                    "jam_mulai": get_sheet_time(value[9][2]),
                    "menit_mulai": get_sheet_time(value[9][3]),
                    "jam_selesai": get_sheet_time(value[9][4]),
                    "menit_selesai": get_sheet_time(value[9][5]),
                    "skp": get_skp_value(value[9][6])[0],
                    "skp_value": get_skp_value(value[9][6])[1],
                    "jumlah_diselesaikan": int(value[9][7])
                }
            ]
        }

    def send_email(self, subject:str, body:str):
        """
        Sends an email to a specified receiver.

        Parameters:
            subject (str): The subject of the email.
            body (str): The body of the email.

        Raises:
            Exception: If there is an error while sending the email.

        Example:
            >>> util = Util()
            >>> util.send_email('Subject', 'Body')
        """
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
        """
        Parses data from the journal into a formatted description.

        Parameters:
            json_data (dict): A dictionary containing the journal data.
            value (str): The key in the dictionary to retrieve the data from.

        Returns:
            str: The formatted description of the journal data.

        Example:
            >>> util = Util()
            >>> jurnal_data = util.get_jurnal()
            >>> pretty_output = util.parse_data_to_pretty_output(jurnal_data, 'senin-kamis')
            >>> print(pretty_output)
            1. Kegiatan 1 dimulai dari jam 09:00:00 hingga 17:00:30
            2. Kegiatan 2 dimulai dari jam 10:00:30 hingga 18:00:00
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