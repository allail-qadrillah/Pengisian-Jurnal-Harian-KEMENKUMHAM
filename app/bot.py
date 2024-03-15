from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, TimeoutException, UnexpectedAlertPresentException
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from dotenv import load_dotenv
from tempfile import mkdtemp
from .utilities import Util
from time import sleep
import logging
import os

load_dotenv()
botlog = logging.getLogger(__name__)
botlog.setLevel(logging.INFO)

class BOT(Util):
    """
    This code snippet defines a class called BOT that is used for automating tasks on a website using Selenium. 
    The class has methods for logging in, filling out a daily journal, checking if the journal has been filled, 
    and starting the automation process. The class uses the Chrome webdriver and supports both local and remote execution. 
    The code also imports necessary modules and defines some utility functions.
    """
    def __init__(self, server='lambda'):
        """Initialize the BOT class.

        Parameters:
        - server (str) {local/lambda}: The server to be used for execution. Default is 'lambda'.

        Attributes:
        - username (str): The username for login.
        - password (str): The password for login.
        - is_complete_fill (bool): Flag to indicate if the journal filling process is complete. Default is False.
        - exception_occured (bool): Flag to indicate if a timeout occurred during the journal filling process. Default is False.
        - server (str): The server to be used for execution.

        Returns:
        - None
        """
        super().__init__()
        self.username = os.getenv('nip')
        self.password = os.getenv('password')
        self.is_complete_fill = False
        self.exception_occured = False
        self.is_login = False
        self.server = server

        if self.server == 'lambda':
            options = webdriver.ChromeOptions()
            service = webdriver.ChromeService("/opt/chromedriver")

            options.binary_location = '/opt/chrome/chrome'
            options.add_argument("--headless=new")
            options.add_argument('--no-sandbox')
            options.add_argument("--disable-gpu")
            options.add_argument("--window-size=1280x1696")
            options.add_argument("--single-process")
            options.add_argument("--disable-dev-shm-usage")
            options.add_argument("--disable-dev-tools")
            options.add_argument("--no-zygote")
            options.add_argument("--remote-debugging-port=9222")
            options.add_argument("--incognito")

            self.driver = webdriver.Chrome(options=options, service=service)

        elif self.server == 'local':
            from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
            self.driver = webdriver.Remote(
                            command_executor='http://localhost:4444/wd/hub',
                            desired_capabilities=DesiredCapabilities.CHROME,
                            options=webdriver.ChromeOptions()
                        )
            
    def get(self, url):
        """Navigate to the specified URL.

        Parameters:
        - url (str): The URL to navigate to.
        """
        self.driver.get(url)

    def wait_element_clear(self, XPATH, time=30):
        """Wait for an element to be clickable and then clear its value.

        This method waits for an element specified by the given XPath to be clickable within the specified time limit. 
        Once the element is clickable, it clears its current value.

        Parameters:
        - XPATH (str): The XPath of the element to wait for and clear.
        - time (int): The maximum time in seconds to wait for the element to be clickable. Default is 30 seconds.

        Returns:
        - None

        Raises:
        - TimeoutException: If the element is not clickable within the specified time limit.
        """
        WebDriverWait(self.driver, time).until(EC.element_to_be_clickable(
            (By.XPATH, XPATH)))
        self.driver.find_element(By.XPATH, XPATH).clear()

    def wait_element_get(self, XPATH, time=30):
        """Wait for an element to be clickable and return it.

        This method waits for an element specified by the given XPath to be clickable within the specified time limit. 
        Once the element is clickable, it returns the element.

        Parameters:
        - XPATH (str): The XPath of the element to wait for and return.
        - time (int): The maximum time in seconds to wait for the element to be clickable. Default is 30 seconds.

        Returns:
        - WebElement: The clickable element specified by the given XPath.
        """
        WebDriverWait(self.driver, time).until(EC.element_to_be_clickable(
            (By.XPATH, XPATH)))
        return self.driver.find_element(By.XPATH, XPATH)

    def wait_element_click(self, XPATH, time=60):
        WebDriverWait(self.driver, time).until(EC.element_to_be_clickable(
            (By.XPATH, XPATH)))
        self.driver.find_element(By.XPATH, XPATH).click()

    def wait_element_input(self, input, XPATH, time=30):
        """Wait for an element to be clickable and click it.

        This method waits for an element specified by the given XPath to be clickable within the specified time limit. 
        Once the element is clickable, it clicks the element.

        Parameters:
        - XPATH (str): The XPath of the element to wait for and click.
        - time (int): The maximum time in seconds to wait for the element to be clickable. Default is 60 seconds.
        """
        WebDriverWait(self.driver, time).until(EC.element_to_be_clickable(
            (By.XPATH, XPATH)))
        self.driver.find_element(By.XPATH, XPATH).send_keys(input)\

    def wait_element_select_value(self, value: str, XPATH, time=30):
        """Wait for an element to be clickable and select an option by its value.

        This method waits for an element specified by the given XPath to be clickable within the specified time limit. 
        Once the element is clickable, it selects the option with the specified value from the dropdown menu.

        Parameters:
        - value (str): The value of the option to be selected.
        - XPATH (str): The XPath of the element to wait for and select.
        - time (int): The maximum time in seconds to wait for the element to be clickable. Default is 30 seconds.
        """
        WebDriverWait(self.driver, time).until(EC.element_to_be_clickable(
            (By.XPATH, XPATH)))
        Select(self.driver.find_element(By.XPATH, XPATH)).select_by_value(value)

    def wait_element_select_index(self, index: int, XPATH, time=30):
        """Wait for an element to be clickable and select an option by its index.

        This method waits for an element specified by the given XPath to be clickable within the specified time limit. 
        Once the element is clickable, it selects the option with the specified index from the dropdown menu.

        Parameters:
        - index (int): The index of the option to be selected.
        - XPATH (str): The XPath of the element to wait for and select.
        - time (int): The maximum time in seconds to wait for the element to be clickable. Default is 30 seconds.
        """
        WebDriverWait(self.driver, time).until(EC.element_to_be_clickable(
            (By.XPATH, XPATH)))
        Select(self.driver.find_element(By.XPATH, XPATH)).select_by_index(index)
    
    def close(self):
        """Close Driver"""
        return self.driver.quit()
    
    def login(self):
        """Login to SIMPEG KEMENKUMHAM

        This method is used to login to the SIMPEG website. It navigates to the login page, 
        fills in the username and password fields, and clicks the login button. After successful login, 
        it waits for 3 seconds. and if an error occurs during the login process, it raises an exception and send email.

        """
        botlog.info("Login ...")
        
        try:
            self.get('https://simpeg.kemenkumham.go.id/devp/siap/signin.php')
            # USERNAME FILL FORM
            self.wait_element_input(input=self.username, XPATH="/html/body/div[1]/div/div/div/div/div/div/div[2]/input[1]")
            # USERNAME CLICK FORM
            self.wait_element_click(XPATH="/html/body/div[1]/div/div/div/div/div/div/div[2]/input[2]")
            # PASSWORD FILL FORM
            self.wait_element_input(input=self.password, XPATH="/html/body/div[2]/div[2]/form/input[7]")
            # PASSWORD CLICK FORM
            self.wait_element_click(XPATH="/html/body/div[2]/div[3]/button[1]")

            botlog.info("Login Done")
            sleep(3)
            return True
        
        except UnexpectedAlertPresentException as uape: 
            botlog.critical(f"Login Failed {repr(uape)}")
            self.exception_occured = True
            return False

        except Exception as e:
            botlog.critical(f"Login Failed {repr(e)}")
            self.send_email(subject=f"Pengisian Jurnal SIMPEG KEMENKUMHAM Tanggal {self.date} Gagal",
                            body=f"Salam. Semoga anda dalam keadaan baik, saya ingin memberitahu anda bahwa tidak dapat login ke SIMPEG KEMENKUMHAM. Terjadi kesalahan {repr(e)}.\n\nTerima kasih atas perhatiannya,\nSalam hormat.")
            self.exception_occured = True
            return False
            
    def fill_jurnal(self, jam_mulai:str, menit_mulai:str, jam_selesai:str, menit_selesai:str,
                    skp: int, skp_value: str, kegiatan: str, jumlah_diselesaikan: int):
        """This method is used to fill out the daily journal on the SIMPEG website. It takes in the following parameters:

        - jam_mulai (str): The starting hour of the activity.
        - menit_mulai (str): The starting minute of the activity.
        - jam_selesai (str): The ending hour of the activity.
        - menit_selesai (str): The ending minute of the activity.
        - skp (int): The SKP value for the activity.
        - skp_value (str): The SKP value for the activity.
        - kegiatan (str): The description of the activity.
        - jumlah_diselesaikan (int): The number of tasks completed for the activity.

        The method performs the following steps:
        1. Navigates to the daily journal page on the SIMPEG website.
        2. Clicks the "Tambah" button to add a new journal entry.
        3. Inputs the starting hour, starting minute, ending hour, and ending minute of the activity.
        4. Inputs the SKP value for the activity.
        5. Inputs the description of the activity.
        6. Inputs the number of tasks completed for the activity.
        7. Clicks the "Simpan" button to save the journal entry.

        If a TimeoutException occurs during the process, it raises an exception and sends an email notification.
        If any other exception occurs, it retries the process up to a maximum of 10 times.

        Returns:
        - None
        """
        max_retries = 15
        retries = 0
        botlog.info(f"Mengisi jurnal harian {kegiatan} | waktu mulai {jam_mulai}:{menit_mulai} & waktu selesai {jam_selesai}:{menit_selesai} | SKP {skp} dan jumlah diselesaikan {jumlah_diselesaikan}")
        # melakukan pengisian form dengan mencoba ulang ketika ada error.
        while retries <= max_retries:
            try:
                # OPEN WEB JURNAL HARIAN
                self.get('https://simpeg.kemenkumham.go.id/devp/siap/skp_journal.php')

                # CLICK BTN TAMBAH
                self.wait_element_click(XPATH="/html/body/div[3]/div[2]/a[1]")

                # INPUT JAM MULAI
                self.wait_element_select_value(value=jam_mulai, XPATH="/html/body/div[4]/div[2]/form/fieldset/div[1]/div/select[1]")
                # INPUT MENIT MULAI
                self.wait_element_select_value(value=menit_mulai, XPATH="/html/body/div[4]/div[2]/form/fieldset/div[1]/div/select[2]")
                # INPUT JAM SELESAI
                self.wait_element_select_value(value=jam_selesai, XPATH="/html/body/div[4]/div[2]/form/fieldset/div[1]/div/select[3]")
                # INPUT MENIT SELESAI
                self.wait_element_select_value(value=menit_selesai, XPATH="/html/body/div[4]/div[2]/form/fieldset/div[1]/div/select[4]")

                try:
                    # INPUT SKP VALUE 
                    self.wait_element_select_value(value=skp_value, XPATH="/html/body/div[4]/div[2]/form/fieldset/div[2]/div/select")
                except NoSuchElementException:
                    # INPUT SKP INDEX 
                    self.wait_element_select_index(index=skp, XPATH="/html/body/div[4]/div[2]/form/fieldset/div[2]/div/select")
            
                # INPUT KEGIATAN
                self.wait_element_input(input=kegiatan, XPATH="/html/body/div[4]/div[2]/form/fieldset/div[3]/div/textarea")
                # INPUT JUMLAH DISELESAIKAN
                self.wait_element_clear(XPATH="/html/body/div[4]/div[2]/form/fieldset/div[4]/div/input")
                self.wait_element_input(input=jumlah_diselesaikan, XPATH="/html/body/div[4]/div[2]/form/fieldset/div[4]/div/input")
                # KLIK BTN SIMPAN
                self.wait_element_click(XPATH="/html/body/div[4]/div[3]/button[2]")
            
                break
                # BTN BATAL
                # self.wait_element_click(XPATH="/html/body/div[4]/div[3]/button[1]")
            
            except TimeoutException: # ERROR Anda tidak terdaftar sebagai pegawai WFH!
                botlog.critical("TimeoutException: Situs pengisian jurnal tidak dapat diakses (Anda tidak terdaftar sebagai pegawai WFH)" )
                self.exception_occured = True
                break
            
            except Exception as e:
                retries += 1
                botlog.error(f"Terjadi kesalahan {repr(e)}. Percobaan ke-{retries} dari {max_retries}")

                if retries == max_retries:
                    botlog.info("Jumlah percobaan maksimum telah tercapai. Tidak dapat melanjutkan.")
                    self.send_email(subject=f"Pengisian Jurnal SIMPEG KEMENKUMHAM Tanggal {self.date}",
                                    body=f"Salam. Semoga anda dalam keadaan baik, saya ingin memberitahu anda bahwa jurnal harian untuk tanggal {self.date} Gagal di isi.\
                                    \n\nTerima kasih atas perhatiannya,\
                                    \nSalam hormat.")
                    break
                
    def is_has_filled(self) -> bool:
        """Check if the journal table has been filled.

        This method navigates to the journal page on the SIMPEG website and checks if the journal table has been filled. 
        It returns True if the table has at least one row, indicating that the journal has been filled. 
        Otherwise, it returns False.

        Returns:
        - bool: True if the journal table has been filled, False otherwise.
        """
        self.get('https://simpeg.kemenkumham.go.id/devp/siap/skp_journal.php')
        # DAPATKAN ELEMENT TABEL
        try:
          table = self.wait_element_get(XPATH="/html/body/div[3]/div[1]/div/table[1]",
                                       time=60)
          tbody = table.find_element(By.TAG_NAME, "tbody")
          rows  = tbody.find_elements(By.TAG_NAME, "tr")
          
          # Periksa jumlah baris
          if len(rows) > 0: return True
          else: return False
        except:
            pass
        
    def start(self):
        """Starts the process of filling out the daily journal on the SIMPEG website.

        This method performs the following steps:
        1. Retrieves the journal data.
        2. Checks if it is a holiday. If not, proceeds with the journal filling process.
        3. Logs in to the SIMPEG website.
        4. Checks the day of the week and fills out the journal accordingly.
        - If it is Monday to Thursday, fills out the journal for that day.
        - If it is Friday or Saturday, fills out the journal for that day.
        5. Sends an email notification with the details of the filled journal.
        6. Closes the driver.

        Returns:
        - None
        """
        botlog.info("================= TASK START =================")

        jurnal = self.get_jurnal()

        while not self.is_complete_fill:
            if self.is_complete_fill: break
            if self.exception_occured: break
            # JIKA TIDAK LIBUR LANJUT TASKS
            if not self.is_holiday():

                if self.login(): # JIKA LOGIN BERHASIL
                    # JIKA HARI INI BELUM TERISI?
                    botlog.info("MENGISI JURNAL HARIAN ...")
                    # APAKAH HARI INI SENIN - KAMIS?
                    if self.is_senin_kamis():
                        for item in jurnal.get("senin-kamis"):
                            # OPEN WEB JURNAL HARIAN
                            self.fill_jurnal(
                                                jam_mulai=item.get("jam_mulai"),
                                                menit_mulai=self.random_time(time=item.get("menit_mulai")),
                                                jam_selesai=item.get("jam_selesai"),
                                                menit_selesai=self.random_time(time=item.get("menit_selesai")),
                                                skp=item.get("skp"),
                                                skp_value=item.get("skp_value"),
                                                kegiatan=item.get("kegiatan"),
                                                jumlah_diselesaikan=item.get("jumlah_diselesaikan")
                                            )
                            if self.exception_occured == True:
                                self.send_email(subject=f"Pengisian Jurnal SIMPEG KEMENKUMHAM Tanggal {self.date} Gagal",
                                    body=f"Salam. Semoga anda dalam keadaan baik, saya ingin memberitahu anda bahwa jurnal harian untuk tanggal {self.date} gagal di isi. Situs pengisian jurnal tidak dapat diakses karena 'Anda tidak terdaftar sebagai pegawai WFH'.\n\nTerima kasih atas perhatiannya,\nSalam hormat.")
                                break
                            
                        if self.exception_occured == False:
                            self.is_complete_fill = True
                            self.send_email(subject=f"Pengisian Jurnal SIMPEG KEMENKUMHAM Tanggal {self.date}",
                                        body=f"Salam. Semoga anda dalam keadaan baik, saya ingin memberitahu anda bahwa jurnal harian untuk tanggal {self.date} telah berhasil di isi. Berikut adalah rincian kegiatan hari ini:\
                                        \n\n{self.parse_data_to_pretty_output(jurnal, 'senin-kamis')} \
                                        \n\nTerima kasih atas perhatiannya,\
                                        \nSalam hormat.")
                            botlog.info("FILL JURNAL SENIN KAMIS DONE")
        
                    # APAKAH HARI INI JUMAT - SABTU?
                    elif self.is_jumat_sabtu():
                        for item in jurnal.get("jumat-sabtu"):
                            # OPEN WEB JURNAL HARIAN
                            self.fill_jurnal(
                                            jam_mulai=item.get("jam_mulai"),
                                            menit_mulai=self.random_time(time=item.get("menit_mulai")),
                                            jam_selesai=item.get("jam_selesai"),
                                            menit_selesai=self.random_time(time=item.get("menit_selesai")),
                                            skp=item.get("skp"),
                                            skp_value=item.get("skp_value"),
                                            kegiatan=item.get("kegiatan"),
                                            jumlah_diselesaikan=item.get("jumlah_diselesaikan")
                                        )
                            if self.exception_occured == True:
                                self.send_email(subject=f"Pengisian Jurnal SIMPEG KEMENKUMHAM Tanggal {self.date} Gagal",
                                    body=f"Salam. Semoga anda dalam keadaan baik, saya ingin memberitahu anda bahwa jurnal harian untuk tanggal {self.date} gagal di isi. Situs pengisian jurnal tidak dapat diakses karena 'Anda tidak terdaftar sebagai pegawai WFH'.\n\nTerima kasih atas perhatiannya,\nSalam hormat.")
                                break

                        if self.exception_occured == False: 
                            self.is_complete_fill = True
                            self.send_email(subject=f"Pengisian Jurnal SIMPEG KEMENKUMHAM Tanggal {self.date}",
                                            body=f"Salam. Semoga anda dalam keadaan baik, saya ingin memberitahu anda bahwa jurnal harian untuk tanggal {self.date} telah berhasil di isi. Berikut adalah rincian kegiatan hari ini:\
                                        \n\n{self.parse_data_to_pretty_output(jurnal, 'jumat-sabtu')} \
                                        \n\nTerima kasih atas perhatiannya,\
                                        \nSalam hormat.")
                            botlog.info("FILL JURNAL JUMAT SABTU DONE")

                else: 
                    botlog.critical("Tidak dapat melanjutkan proses karena login gagal.")
                    break
            else:
                botlog.info("HARI INI LIBUR")
                break
            
        # self.driver.close()
        botlog.info("================= TASK DONE =================")
