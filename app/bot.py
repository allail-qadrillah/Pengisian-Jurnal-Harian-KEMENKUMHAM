from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, TimeoutException
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
    def __init__(self, server='lambda'):
        super().__init__()
        self.username = os.getenv('nip')
        self.password = os.getenv('password')
        self.is_complete_fill = False
        self.timeout_occured = False
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
        self.driver.get(url)

    def wait_element_clear(self, XPATH, time=30):
        WebDriverWait(self.driver, time).until(EC.element_to_be_clickable(
            (By.XPATH, XPATH)))
        self.driver.find_element(By.XPATH, XPATH).clear()

    def wait_element_get(self, XPATH, time=30):
        WebDriverWait(self.driver, time).until(EC.element_to_be_clickable(
            (By.XPATH, XPATH)))
        return self.driver.find_element(By.XPATH, XPATH)

    def wait_element_click(self, XPATH, time=60):
        WebDriverWait(self.driver, time).until(EC.element_to_be_clickable(
            (By.XPATH, XPATH)))
        self.driver.find_element(By.XPATH, XPATH).click()

    def wait_element_input(self, input, XPATH, time=30):
        WebDriverWait(self.driver, time).until(EC.element_to_be_clickable(
            (By.XPATH, XPATH)))
        self.driver.find_element(By.XPATH, XPATH).send_keys(input)\

    def wait_element_select_value(self, value: str, XPATH, time=30):
        WebDriverWait(self.driver, time).until(EC.element_to_be_clickable(
            (By.XPATH, XPATH)))
        Select(self.driver.find_element(By.XPATH, XPATH)).select_by_value(value)

    def wait_element_select_index(self, index: int, XPATH, time=30):
        WebDriverWait(self.driver, time).until(EC.element_to_be_clickable(
            (By.XPATH, XPATH)))
        Select(self.driver.find_element(By.XPATH, XPATH)).select_by_index(index)
    
    def close(self):
        """keluar dari chrome"""
        return self.driver.quit()
    
    def login(self):
        """Login ke SIMPEG"""
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

        except Exception as e:
          print(e)
          

    def fill_jurnal(self, jam_mulai:str, menit_mulai:str, jam_selesai:str, menit_selesai:str,
                    skp: int, skp_value: str, kegiatan: str, jumlah_diselesaikan: int):
        """isi jurnal"""
        max_retries = 10
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
            
                # JIKA BERHASIL TANPA ERROR. KELUAR DARI LOOP
                break
                # BTN BATAL
                # self.wait_element_click(XPATH="/html/body/div[4]/div[3]/button[1]")
            except TimeoutException:
                botlog.critical("TimeoutException: Situs pengisian jurnal tidak dapat diakses (Anda tidak terdaftar sebagai pegawai WFH)" )
                self.send_email(subject=f"Pengisian Jurnal SIMPEG KEMENKUMHAM Tanggal {self.date} Gagal",
                                body=f"Salam. Semoga anda dalam keadaan baik, saya ingin memberitahu anda bahwa jurnal harian untuk tanggal {self.date} gagal di isi. Situs pengisian jurnal tidak dapat diakses karena 'Anda tidak terdaftar sebagai pegawai WFH'.\n\nTerima kasih atas perhatiannya,\nSalam hormat.")
                self.timeout_occured = True
                break
            
            except Exception as e:
                retries += 1
                # botlog.error(f"Terjadi kesalahan {repr(e)}. Percobaan ke-{retries} dari {max_retries}")
                print(e)

                if retries == max_retries:
                    botlog.info("Jumlah percobaan maksimum telah tercapai. Tidak dapat melanjutkan.")
                    break
                
    def is_has_filled(self) -> bool:
        """Cek apakah tabel jurnalnya sudah terisi? jika sudah return True"""
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
        """
        memulai pengisian jurnal
        """
        botlog.info("================= TASK START =================")

        jurnal = self.get_jurnal()
        timeout_occured = False

        while not self.is_complete_fill:
            if self.is_complete_fill: break
            # JIKA TIDAK LIBUR LANJUT TASKS
            if not self.is_holiday():

                self.login()
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
                        if self.timeout_occured:break

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
                        if self.timeout_occured:break
                        
                    self.is_complete_fill = True
                    self.send_email(subject=f"Pengisian Jurnal SIMPEG KEMENKUMHAM Tanggal {self.date}",
                                    body=f"Salam. Semoga anda dalam keadaan baik, saya ingin memberitahu anda bahwa jurnal harian untuk tanggal {self.date} telah berhasil di isi. Berikut adalah rincian kegiatan hari ini:\
                                \n\n{self.parse_data_to_pretty_output(jurnal, 'jumat-sabtu')} \
                                \n\nTerima kasih atas perhatiannya,\
                                \nSalam hormat.")
                    botlog.info("FILL JURNAL JUMAT SABTU DONE")

            else:
                botlog.info("HARI INI LIBUR")
                break
            
        self.driver.close()
        botlog.info("================= TASK DONE =================")
