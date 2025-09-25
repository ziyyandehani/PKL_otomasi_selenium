from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import pandas as pd
import requests
import time
import os
import undetected_chromedriver as uc
import json
from datetime import datetime
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import sys

# === Setup log terminal ===
os.makedirs("log_terminal", exist_ok=True)
timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
log_path = os.path.join("log_terminal", f"terminal_{timestamp}.log")

class Tee:
    def __init__(self, *streams):
        self.streams = streams
    def write(self, data):
        for s in self.streams:
            s.write(data)
            s.flush()
    def flush(self):
        for s in self.streams:
            s.flush()

log_file = open(log_path, "w", encoding="utf-8")
sys.stdout = Tee(sys.stdout, log_file)
sys.stderr = Tee(sys.stderr, log_file)

print(f"📋 Logging terminal aktif. Semua output disimpan di: {log_path}")

def buka_dengan_cookie():
    options = uc.ChromeOptions()
    driver = uc.Chrome(options=options)
    driver.get("https://suradi.malangkota.go.id/")
    time.sleep(2)

    with open(r"D:/PKL/code/2. TUGAS 2/cookie.json", "r") as f:
        cookies = json.load(f)
        for cookie in cookies:
            for key in ['sameSite', 'storeId', 'session', 'id', 'hostOnly', 'expirationDate']:
                cookie.pop(key, None)
            if cookie.get("domain", "").startswith("."):
                cookie["domain"] = cookie["domain"][1:]
            driver.add_cookie(cookie)

    driver.get("https://suradi.malangkota.go.id/admin/dashboard")
    print("✅ Cookie berhasil digunakan, masuk dashboard")
    return driver

def verifikasi_dan_upload_ulang(driver, driver_suradi, pegawai_id, nip):
    print(f"\n🔁 Verifikasi ulang cuti 2025 ID Pegawai: {pegawai_id}")
    driver.get(f"https://simpeg.malangkota.go.id/kepegawaian/informasi/daftar_pegawai/detail_pegawai/{pegawai_id}/tab_disiplin")
    time.sleep(3)

    rows = driver.find_elements(By.XPATH, "//table[@id='datatable_cuti']//tbody/tr")
    ulang_upload = []

    for row in rows:
        tds = row.find_elements(By.TAG_NAME, "td")
        if len(tds) < 8:
            continue
        tanggal = tds[2].text.strip()
        tahun = tanggal.split("-")[-1]
        if tahun != "2025":
            continue

        td_file = tds[7]
        if td_file.find_elements(By.TAG_NAME, "a"):
            continue

        nomor_surat = tds[3].text.strip()
        td_opsi = tds[8]
        link_edit = td_opsi.find_element(By.XPATH, ".//a[contains(@href, 'edit_cuti')]").get_attribute("href")

        ulang_upload.append({
            'tanggal': tanggal,
            'nomor_surat': nomor_surat,
            'link_edit': link_edit
        })

    print(f"🔍 Masih ada {len(ulang_upload)} cuti 2025 tanpa file")

    for surat in ulang_upload:
        path_file, status = cek_dan_download_suradi(driver_suradi, surat['nomor_surat'])

        if status == "Ditolak":
            print(f"🚫 Surat {surat['nomor_surat']} tetap ditolak, dilewati")
            continue

        if path_file:
            try:
                upload_ke_simpeg(driver, surat['link_edit'], path_file)
                print("✅ Upload ulang berhasil")
            except Exception as e:
                print(f"❌ Gagal upload ulang: {e}")
        else:
            print(f"⚠️ Gagal download ulang surat: {status}")

def cek_dan_download_suradi(driver, nomor_surat, max_retry=3, delay=5):
    """
    Mengecek dan mendownload surat TTD dari SURADI dengan retry jika error.
    """
    for attempt in range(1, max_retry + 1):
        try:
            driver.get("https://suradi.malangkota.go.id/surat_bkpsdm/surat_pengajuan/TJS202206060000083")
            time.sleep(3)

            # Cari input search
            search_input = driver.find_element(By.CSS_SELECTOR, "input[type='search']")
            search_input.clear()
            search_input.send_keys(nomor_surat)
            time.sleep(2)
            search_input.send_keys(Keys.RETURN)
            time.sleep(2)

            rows = driver.find_elements(By.XPATH, "//table[@id='table_server']//tbody/tr")
            if not rows:
                print(f"❌ Nomor surat {nomor_surat} tidak ditemukan")
                return None, "Nomor surat tidak ditemukan"

            row = rows[0]
            status_cells = row.find_elements(By.TAG_NAME, "td")

            if any("Sudah TTE" in cell.text for cell in status_cells):
                try:
                    link_ttd = row.find_element(By.XPATH, ".//a[contains(text(), 'Surat TTD')]")
                    url_pdf = link_ttd.get_attribute("href")
                    print(f"✅ Surat TTD ditemukan: {url_pdf}")
                    file_name = url_pdf.split("/")[-1]
                    os.makedirs("hasil_download", exist_ok=True)
                    path = os.path.join("hasil_download", file_name)
                    r = requests.get(url_pdf, timeout=10)
                    with open(path, "wb") as f:
                        f.write(r.content)
                    print(f"📄 File disimpan: {path}")
                    return path, "Berhasil didownload"
                except Exception as e:
                    print(f"⚠️ Tidak ditemukan tombol download: {e}")
                    return None, "Link download tidak ditemukan"

            elif any("Ditolak" in cell.text for cell in status_cells):
                print(f"🚫 Nomor surat {nomor_surat} status: Ditolak")
                return None, "Ditolak"

            elif any("Diproses" in cell.text for cell in status_cells):
                print(f"🟡 Nomor surat {nomor_surat} masih diproses")
                return None, "Masih Diproses"

            else:
                print("⚠️ Status tidak dikenali")
                return None, "Status tidak dikenali"

        except Exception as e:
            print(f"⚠️ Error akses SURADI (attempt {attempt}/{max_retry}): {e}")
            if attempt < max_retry:
                print(f"🔄 Retry dalam {delay} detik...")
                time.sleep(delay)
                delay *= 2  # delay adaptif (5s -> 10s -> 20s)
            else:
                return None, f"Error SURADI setelah {max_retry} kali percobaan"

def upload_ke_simpeg(driver, link_edit, file_path):
    """
    Buka halaman edit cuti di SIMPEG dan unggah file surat cuti.

    Fungsi menavigasi ke `link_edit`, mengisi input file (name="file_surat_cuti") dan
    menekan tombol submit. Bila tombol tidak bisa diklik biasa, digunakan JS click.

    Args:
        driver: Selenium WebDriver yang sudah login ke SIMPEG.
        link_edit: URL halaman edit cuti yang akan diakses.
        file_path: path ke file PDF yang akan diunggah.

    Raises:
        NoSuchElementException: bila elemen input file atau tombol tidak ditemukan.
    """

    driver.get(link_edit)
    time.sleep(2)
    upload_input = driver.find_element(By.NAME, "file_surat_cuti")
    upload_input.send_keys(os.path.abspath(file_path))
    print(f"📤 Mengunggah file: {file_path}")
    dash_kota = driver.find_element(By.NAME, "lampiran_kota")
    driver.execute_script("arguments[0].value = arguments[1];",dash_kota, "-")

    wait = WebDriverWait(driver, 10)
    simpan_btn = wait.until(EC.element_to_be_clickable((By.ID, "submit_button")))
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", simpan_btn)
    time.sleep(0.5)

    try:
        simpan_btn.click()
    except Exception as e:
        print("⚠️ Tombol tidak bisa diklik biasa, pakai JS klik")
        driver.execute_script("arguments[0].click();", simpan_btn)
    time.sleep(2)
    print("✅ Upload ke SIMPEG berhasil")

service = Service("C:/Users/HP/Downloads/chromedriver-win64/chromedriver-win64/chromedriver.exe")
driver = webdriver.Chrome(service=service)
wait = WebDriverWait(driver, 10)

driver.get("https://simpeg.malangkota.go.id/login")
driver.find_element(By.ID, "username").send_keys("the-username")
driver.find_element(By.ID, "password").send_keys("the-password")
driver.find_element(By.CLASS_NAME, "btn-primary").click()
time.sleep(2)
tab_awal = driver.current_window_handle

df = pd.read_excel(r"D:/PKL/2. TUGAS 2/Data_Januari.xlsx")
driver_suradi = buka_dengan_cookie()

log_list = []

for index, row in df.iterrows():
    start_time = time.time()
    nip_raw = row['NIP Baru']
    nip = str(int(nip_raw)) if isinstance(nip_raw, float) else str(nip_raw)
    print(f"\n🔍 Proses NIP: {nip}")

    driver.get("https://simpeg.malangkota.go.id/kepegawaian/informasi/pencarian_pegawai")
    time.sleep(2)
    nip_input = wait.until(EC.presence_of_element_located((By.NAME, "nip_baru"))) 
    nip_input.clear() 
    nip_input.send_keys(nip)

    # Pilih opsi dari dropdown
    select_dropdown = Select(driver.find_element(By.NAME, "status_aktif"))
    select_dropdown.select_by_visible_text("Pegawai Aktif dan Non Aktif")
    time.sleep(2)
    # Klik tombol 'Cari'
    cari_button = wait.until(EC.element_to_be_clickable((By.ID, "search_button")))
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", cari_button)
    time.sleep(0.5)

    try:
        cari_button.click()
    except Exception:
        print("⚠️ Gagal klik biasa, pakai JS click...")
        driver.execute_script("arguments[0].click();", cari_button)

    time.sleep(4)
    # Klik tombol Detil (pakai JavaScript supaya pasti)
    detail_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(@title, 'Detil Data')]")))
    driver.execute_script("arguments[0].click();", detail_btn)
    time.sleep(3)

    driver.switch_to.window(driver.window_handles[-1])
    pegawai_id = driver.current_url.rstrip('/').split("/")[-1]
    print(f"✅ ID Pegawai: {pegawai_id}")

    driver.get(f"https://simpeg.malangkota.go.id/kepegawaian/informasi/daftar_pegawai/detail_pegawai/{pegawai_id}/tab_disiplin")
    time.sleep(4)

    baris_cuti_2025 = []

    while True:
        rows = driver.find_elements(By.XPATH, "//table[@id='datatable_cuti']//tbody/tr")
        ada_2025_di_halaman_ini = False

        for row in rows:
            tds = row.find_elements(By.TAG_NAME, "td")
            if len(tds) < 8:
                continue
            tanggal = tds[2].text.strip()
            tahun = tanggal.split("-")[-1]
            if tahun != "2025":
                continue

            ada_2025_di_halaman_ini = True  

            nomor_surat = tds[3].text.strip()
            td_file = tds[7]
            ada_file = bool(td_file.find_elements(By.TAG_NAME, "a"))

            td_opsi = tds[8]
            link_edit = td_opsi.find_element(By.XPATH, ".//a[contains(@href, 'edit_cuti')]").get_attribute("href")

            baris_cuti_2025.append({
                'tanggal': tanggal,
                'nomor_surat': nomor_surat,
                'link_edit': link_edit,
                'ada_file': ada_file
            })

        if not ada_2025_di_halaman_ini:
            break  # halaman ini gak ada 2025, langsung stop, hemat waktu

        try:
            next_button = driver.find_element(By.CSS_SELECTOR, "#datatable_cuti_paginate a.paginate_button.next:not(.disabled)")
            driver.execute_script("arguments[0].scrollIntoView(true);", next_button)
            time.sleep(0.5)
            next_button.click()
            time.sleep(1.5)
        except:
            break  # tidak ada tombol next, selesai

    print(f"📋 Total cuti 2025: {len(baris_cuti_2025)}")

    for surat in baris_cuti_2025:
        print(f"📄 Nomor surat: {surat['nomor_surat']} | Ada file: {'✅ Ya' if surat['ada_file'] else '❌ Tidak'}")

        if surat['ada_file']:
            status_upload = "Sudah Ada File"
        else:
            try:
                path_file, status = cek_dan_download_suradi(driver_suradi, surat['nomor_surat'])
                if path_file:
                    upload_ke_simpeg(driver, surat['link_edit'], path_file)
                    status_upload = "Sukses Upload"
                else:
                    print(f"⚠️ Gagal download: {status}")
                    status_upload = f"Gagal: {status}"
            except Exception as e:
                print(f"❌ Error upload: {str(e)}")
                status_upload = f"Error: {str(e)}"

        log_list.append({
            'NIP': nip,
            'Tanggal Surat': surat['tanggal'],
            'Nomor Surat': surat['nomor_surat'],
            'Status': status_upload
        })
    end_time = time.time()  # waktu selesai
    durasi = end_time - start_time
    print(f"⏱️ Durasi NIP {nip}: {durasi:.2f} detik")
    driver.close()
    driver.switch_to.window(tab_awal)
    verifikasi_dan_upload_ulang(driver, driver_suradi, pegawai_id, nip)

os.makedirs("log", exist_ok=True)
timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
log_filename = f"log/log_upload_cutian_{timestamp}.xlsx"
log_df = pd.DataFrame(log_list)
log_df_cleaned = log_df.copy()

log_df_cleaned['No'] = ''
log_df_cleaned['Total Surat 2025'] = ''
no = 1
last_nip = None
jumlah_surat_per_nip = log_df_cleaned.groupby('NIP').size().to_dict()

for i, row in log_df_cleaned.iterrows():
    nip = row['NIP']
    if nip and nip != last_nip:
        log_df_cleaned.at[i, 'No'] = no
        log_df_cleaned.at[i, 'Total Surat 2025'] = jumlah_surat_per_nip[nip]
        last_nip = nip
        no += 1
    else:
        log_df_cleaned.at[i, 'NIP'] = ''

log_df_cleaned = log_df_cleaned[['No', 'NIP', 'Total Surat 2025', 'Tanggal Surat', 'Nomor Surat', 'Status']]
log_df_cleaned.to_excel(log_filename, index=False)

print(f"✅ Semua NIP selesai diproses dan hasil dicatat di {log_filename}")
for d in [driver, driver_suradi]:
    try:
        d.quit()
    except Exception:
        pass
