from selenium import webdriver #Digunkaan untuk mengendalikan browser 
from selenium.webdriver.chrome.service import Service #Untuk mengatur service chromedriver
from selenium.webdriver.common.by import By #Untuk memilih elemen berdasarkan tipe tertentu
from selenium.webdriver.support.ui import Select #Untuk mengelola elemen <select> di HTML
import pandas as pd #Untuk mengelola data dalam format tabel (DataFrame)
from pathlib import Path    #Untuk mengelola path file dan direktori
import re #Untuk operasi berbasis pola pada string
import time  #Untuk menambahkan jeda waktu dalam eksekusi kode
from selenium.webdriver.support.ui import WebDriverWait #Untuk menunggu kondisi tertentu sebelum melanjutkan eksekusi
from selenium.webdriver.support import expected_conditions as EC #Untuk mendefinisikan kondisi yang diharapkan saat menunggu
from selenium.webdriver.common.keys import Keys #Untuk mengirimkan input keyboard ke elemen web

def normalisasi_jurusan(text):
    """
    Normalisasi teks jurusan/pendidikan agar lebih seragam.

    Fungsi ini mengubah input jurusan mentah (misalnya dari data pegawai) 
    ke format baku dengan beberapa aturan standarisasi dan mapping tertentu.

    Aturan normalisasi:
    - Mengubah huruf menjadi kapital, lalu menghapus spasi berlebih di awal/akhir.
    - Mengganti pola tertentu, misalnya:
        * "S-1 " → "SARJANA-"
        * "D-3 " → "DIPLOMA III-"
        * "SMA PAKET C " → "SMA-Paket C"
    - Jika ditemukan kata kunci dalam `mapping`, maka diganti sesuai nilai normalisasi.
      Contoh mapping:
        * "SLTA SEDERAJAT" → "Sekolah Menengah Atas"

    Args:
        text (str): Teks jurusan mentah.

    Returns:
        str: Hasil normalisasi jurusan dengan format baku.
             Jika tidak ada yang cocok, mengembalikan teks dengan kapitalisasi awal kata.

    Examples:
        >>> normalisasi_jurusan("S-1 TEKNIK INFORMATIKA")
        'Sarjana-Teknik Informatika'
        >>> normalisasi_jurusan("D-3 Akuntansi")
        'DIPLOMA III-Akuntansi'
        >>> normalisasi_jurusan("SLTA SEDERAJAT")
        'Sekolah Menengah Atas'
        >>> normalisasi_jurusan("SMA PAKET C IPS")
        'Sma-Paket C Ips'


        
    """
    text = text.upper().strip()

    # Standarisasi awal
    text = text.replace("S-1 ", "SARJANA-")
    text = text.replace("D-3 ", "DIPLOMA III-")
    text = text.replace("SMA PAKET C ", "SMA-Paket C")

    # Mapping kata kunci ke hasil normalisasi
    mapping = {
        #"SMA PAKET C": "SMA-Paket C",
        "SLTA SEDERAJAT": "SEKOLAH MENENGAH ATAS",
        #"STM": "SEKOLAH TEKNIK MENENGAH"
    }

    for keyword, hasil in mapping.items():
        if keyword in text:
            return hasil.title()

    return text.title()  # fallback kalau tidak cocok apa-apa

def normalisasi_jabatan(jabatan_raw):
    """
    Normalisasi teks jabatan mentah menjadi kategori baku.

    Fungsi ini mengubah input jabatan mentah menjadi salah satu dari empat kategori:
    - "Pelaksana" jika mengandung kata "pelaksana"
    - "Fungsional" jika mengandung kata "fungsional"
    - "Struktural" jika mengandung kata "struktural"
    - "lainnya" jika tidak cocok dengan tiga kategori di atas

    Args:
        jabatan_raw (str): Teks jabatan mentah (misalnya dari input spreadsheet atau database).

    Returns:
        str: Kategori jabatan yang sudah dinormalisasi.
             Nilai yang mungkin: "Pelaksana", "Fungsional", "Struktural", atau "lainnya".

    Examples:
        >>> normalisasi_jabatan("Jabatan Pelaksana")
        'Pelaksana'
        >>> normalisasi_jabatan("Tenaga Fungsional Umum")
        'Fungsional'
        >>> normalisasi_jabatan("Pejabat Struktural X")
        'Struktural'
        >>> normalisasi_jabatan("Magang")
        'lainnya'
    """
    jabat = str(jabatan_raw).lower()
    if "pelaksana" in jabat:
        return "Pelaksana"
    if "fungsional" in jabat:
        return "Fungsional"
    if "struktural" in jabat:
        return "Struktural"
    return "lainnya"

def ekstrak_sub_unit_unit_skpd(unor):
    """
    Ekstrak sub-unit, unit, dan SKPD dari field 'Unor' yang berupa teks lurus.

    Format input umumnya: "<Sub Unit> <Unit Keyword> <SKPD Keyword> <Nama SKPD>". 
    Fungsi mencari kata kunci SKPD ("Dinas", "Satuan", "Kecamatan") dari akhir teks,
    lalu mencari kata kunci unit ("Bidang", "Seksi", "UPT") di bagian sebelum SKPD.

    Args:
        unor: String yang berisi nama Unor (contoh: "Bidang Pelayanan Perizinan dan Nonperizinan Ekonomi, 
        Pariwisata dan Sosial Budaya Dinas Tenaga Kerja, Penanaman Modal dan Pelayanan Terpadu Satu Pintu").

    Returns:
        tuple yang berisi (sub_unit, unit_kerja, skpd).
        - sub_unit: teks sebelum kata kunci unit (atau seluruh sisa bila tidak ditemukan)
        - unit_kerja: bagian yang dimulai dari kata kunci unit (atau "" jika tidak ada)
        - skpd: bagian yang dimulai dari kata kunci SKPD (atau "" jika tidak ada)

    Examples:
        >>> ekstrak_sub_unit_unit_skpd("Seksi Pengendalian Bidang Ketertiban Dinas X")
        ("Seksi Pengendalian", "Bidang Ketertiban", "Dinas X")
    """
    unor = str(unor)

    skpd_keywords = ["Dinas", "Satuan", "Kecamatan"]
    unit_keywords = ["Bidang", "SMPN", "Kelurahan", "Seksi", "Sekretariat", "UPT", "Puskesmas"]

    skpd1 = ""
    unit_kerja1 = ""
    sub_unit1 = ""

    skpd_idx = -1
    unit_idx = -1

    # Cari SKPD paling akhir (karena biasanya di akhir kalimat)
    for kata in skpd_keywords:
        idx = unor.rfind(kata)
        if idx != -1:
            skpd_idx = idx
            skpd1 = unor[idx:].strip()
            break

    # Sisa sebelum SKPD → cari unit kerja
    sisa = unor[:skpd_idx].strip() if skpd_idx != -1 else unor

    for kata in unit_keywords:
        idx = sisa.find(kata)
        if idx != -1:
            unit_idx = idx
            unit_kerja1 = sisa[idx:].strip()
            break

    # Sisanya jadi sub unit
    if unit_idx != -1:
        sub_unit1 = sisa[:unit_idx].strip()
    else:
        sub_unit1 = sisa

    return sub_unit1, unit_kerja1, skpd1
   
# === 1. Setup Driver & Login SIMPEG ===
# Bagian ini bertugas untuk:
# - Menginisialisasi ChromeDriver dengan path lokal
# - Membuka halaman login SIMPEG Kota Malang
# - Mengisi username & password lalu menekan tombol login
# - Menunggu halaman beralih setelah login
# Catatan: ganti "the-username" dan "the-password" dengan kredensial asli.

service = Service("C:/Users/HP/Downloads/chromedriver-win64/chromedriver-win64/chromedriver.exe")
driver = webdriver.Chrome(service=service)

driver.get("https://simpeg.malangkota.go.id/login")
driver.find_element(By.ID, "username").send_keys("the-username")
driver.find_element(By.ID, "password").send_keys("the-password")
driver.find_element(By.CLASS_NAME, "btn-primary").click()
time.sleep(2)
wait = WebDriverWait(driver, 10)

# === 2. Baca File Excel ===
# Bagian ini bertugas untuk:
# - Membaca data pegawai dari file Excel "data_jabatan.xlsx"
# - Menyediakan list kosong `log_gagal` untuk mencatat data yang gagal diproses
# - Menentukan path folder tempat file SK (dokumen TTE) disimpan
#   agar bisa dicocokkan dengan data pegawai
data = pd.read_excel("data_jabatan.xlsx")
log_gagal = []
folder_path = Path(r"D:\PKL\code\cek_dan_perbaiki\TTE SPMT PPPK T1 2024")


for index, row in data.iterrows():
    # === Proses Data Pegawai ===
    # Bagian ini melakukan loop per baris data di file Excel.
    # Setiap baris berisi informasi satu pegawai, NIP, data pendidikan, dan jabatan.
    # Tahapan utamanya:
    # 1. Ambil NIP pegawai (dibersihkan agar konsisten dalam bentuk string).
    # 2. Ambil data pendidikan: tahun lulus, nomor ijazah, kepala sekolah, tanggal ijazah, jurusan, lembaga.
    # 3. Ambil data jabatan: nomor SPMT, tanggal SPMT, TMT SPMT, jenis jabatan, nama jabatan.
    # 4. Lakukan normalisasi data (jurusan & jabatan) agar sesuai format input SIMPEG.
    # 5. Ekstrak unit kerja, sub unit, dan SKPD dari kolom 'Unor'.

    print(f"\n🚀 Proses baris ke-{index + 1} NIP: {row['NIP Baru']}")
    nip_raw = row['NIP Baru']
    nip = str(int(nip_raw)) if isinstance(nip_raw, float) else str(nip_raw)
    data = pd.read_excel("data_jabatan.xlsx", dtype={"Tanggal SPMT": str, "TMT SPMT": str, "Tanggal Ijazah": str})

    tahun_lulus = str(row['Tahun Lulus'])
    no_ijazah = row['No. Ijazah']
    kepala = row['Kepala Sekolah']
    tanggal = pd.to_datetime(row['Tanggal Ijazah'], dayfirst=True).strftime("%d-%m-%Y")
    jurusan_excel = normalisasi_jurusan(row['Jurusan'])
    lembaga_excel = row['Lembaga']

    keterangan = str(row['No.SPMT']) 
    parts = keterangan.split('/')
    no_spmt = str(row['No.SPMT'])
    tanggal_spmt = pd.to_datetime(row['Tanggal SPMT'], dayfirst=True).strftime("%d-%m-%Y")
    tanggal_tmt = pd.to_datetime(row['TMT SPMT'], dayfirst=True).strftime("%d-%m-%Y")
    
    jabatan_excel = normalisasi_jabatan(row['JENIS JABATAN NAMA'])
    jabatan_nama = row['JABATAN NAMA']
    sub_unit, unit_kerja, skpd = ekstrak_sub_unit_unit_skpd(row['Unor'])

    # === Navigasi ke Halaman Detail Pegawai ===
    # Bagian ini melakukan pencarian data pegawai berdasarkan NIP, 
    # lalu membuka halaman detail pegawai untuk mendapatkan ID Pegawai.
    # Tahapannya:
    # 1. Buka halaman daftar pegawai SIMPEG.
    # 2. Masukkan NIP ke input pencarian.
    # 3. Klik tombol 'Cari Data' (menggunakan scroll + JavaScript click agar lebih stabil).
    # 4. Tunggu hasil pencarian muncul, lalu klik tombol 'Detil'.
    # 5. Berpindah ke tab baru yang menampilkan detail pegawai.
    # 6. Ambil ID Pegawai dari URL halaman.

    driver.get("https://simpeg.malangkota.go.id/kepegawaian/informasi/daftar_pegawai")
    time.sleep(3)
    tab_awal = driver.current_window_handle
    nip_input = driver.find_element(By.ID, "nip_baru")
    nip_input.clear()
    nip_input.send_keys(nip)
    cari_btn = driver.find_element(By.CLASS_NAME, "btn-primary") # Tombol 'Cari Data'

    driver.execute_script("arguments[0].scrollIntoView(true);", cari_btn) # Scroll dulu 
    time.sleep(0.5) 
    driver.execute_script("arguments[0].click();", cari_btn) # baru klik cari data
    time.sleep(4)
    
    detail_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(@title, 'Detil')]")))
    driver.execute_script("arguments[0].click();", detail_btn) # klik tombol detil
    time.sleep(2)
    
    driver.switch_to.window(driver.window_handles[-1]) # Ganti fokus ke tab baru
    print("🧭 Berpindah ke tab baru.")

    current_url = driver.current_url
    pegawai_id = current_url.rstrip('/').split("/")[-1]
    print("🆔 ID Pegawai:", pegawai_id)

    # === Mengisi Data Pendidikan Pegawai ===
    # Bagian ini langsung mengakses form tambah pendidikan berdasarkan ID Pegawai.
    # Tahapannya:
    # 1. Susun URL form tambah pendidikan menggunakan ID Pegawai.
    # 2. Akses halaman form pendidikan dengan driver.get().
    # 3. Isi kolom-kolom form pendidikan sesuai data dari Excel:
    #    - Tanggal Ijazah
    #    - Nama Kepala Sekolah
    #    - Tahun Lulus
    #    - Nomor Ijazah
    # 4. Pastikan field 'keterangan' dikosongkan dulu sebelum diisi ulang.
    # 5. Masukkan nomor ijazah ke field keterangan.

    print ("Mengisi Pendidikan")
    form_url = f"https://simpeg.malangkota.go.id/kepegawaian/informasi/pegawai_pendidikan/add_pendidikan/{pegawai_id}"
    driver.get(form_url)
    print("🔗 Akses langsung ke form tambah pendidikan:", form_url)
    time.sleep(3)

    driver.find_element(By.ID, "tanggal_ijazah").send_keys(tanggal)
    driver.find_element(By.NAME, "nama_kepala").send_keys(kepala)
    driver.find_element(By.ID, "tahun_lulus").send_keys(tahun_lulus)
    driver.find_element(By.NAME, "keterangan").clear()
    driver.find_element(By.NAME, "keterangan").send_keys(no_ijazah)

    try: # === Jurusan ===
        select_jurusan = Select(driver.find_element(By.ID, "jurusan"))
        select_jurusan.select_by_visible_text(jurusan_excel.upper())
        time.sleep(1)
        print(f"✅ Jurusan '{jurusan_excel.upper()}' berhasil dipilih.")
    except Exception as e:
        log_gagal.append({
            "NIP": nip,
            "Keterangan": f"error saat memilih jurusan: {e}"
        })
        print (f"❌ Gagal memilih jurusan: {e}")

    try: # === Lembaga ===
        select_box = wait.until(EC.element_to_be_clickable((By.ID, "select2-categories-container")))
        select_box.click()
        time.sleep(0.5)

        input_box = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input.select2-search__field")))
        input_box.send_keys(lembaga_excel.upper())
        time.sleep(1) 

        input_box.send_keys(Keys.ENTER)
        time.sleep(1)

        selected_lembaga = driver.find_element(By.ID, "select2-categories-container").text.strip().upper()
        if lembaga_excel.upper() in selected_lembaga:
            print(f"✅ Lembaga: '{selected_lembaga}' berhasil dipilih.")
        else:
            raise ValueError(f"Lembaga tidak cocok: '{selected_lembaga}' ≠ '{lembaga_excel.upper()}'")
    except Exception as e:
        log_gagal.append({
            "NIP": nip,
            "Keterangan": f"Error saat milih lembaga: {e}"
        })
        print (f"❌ Gagal memilih lembaga: {e}")

    try: # Pendidikan CPNS 
        pendidikan_cpns_box = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[@id='select2-pendidikan_cpns-container']")))
        pendidikan_cpns_box.click()
        time.sleep(0.5)

        pendidikan_cpns_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input.select2-search__field")))
        pendidikan_cpns_input.send_keys("Ya")
        time.sleep(0.5)
        pendidikan_cpns_input.send_keys(Keys.ENTER)
    except Exception as e:
        log_gagal.append({
            "NIP": nip,
            "Keterangan": f"Error saat milih pendik_cpns: {e}"
        })
        print (f"❌ Gagal memilih pendik_cpns: {e}")
            
    time.sleep(1)
    submit_btn = wait.until(EC.element_to_be_clickable((By.ID, "submit_button")))
    submit_btn.click()
    time.sleep(3)
    errors = driver.find_elements(By.CSS_SELECTOR, ".error-block")
    if errors:
        error_messages = [e.text.strip() for e in errors if e.text.strip()]
        combined_error = "; ".join(error_messages)
            
        log_gagal.append({
            "NIP": nip,
            "Keterangan": f"Form pendidikan tidak berhasil disimpan. Error: {combined_error}"
        })
        print(f"❌ Form pendidikan gagal disimpan: {combined_error}")
    else:
        print(f"✅ Data pendidikan {nip} berhasil disimpan.")      

    print ("\n ======= Mengisi Main Jabatan ========")
    form_url = f"https://simpeg.malangkota.go.id/kepegawaian/informasi/pegawai_jabatan/tambah_jabatan/{pegawai_id}"
    driver.get(form_url)
    print("🔗 Akses langsung ke form tambah jabatan:", form_url)
    time.sleep(3)

    try:
        print(f"🎯 Mencoba pilih jabatan: {jabatan_excel}")
        jabatan_box = wait.until(EC.element_to_be_clickable((By.ID, "select2-jenis_jabatan-container")))
        jabatan_box.click()
        time.sleep(1)

        jabatan_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input.select2-search__field")))
        jabatan_input.send_keys(jabatan_excel.upper()) 
        time.sleep(1)
        jabatan_input.send_keys(Keys.ENTER)
        time.sleep(1)
        print(f"✅ Jabatan '{jabatan_excel}' berhasil dipilih.")
    except Exception as e:
        log_gagal.append({
            "NIP": nip,
            "Keterangan": f"Gagal memilih jabatan: {e}"
        })
        print(f"❌ Gagal memilih jabatan: {e}")

    try:
        print(f"🎯 Mencoba pilih SKPD: {skpd}")

        skpd_box = wait.until(EC.element_to_be_clickable((By.ID, "select2-skpd-container")))
        skpd_box.click()
        time.sleep(1)

        skpd_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input.select2-search__field")))
        skpd_input.send_keys(skpd.upper())
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".select2-results__option--highlighted")))
        skpd_input.send_keys(Keys.ENTER)
        time.sleep(1)
        print(f"✅ SKPD '{skpd}' berhasil dipilih.")
    except Exception as e:
        log_gagal.append({
            "NIP": nip,
            "Keterangan": f"Gagal memilih SKPD: {e}"
        })
        print(f"❌ Gagal memilih SKPD: {e}")

    try:
        print(f"🎯 Mencoba pilih unit kerja: {unit_kerja}")

        unit_box = wait.until(EC.element_to_be_clickable((By.ID, "select2-unit_kerja-container")))
        unit_box.click()
        time.sleep(1)

        unit_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input.select2-search__field")))
        unit_input.send_keys(unit_kerja.upper())
        time.sleep(2.5)
        unit_input.send_keys(Keys.ENTER)
        time.sleep(1)

        print(f"✅ unit kerja '{unit_kerja}' berhasil dipilih.")
    except Exception as e:
        log_gagal.append({
            "NIP": nip,
            "Keterangan": f"Gagal memilih unit kerja: {e}"
        })
        print(f"❌ Gagal memilih unit kerja: {e}")

    if sub_unit:
        try:
            print(f"🎯 Mencoba pilih sub unit kerja: {sub_unit}")

            subunit_box = wait.until(EC.element_to_be_clickable((By.ID, "select2-sub_unit_kerja-container")))
            subunit_box.click()
            time.sleep(1)

            subunit_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input.select2-search__field")))
            subunit_input.send_keys(sub_unit.upper())
            time.sleep(2.5)
            subunit_input.send_keys(Keys.ENTER)
            time.sleep(1)

            print(f"✅ Sub unit kerja '{sub_unit}' berhasil dipilih.")
        except Exception as e:
            log_gagal.append({
                "NIP": nip,
                "Keterangan": f"Gagal memilih sub unit kerja: {e}"
            })
            print(f"❌ Gagal memilih sub unit kerja: {e}")
    else:
        print("ℹ️ Tidak ada sub unit kerja untuk dipilih.")

    try:
        print(f"🎯 Mencoba pilih nama jabatan {jabatan_excel}: {jabatan_nama}")
        jabfung_box = None
        if jabatan_excel == "Fungsional":
            jabfung_box = wait.until(EC.presence_of_element_located((By.ID, "select2-jab_fungsional-container")))
        elif jabatan_excel == "Pelaksana":
            jabfung_box = wait.until(EC.presence_of_element_located((By.ID, "select2-jab_pelaksana-container")))
        elif jabatan_excel == "Struktural":
            jabfung_box = wait.until(EC.presence_of_element_located((By.ID, "select2-jab_struktural-container")))
        jabfung_box.click()
        time.sleep(0.5)
            
        jabfung_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input.select2-search__field")))
        jabfung_input.clear()
        jabfung_input.send_keys(jabatan_nama.upper())
        time.sleep(2)
        jabfung_input.send_keys(Keys.ENTER)
        time.sleep(1)

        print(f"✅ Jabatan '{jabatan_excel}' '{jabatan_nama}' berhasil dipilih.")
    except Exception as e:
        log_gagal.append({
            "NIP": nip,
            "Keterangan": f"Gagal memilih nama jabatan '{jabatan_excel}': {e}"
        })
        print(f"❌ Gagal memilih nama jabatan '{jabatan_excel}': {e}")

    try:
        select_pejabat = Select(driver.find_element(By.ID, "pejabat"))
        select_pejabat.select_by_visible_text("Sekretaris Daerah")
        print("✅ 'Sekretaris Daerah' berhasil dipilih.")
    except Exception as e:
        log_gagal.append({
            "NIP": nip,
            "Keterangan": f"Gagal memilih pejabat pengangkatan: {e}"
        })
        print(f"❌ Gagal memilih pejabat pengangkatan: {e}")

    try:
        nomor_sk_input = wait.until(EC.presence_of_element_located((By.ID, "nomor_sk")))
        nomor_sk_input.clear()
        nomor_sk_input.send_keys(no_spmt)

        tanggal_sk_input = driver.find_element(By.ID, "inp_tanggal_sk")
        driver.execute_script("arguments[0].value = arguments[1];", tanggal_sk_input, tanggal_spmt)

        tmt_sk_input = driver.find_element(By.ID, "inp_tmt_sk")
        driver.execute_script("arguments[0].value = arguments[1];", tmt_sk_input, tanggal_tmt)

        tmt_pelantikan_input = driver.find_element(By.ID, "inp_tmt_pelantikan")
        driver.execute_script("arguments[0].value = arguments[1];", tmt_pelantikan_input, tanggal_tmt)

        tmt_mutasi_input = driver.find_element(By.ID, "inp_tmt_mutasi")
        driver.execute_script("arguments[0].value = arguments[1];", tmt_mutasi_input, tanggal_tmt)

        print("✅ Nomor SK, Tanggal SK, TMT pelantikan, dan TMT mutasi Jabatan berhasil diisi.")

    except Exception as e:
        log_gagal.append({
            "NIP": nip,
            "Keterangan": f"Gagal mengisi data no.SK: {e}"
        })
        print(f"❌ Gagal mengisi data no.SK: {e}")
    
    driver.find_element(By.NAME, "ket_pejabat").send_keys("Sekretaris Daerah Kota Malang")

    if len(parts) >= 2:
        kode = parts[1]
        pattern = re.compile(rf"^SPMT_PPPK_T1_\d+_{kode}_")

        matching_files = [
            f for f in folder_path.iterdir()
            if f.is_file() and pattern.match(f.name)
        ]

        if matching_files:
            file_sk_path = str(matching_files[0])
            print(f"✅ File ditemukan: {file_sk_path}")

            file_input = wait.until(EC.presence_of_element_located((By.ID, "file_spmt")))
            driver.execute_script("arguments[0].scrollIntoView(true);", file_input)
            file_input.send_keys(file_sk_path)
            time.sleep(3)

        else:
            log_gagal.append({
                "NIP": nip,
                "Keterangan": f"file SPMT nya tidak ada"                    
            })
            print(f"❌ Tidak ditemukan file untuk kode {kode} dengan pola prefix 'SPMT_PPPK_T1_<x>_{kode}_'")
    else:
        print(f"❌ Format SPMT tidak valid: {keterangan}")

    try:
        time.sleep(1)

        # Scroll dan klik tombol Submit
        submit_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Submit')]")))
        driver.execute_script("arguments[0].scrollIntoView(true);", submit_btn)
        time.sleep(0.5)
        driver.execute_script("arguments[0].click();", submit_btn)  
        time.sleep(2)

        try:
            WebDriverWait(driver, 3).until(EC.alert_is_present())
            alert = driver.switch_to.alert
            print(f"⚠ Alert muncul: {alert.text}")
            alert.accept()
            print("✅ Alert dikonfirmasi.")
            time.sleep(2)
        except:
            print("ℹ️ Tidak ada alert konfirmasi setelah submit.")

        errors = driver.find_elements(By.CSS_SELECTOR, ".error-block")
        if errors:
            error_messages = [e.text.strip() for e in errors if e.text.strip()]
            combined_error = "; ".join(error_messages)

            log_gagal.append({
                "NIP": nip,
                "Keterangan": f"Form tidak berhasil disimpan. Error: {combined_error}"
            })
            print(f"❌ Form Jabatan gagal disimpan: {combined_error}")
        else:
            print(f"✅ Data Jabatan {nip} berhasil disimpan.")

    except Exception as e:
        log_gagal.append({
            "NIP": nip,
            "Keterangan": f"Gagal saat submit: {e}"
        })
        print(f"❌ Error saat proses submit: {e}")

    driver.close()
    driver.switch_to.window(tab_awal)

driver.quit()
print("✨ Semua data selesai diproses.")
if log_gagal:
    df_log = pd.DataFrame(log_gagal)
    df_log.to_excel("log-semua-kegagalan.xlsx", index=False)
    print("📝 Log kegagalan disimpan di 'log-semua-kegagalan.xlsx'")
else:
    print("🎉 Tidak ada data yang gagal diproses.")