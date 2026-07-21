import math
import os
from datetime import datetime
import openpyxl
import pandas as pd
import plotly.express as px
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from scipy.stats import norm
import streamlit as st

# ============================================================
# 1. KONFIGURASI HALAMAN DASHBOARD & HEADER UTAMA
# ============================================================
st.set_page_config(
    page_title="Smart Inventory Systems",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Tampilan Header Rapi (Tidak Terpotong & Tidak Duplikat)
st.markdown(
    """
    <div style="padding-bottom: 10px;">
        <h2 style="font-size: 30px; font-weight: 700; margin-bottom: 4px;">
            📦 INTEGRATED INVENTORY CONTROL & REAL-TIME MONITORING SYSTEM
        </h2>
        <p style="font-size: 21px; color: #A0AAB8; font-weight: 500; margin-top: 0;">
            Alat Pendukung Keputusan Persediaan Probabilistik Berbasis Algoritma Hadley–Whitin
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)

# Nama file lokal untuk menyimpan riwayat & template
FILE_RIWAYAT = "riwayat_transaksi.xlsx"
FILE_TEMPLATE = "Template_Master_Data_Inventory_Hadley_Whitin.xlsx"


# ============================================================
# FUNGSI PEMBUAT TEMPLATE EXCEL OTOMATIS (Selalu Kosongan)
# ============================================================
def buat_template_excel():
  wb = openpyxl.Workbook()

  # --- SHEET 1: Master Data Parameter ---
  ws1 = wb.active
  ws1.title = "Master Data Parameter"
  ws1.views.sheetView[0].showGridLines = True

  header_fill = PatternFill(
      start_color="1F4E78", end_color="1F4E78", fill_type="solid"
  )
  header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")

  ws1["A1"] = "TEMPLATE MASTER DATA PERSEDIAAN GUDANG (HADLEY-WHITIN)"
  ws1["A1"].font = Font(name="Arial", size=14, bold=True, color="1F4E78")

  headers_ws1 = ["Jenis Barang", "D", "Sigma", "L", "A", "h", "Cu"]
  ws1.row_dimensions[3].height = 25
  for col_num, h_text in enumerate(headers_ws1, 1):
    cell = ws1.cell(row=3, column=col_num)
    cell.value = h_text
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = Alignment(horizontal="center", vertical="center")

  # --- SHEET 2: Kalkulator Pemula ---
  ws2 = wb.create_sheet(title="Kalkulator Pemula")
  ws2.views.sheetView[0].showGridLines = True

  ws2["A1"] = "KALKULATOR PENENTU PARAMETER OTOMATIS UNTUK ORANG AWAM"
  ws2["A1"].font = Font(name="Arial", size=14, bold=True, color="1F4E78")

  headers_ws2 = [
      "Jenis Barang",
      "Jan",
      "Feb",
      "Mar",
      "Apr",
      "Mei",
      "Jun",
      "Jul",
      "Agt",
      "Sep",
      "Okt",
      "Nov",
      "Des",
      "D (Total Demand)",
      "Sigma (Std Deviasi)",
      "Lead Time (Hari)",
      "L (Tahun)",
      "Biaya Pesan (A)",
      "Biaya Simpan (h)",
      "Biaya Shortage (Cu)",
  ]
  ws2.row_dimensions[3].height = 25
  for col_num, h_text in enumerate(headers_ws2, 1):
    cell = ws2.cell(row=3, column=col_num)
    cell.value = h_text
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = Alignment(horizontal="center", vertical="center")

  wb.save(FILE_TEMPLATE)


# Panggil fungsi ini
buat_template_excel()

# ============================================================
# 2. INISIALISASI DATABASE & LOG TRANSAKSI
# ============================================================
if "data_gudang" not in st.session_state:
  st.session_state["data_gudang"] = None
if "stok_realtime" not in st.session_state:
  st.session_state["stok_realtime"] = {}


def bersihkan_angka(nilai):
  if pd.isna(nilai):
    return 0
  return pd.to_numeric(
      str(nilai).strip().replace(".", "").replace(",", "."), errors="coerce"
  )


def format_indonesia_kg(nilai):
  if nilai == "-" or pd.isna(nilai) or nilai is None:
    return "-"
  try:
    val = float(nilai)
    return (
        f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        + " kg"
    )
  except:
    return str(nilai)


def format_rupiah(nilai):
  try:
    val = float(nilai)
    return f"Rp {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
  except:
    return str(nilai)


def muat_riwayat():
  if os.path.exists(FILE_RIWAYAT):
    try:
      df_log = pd.read_excel(FILE_RIWAYAT)
      df_log["Waktu"] = pd.to_datetime(df_log["Waktu"])
      return df_log
    except:
      pass
  return pd.DataFrame(
      columns=[
          "Waktu",
          "Tanggal",
          "Bulan",
          "Bahan Baku",
          "Aktivitas",
          "Jumlah (kg)",
      ]
  )


def catat_transaksi(bahan, aktivitas, jumlah):
  df_log = muat_riwayat()
  sekarang = datetime.now()

  data_baru = pd.DataFrame([{
      "Waktu": sekarang,
      "Tanggal": sekarang.strftime("%Y-%m-%d"),
      "Bulan": sekarang.strftime("%Y-%m"),
      "Bahan Baku": bahan,
      "Aktivitas": aktivitas,
      "Jumlah (kg)": jumlah,
  }])

  df_log = pd.concat([df_log, data_baru], ignore_index=True)
  df_log.to_excel(FILE_RIWAYAT, index=False)


# ============================================================
# 3. SIDEBAR UPLOAD & PENGATURAN PARAMETER
# ============================================================
st.sidebar.header("⚙️ Langkah 1: Upload Master Data")
uploaded_file = st.sidebar.file_uploader(
    "Upload File Excel/CSV Data Logistik", type=["csv", "xlsx"]
)

if os.path.exists(FILE_TEMPLATE):
  with open(FILE_TEMPLATE, "rb") as f_temp:
    st.sidebar.download_button(
        label="📥 Unduh Template Excel Otomatis",
        data=f_temp.read(),
        file_name="Template_Master_Data_Inventory.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
        help="Unduh file kosongan ini jika Anda ingin template master data!",
    )

max_iter = st.sidebar.slider(
    "Batas Maksimum Iterasi Matematika", min_value=10, max_value=100, value=50
)

# ============================================================
# 4. PROSES HITUNG HADLEY-WHITIN KETIKA FILE DI-UPLOAD
# ============================================================
if uploaded_file is not None and st.session_state["data_gudang"] is None:
  try:
    df = (
        pd.read_csv(uploaded_file, delimiter=";")
        if uploaded_file.name.endswith(".csv")
        else pd.read_excel(uploaded_file)
    )
    df.columns = df.columns.str.strip()

    hasil_akhir = []
    df_log_ada = muat_riwayat()

    for index, row in df.iterrows():
      bahan = (
          row["Jenis Barang"] if "Jenis Barang" in row else row["Bahan Baku"]
      )
      D = bersihkan_angka(row["D"])
      sigma = bersihkan_angka(row["Sigma"])
      L = bersihkan_angka(row["L"])
      A = bersihkan_angka(row["A"])
      h = bersihkan_angka(row["h"])
      Cu = bersihkan_angka(row["Cu"])

      nama_bahan = str(bahan).strip().lower()

      # Klasifikasi Kategori dan Sistem Kontrol
      if "aluminium" in nama_bahan or "copper" in nama_bahan:
        kategori = "A"
        metode = "Continuous Review (s,S)"
      else:
        kategori = "B/C"
        metode = "Continuous Review (s,Q)"

      sigma_L, demand_L = sigma * math.sqrt(L), D * L
      q_prev = math.sqrt((2 * A * D) / h) if h > 0 else 0
      r_prev = demand_L + (
          norm.ppf(
              1
              - max(
                  min((h * q_prev) / (Cu * D if Cu * D > 0 else 1), 0.999999),
                  0.000001,
              )
          )
          * sigma_L
      )

      for i in range(1, max_iter + 1):
        alpha = max(
            min((h * q_prev) / (Cu * D if Cu * D > 0 else 1), 0.999999), 0.000001
        )
        Z = norm.ppf(1 - alpha)
        N = sigma_L * (norm.pdf(Z) - Z * (1 - norm.cdf(Z)))
        q_new = math.sqrt((2 * D * (A + Cu * N)) / h) if h > 0 else 0
        r_new = demand_L + (
            norm.ppf(
                1
                - max(
                    min((h * q_new) / (Cu * D if Cu * D > 0 else 1), 0.999999),
                    0.000001,
                )
            )
            * sigma_L
        )

        if round(r_new, 2) == round(r_prev, 2):
          q_prev, r_prev = q_new, r_new
          break
        q_prev, r_prev = q_new, r_new

      SS_final = r_prev - demand_L

      # Nilai S_max HANYA ADA jika Kategori A (s,S). Jika Kategori B/C (s,Q), diset None
      S_final = round(q_prev + r_prev, 2) if kategori == "A" else None

      OP = (A * D) / q_prev if q_prev > 0 else 0
      OS = h * ((q_prev / 2) + r_prev - demand_L)
      alpha_f = max(
          min((h * q_prev) / (Cu * D if Cu * D > 0 else 1), 0.999999), 0.000001
      )
      OK = (
          Cu
          * (D / q_prev if q_prev > 0 else 0)
          * (
              sigma_L
              * (
                  norm.pdf(norm.ppf(1 - alpha_f))
                  - norm.ppf(1 - alpha_f) * (1 - norm.cdf(norm.ppf(1 - alpha_f)))
              )
          )
      )

      stok_kalkulasi = round(r_prev + SS_final, 2)

      if not df_log_ada.empty:
        log_bahan = df_log_ada[df_log_ada["Bahan Baku"] == bahan]
        masuk = log_bahan[
            log_bahan["Aktivitas"] == "Barang Datang (Stok Masuk)"
        ]["Jumlah (kg)"].sum()
        keluar = log_bahan[
            log_bahan["Aktivitas"] == "Diambil Produksi (Stok Keluar)"
        ]["Jumlah (kg)"].sum()
        stok_kalkulasi = stok_kalkulasi + masuk - keluar

      stok_kalkulasi = float(stok_kalkulasi)
      st.session_state["stok_realtime"][bahan] = stok_kalkulasi

      hasil_akhir.append({
          "Bahan Baku": bahan,
          "Kategori": kategori,
          "Metode": metode,
          "Q_opt (kg)": round(q_prev, 2),
          "s_opt (kg)": round(r_prev, 2),
          "S_max (kg)": S_final,  # Diset None jika B/C
          "SS (kg)": round(SS_final, 2),
          "Total Cost (OT)": round(OP + OS + OK, 2),
      })

    st.session_state["data_gudang"] = pd.DataFrame(hasil_akhir)
  except Exception as e:
    st.error(f"Gagal membaca file: {e}")

# ============================================================
# 5. TAMPILAN UTAMA DASHBOARD
# ============================================================
if st.session_state["data_gudang"] is not None:
  df_hasil = st.session_state["data_gudang"].copy()
  df_hasil["Stok Saat Ini (kg)"] = df_hasil["Bahan Baku"].map(
      st.session_state["stok_realtime"]
  )

  def cek_status(row):
    if row["Stok Saat Ini (kg)"] <= row["s_opt (kg)"]:
      return "🚨 HARUS REORDER!"
    return "✅ Stok Aman"

  df_hasil["Status Gudang"] = df_hasil.apply(cek_status, axis=1)

  # --------------------------------------------------------
  # BAGIAN A: PANEL NOTIFIKASI DARURAT
  # --------------------------------------------------------
  item_butuh_reorder = df_hasil[df_hasil["Status Gudang"] == "🚨 HARUS REORDER!"]
  if not item_butuh_reorder.empty:
    st.error("### ⚠️ PERINGATAN REORDER DI GUDANG!")
    for idx, row in item_butuh_reorder.iterrows():
      stok_sekarang = row["Stok Saat Ini (kg)"]

      # REVISI 3: Logika Rekomendasi Reorder Berdasarkan Kategori A (s,S) atau B/C (s,Q)
      if row["Kategori"] == "A":
        target_S = row["S_max (kg)"]
        jumlah_harus_dipesan = max(
            round(target_S - stok_sekarang, 2), row["Q_opt (kg)"]
        )
        pesan_rekomendasi = (
            f"👉 **Rekomendasi Sistem (s,S):** Segera lakukan pemesanan dinamis"
            f" sebanyak **{format_indonesia_kg(jumlah_harus_dipesan)}** untuk"
            " mengejar Target Maximum Inventory"
            f" (**{format_indonesia_kg(target_S)}**)."
        )
      else:
        jumlah_harus_dipesan = row["Q_opt (kg)"]
        pesan_rekomendasi = (
            f"👉 **Rekomendasi Sistem (s,Q):** Segera lakukan pemesanan ulang"
            " sebesar lot tetap ($Q_{\\text{opt}}$) sebanyak"
            f" **{format_indonesia_kg(jumlah_harus_dipesan)}**."
        )

      st.warning(
          f"**{row['Bahan Baku']}** ({row['Kategori']}) Kritis! Sisa stok"
          f" **{format_indonesia_kg(stok_sekarang)}** telah melewati batas"
          " aman reorder"
          f" **{format_indonesia_kg(row['s_opt (kg)'])}**.\n\n{pesan_rekomendasi}"
      )
  else:
    st.success("### ✅ Semua stok persediaan berada dalam kondisi aman.")

  st.markdown("---")

  # --------------------------------------------------------
  # BAGIAN B: PANEL TRANSAKSI GUDANG
  # --------------------------------------------------------
  st.markdown("### 📥 📤 Panel Transaksi Gudang (Real-Time Mutasi)")
  col_kiri, col_tengah, col_kanan = st.columns(3)

  with col_kiri:
    pilih_bahan = st.selectbox(
        "Pilih Jenis Barang:", df_hasil["Bahan Baku"].tolist()
    )
  with col_tengah:
    jenis_transaksi = st.radio(
        "Jenis Aktivitas:",
        ["Barang Datang (Stok Masuk)", "Diambil Produksi (Stok Keluar)"],
    )
  with col_kanan:
    jumlah_mutasi = st.number_input("Jumlah (kg):", min_value=0.0, step=10.0)
    submit_button = st.button("Simpan Transaksi 💾")

  if submit_button and jumlah_mutasi > 0:
    if jenis_transaksi == "Barang Datang (Stok Masuk)":
      catat_transaksi(pilih_bahan, jenis_transaksi, jumlah_mutasi)
      st.session_state["stok_realtime"][pilih_bahan] += jumlah_mutasi
      st.toast(f"Berhasil mencatat barang masuk!", icon="📥")
    else:
      if st.session_state["stok_realtime"][pilih_bahan] - jumlah_mutasi < 0:
        st.error("Transaksi Ditolak! Stok di gudang tidak mencukupi.")
      else:
        catat_transaksi(pilih_bahan, jenis_transaksi, jumlah_mutasi)
        st.session_state["stok_realtime"][pilih_bahan] -= jumlah_mutasi
        st.toast(f"Berhasil mencatat pengeluaran produksi!", icon="📤")
    st.rerun()

  st.markdown("---")

  # --------------------------------------------------------
  # BAGIAN C: TAB MONITORING LENGKAP & HISTOGRAM DI BAWAH
  # --------------------------------------------------------
  tab1, tab2, tab3 = st.tabs([
      "📊 Live Monitoring & Chart",
      "📜 Laporan Riwayat Keluar-Masuk",
      "⚙️ Reset Data",
  ])

  with tab1:
    st.markdown("#### Kondisi Real-Time Seluruh Jenis Barang 📝")
    df_tampilan = df_hasil.copy()
    df_tampilan["Stok Saat Ini"] = df_tampilan["Stok Saat Ini (kg)"].apply(
        format_indonesia_kg
    )
    df_tampilan["s_opt (Batas Aman/r)"] = df_tampilan["s_opt (kg)"].apply(
        format_indonesia_kg
    )
    df_tampilan["Q_opt (Pemesanan Dasar)"] = df_tampilan["Q_opt (kg)"].apply(
        format_indonesia_kg
    )

    # REVISI 2: Tampilan S_max hanya diformat ke kg jika tidak None (Kategori A)
    df_tampilan["S_max (Target Maksimum)"] = df_tampilan["S_max (kg)"].apply(
        lambda x: format_indonesia_kg(x) if pd.notna(x) else "-"
    )

    df_tampilan["Safety Stock (SS)"] = df_tampilan["SS (kg)"].apply(
        format_indonesia_kg
    )
    df_tampilan["Total Cost (OT)"] = df_tampilan["Total Cost (OT)"].apply(
        format_rupiah
    )

    kolom_tampil = [
        "Bahan Baku",
        "Kategori",
        "Stok Saat Ini",
        "Status Gudang",
        "s_opt (Batas Aman/r)",
        "Q_opt (Pemesanan Dasar)",
        "S_max (Target Maksimum)",
        "Safety Stock (SS)",
        "Total Cost (OT)",
    ]
    st.dataframe(df_tampilan[kolom_tampil], use_container_width=True)

    st.markdown("---")

    # Penempatan grafik: Berada di bawah tabel
    st.markdown(
        "#### 📊 Histogram Pemantauan Batas Stok Aktual vs Parameter Persediaan"
    )

    df_grafik = df_hasil.copy()
    df_grafik["Stok Saat Ini (kg)"] = pd.to_numeric(
        df_grafik["Stok Saat Ini (kg)"], errors="coerce"
    ).fillna(0)
    df_grafik["s_opt (kg)"] = pd.to_numeric(
        df_grafik["s_opt (kg)"], errors="coerce"
    ).fillna(0)
    df_grafik["S_max (kg)"] = pd.to_numeric(
        df_grafik["S_max (kg)"], errors="coerce"
    ).fillna(0)

    fig = px.bar(
        df_grafik,
        x="Bahan Baku",
        y=["Stok Saat Ini (kg)", "s_opt (kg)", "S_max (kg)"],
        barmode="group",
        title=(
            "Komparasi Visual Stok Riil Terhadap Batas Aman (s) dan Target"
            " Maksimum (S)"
        ),
        labels={"value": "Volume (kg)", "variable": "Parameter Gudang"},
        color_discrete_sequence=["#3399FF", "#FF3333", "#00CC66"],
    )
    st.plotly_chart(fig, use_container_width=True)

  with tab2:
    st.markdown("### 📜 Log Laporan Mutasi Barang Gudang")
    df_log_tampil = muat_riwayat()

    if not df_log_tampil.empty:
      col_f1, col_f2 = st.columns(2)
      with col_f1:
        filter_hari = st.multiselect(
            "Filter Berdasarkan Tanggal (Hari):",
            sorted(df_log_tampil["Tanggal"].unique(), reverse=True),
        )
      with col_f2:
        filter_bulan = st.multiselect(
            "Filter Berdasarkan Bulan:",
            sorted(df_log_tampil["Bulan"].unique(), reverse=True),
        )

      df_filtered = df_log_tampil.copy()
      if filter_hari:
        df_filtered = df_filtered[df_filtered["Tanggal"].isin(filter_hari)]
      if filter_bulan:
        df_filtered = df_filtered[df_filtered["Bulan"].isin(filter_bulan)]

      df_filtered = df_filtered.sort_values(by="Waktu", ascending=False)
      df_filtered["Jumlah (kg)"] = df_filtered["Jumlah (kg)"].apply(
          format_indonesia_kg
      )

      st.dataframe(
          df_filtered[["Waktu", "Bahan Baku", "Aktivitas", "Jumlah (kg)"]],
          use_container_width=True,
      )

      st.download_button(
          label="📥 Unduh Laporan Riwayat Lengkap (.xlsx)",
          data=open(FILE_RIWAYAT, "rb").read(),
          file_name=f"Laporan_Gudang_{datetime.now().strftime('%Y%m%d')}.xlsx",
          mime=(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          ),
      )
    else:
      st.info("Belum ada riwayat transaksi keluar-masuk barang yang tercatat.")

  with tab3:
    st.warning("Menghapus data akan membersihkan memori internal dashboard.")
    if st.button("Hapus & Reset Master Data Gudang 🗑️"):
      st.session_state["data_gudang"] = None
      st.session_state["stok_realtime"] = {}
      if os.path.exists(FILE_RIWAYAT):
        os.remove(FILE_RIWAYAT)
      st.rerun()

else:
  st.info(
      "💡 Silakan upload file Excel data Inventory Anda di panel sebelah kiri"
      " untuk Mulai Perhitungan & mengaktifkan sistem monitoring."
  )

  st.markdown("---")
  st.markdown("### 📋 Rincian Data Parameter yang Dibutuhkan System")
  st.markdown(
      "Agar sistem dapat menghitung Reorder Point ($s$), Ukuran Pemesanan ($Q$), dan"
      " Maximum Stock ($S$), file Excel Anda wajib memiliki kolom berikut:"
  )

  col_a, col_b = st.columns(2)

  with col_a:
    st.markdown("""
        **1. Demand ($D$) - Total Permintaan Tahunan**
        * **Apa itu?** Total jumlah pemakaian barang dalam 1 tahun (dalam kg/unit).
        * **Cara dapatnya:** Jumlahkan total pemakaian dari bulan Januari sampai Desember.

        **2. Standar Deviasi / Sigma ($\sigma$) - Fluktuasi Pemakaian**
        * **Apa itu?** Angka yang menunjukkan seberapa fluktuatif pemakaian barang tiap bulannya.
        * **Cara dapatnya:** Gunakan rumus Excel `=STDEV.S(data_12_bulan)*SQRT(12)`.

        **3. Lead Time ($L$) - Waktu Tunggu Pengiriman**
        * **Apa itu?** Durasi waktu sejak barang dipesan sampai tiba di gudang (dalam satuan TAHUN).
        * **Cara dapatnya:** Jika supplier butuh waktu 30 hari, maka $L = 30 / 365 = 0.0822$ tahun.
        """)

  with col_b:
    st.markdown("""
        **4. Biaya Pesan ($A$) - Ordering Cost**
        * **Apa itu?** Biaya yang keluar setiap kali Anda melakukan 1 kali pemesanan.
        * **Contoh:** Biaya administrasi, ongkos kirim, telepon, pemeriksaan barang.

        **5. Biaya Simpan ($h$) - Holding Cost**
        * **Apa itu?** Biaya untuk menyimpan 1 unit barang di gudang selama 1 TAHUN.
        * **Contoh:** Biaya sewa tempat, listrik gudang, perawatan, atau kerusakan per unit.

        **6. Biaya Kekurangan ($C_u$) - Shortage Cost**
        * **Apa itu?** Kerugian atau estimasi biaya denda/kehilangan profit jika persediaan gudang HABIS saat produksi berjalan.
        """)

  st.markdown("---")
  st.success(
      "💡 **Sudah Dijelaskan Tapi Masih Bingung wkwkwk. Tenang, Udah Disediakan"
      " Template Kosongan Kok 🥳** Klik tombol **'📥 Unduh Template Excel"
      " Otomatis'** di menu sebelah kiri untuk mengunduh template Excel siap"
      " isi!"
  )