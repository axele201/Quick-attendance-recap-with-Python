import pandas as pd

# Membaca file absensi yang baru
file_absensi = 'absensi.xlsx'
df_absensi = pd.read_excel(file_absensi)

# Membaca file rekap absensi yang sudah ada (atau membuat baru jika tidak ada)
file_rekap = 'rekap_absensi.xlsx'
try:
    df_rekap = pd.read_excel(file_rekap)
except FileNotFoundError:
    df_rekap = pd.DataFrame(columns=['Nama', 'Clock In', 'Clock Out'])

# Memilih hanya kolom 'Nama', 'Clock In', dan 'Clock Out' dari file absensi
df_absensi_selected = df_absensi[['Nama', 'Clock In', 'Clock Out']]

# Mengidentifikasi nama yang ada di absensi tetapi tidak ada di rekap absensi
for index, row in df_absensi_selected.iterrows():
    nama = row['Nama']
    clock_in = row['Clock In']
    clock_out = row['Clock Out']
    
    # Mengecek apakah nama sudah ada di rekap_absensi
    if nama in df_rekap['Nama'].values:
        # Jika sudah ada, update Clock In dan Clock Out
        df_rekap.loc[df_rekap['Nama'] == nama, 'Clock In'] = clock_in
        df_rekap.loc[df_rekap['Nama'] == nama, 'Clock Out'] = clock_out
    else:
        # Jika nama belum ada, tambahkan baris baru menggunakan pd.concat
        new_row = pd.DataFrame([row])
        df_rekap = pd.concat([df_rekap, new_row], ignore_index=True)

# Menyimpan kembali file rekap absensi
df_rekap.to_excel(file_rekap, index=False)

print("Data absensi berhasil ditambahkan atau diperbarui di file rekap_absensi.xlsx")
