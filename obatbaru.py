# Mengimpor pustaka utama Tkinter untuk membuat antarmuka grafis (GUI)
import tkinter as tk

# Mengimpor komponen modern ttk (untuk style), kotak pesan (alert), dan dialog pemilih file
from tkinter import ttk, messagebox, filedialog

# Mengimpor pustaka Pandas untuk membaca data dari file tabel (Excel/CSV/TXT)
import pandas as pd

# Mengimpor pustaka OS bawaan Python untuk interaksi dengan file sistem komputer
import os

class ApotekApp:
    def __init__(self, root):
        # Menyimpan objek window utama
        self.root = root
        
        # Mengatur teks judul pada window aplikasi
        self.root.title("Sistem Informasi Apotek")
        
        # Mengatur ukuran lebar (500px) dan tinggi (650px) window aplikasi agar lebih lega
        self.root.geometry("500x650")
        
        # Menyimpan kode warna tema hijau mint medis (Medical Mint) untuk latar belakang
        bg_medical = "#F0FDFA"
        
        # Menerapkan warna latar belakang pada window utama aplikasi
        self.root.configure(bg=bg_medical)

        # List untuk menampung nama obat dari file
        self.data_obat = []
        
        # --- KONFIGURASI STYLE TTK (Membuat Tampilan Modern & Elegan) ---
        # Membuat objek Style untuk memodifikasi komponen bawaan Tkinter
        self.style = ttk.Style()
        
        # Menggunakan tema 'clam' sebagai dasar agar warna tombol bisa diubah total
        self.style.theme_use("clam")
        
        # Mengatur warna latar belakang frame pembungkus agar menyatu dengan background utama
        self.style.configure("TFrame", background=bg_medical)
        
        # Mengatur teks label biasa menggunakan font Segoe UI dengan warna hijau tua pekat
        self.style.configure("TLabel", background=bg_medical, foreground="#115E59", font=("Segoe UI", 10))
        
        # Mengatur teks judul utama (Title) menjadi besar (ukuran 16) dan tebal (bold)
        self.style.configure("Title.TLabel", background=bg_medical, foreground="#134E4A", font=("Segoe UI", 16, "bold"))
        
        # Mengatur teks sub-judul (Subtitle) menjadi miring (italic) dengan warna hijau pudar
        self.style.configure("Subtitle.TLabel", background=bg_medical, foreground="#0D9488", font=("Segoe UI", 9, "italic"))

        # --- DESAIN TOMBOL MODERN DENGAN VARIASI WARNA (UX HIGHLIGHT) ---
        # Tombol Langkah 1: Warna Teal Medis
        self.style.configure("Load.TButton", font=("Segoe UI", 10, "bold"), foreground="white", background="#0D9488", borderwidth=0)
        # Efek Hover: Berubah menjadi teal lebih gelap saat kursor mouse menyentuh tombol
        self.style.map("Load.TButton", background=[("active", "#0F766E")])
        
        # Tombol Langkah 2: Warna Hijau Emerald Tua
        self.style.configure("Sort.TButton", font=("Segoe UI", 10, "bold"), foreground="white", background="#047857", borderwidth=0)
        # Efek Hover: Berubah menjadi hijau lebih gelap saat disentuh mouse
        self.style.map("Sort.TButton", background=[("active", "#065F46")])
        
        # Tombol Langkah 3: Warna Biru Cyan Medis
        self.style.configure("Search.TButton", font=("Segoe UI", 10, "bold"), foreground="white", background="#0284C7", borderwidth=0)
        # Efek Hover: Berubah menjadi biru lebih gelap saat disentuh mouse
        self.style.map("Search.TButton", background=[("active", "#0369A1")])

        # --- TATA LETAK / LAYOUT UI ---
        
        # 1. BAGIAN HEADER (Judul Aplikasi)
        # Membuat frame atas dengan jarak padding dalam sebesar 20px
        self.header_frame = ttk.Frame(root, padding=20)
        self.header_frame.pack(fill="x")
        
        # Memasang teks judul utama di area header
        ttk.Label(self.header_frame, text="Sistem Informasi Apotek Wahab", style="Title.TLabel").pack(anchor="w")
        # Memasang teks keterangan di bawah judul utama
        ttk.Label(self.header_frame, text="Manajemen, Pengurutan, dan Pencarian Data Obat", style="Subtitle.TLabel").pack(anchor="w", pady=(2, 0))
        
        # Membuat garis pembatas horizontal (divider) berwarna hijau mint terang
        divider = tk.Frame(root, height=1, bg="#CCFBF1")
        divider.pack(fill="x", padx=20)

        # BAGIAN UTAMA (Main Container)
        # Membuat frame penampung konten utama dengan jarak padding 20px
        self.main_frame = ttk.Frame(root, padding=20)
        self.main_frame.pack(fill="both", expand=True)

        # 2. BAGIAN KONTROL FILE (Langkah 1)
        # Membuat label petunjuk untuk memuat file master obat
        ttk.Label(self.main_frame, text="Langkah 1: Ambil Data Master", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        
        # Membuat tombol modern untuk memicu fungsi load_file (Kursor berubah jadi icon tangan saat di-hover)
        self.btn_load = ttk.Button(self.main_frame, text="📂 Pilih File Obat (.xlsx, .csv, .txt)", 
                                   command=self.load_file, style="Load.TButton", cursor="hand2")
        self.btn_load.pack(fill="x", pady=(5, 15))

        # 3. BAGIAN DAFTAR DATA & SORTIR (Langkah 2)
        # Membuat label petunjuk untuk daftar data dan sorting
        ttk.Label(self.main_frame, text="Langkah 2: Tampilan & Pengurutan", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        
        # Membuat frame khusus untuk membungkus listbox agar memiliki garis pinggir (border) hitam tipis yang elegan
        self.listbox_frame = tk.Frame(self.main_frame, bg="#FFFFFF", bd=1, relief="solid", highlightbackground="#99F6E4")
        self.listbox_frame.pack(fill="both", expand=True, pady=(5, 5))
        
        # Membuat komponen Scrollbar vertikal agar daftar data bisa digulir ke atas-bawah dengan mulus
        self.scrollbar = ttk.Scrollbar(self.listbox_frame, orient="vertical")
        self.scrollbar.pack(side="right", fill="y")

        # Membuat komponen Listbox berwarna putih bersih dengan teks hijau tua, tanpa border bawaan yang kaku
        self.listbox = tk.Listbox(self.listbox_frame, height=8, font=("Segoe UI", 10), 
                                  bg="#FFFFFF", fg="#115E59", selectbackground="#0D9488", 
                                  selectforeground="white", borderwidth=0, highlightthickness=0,
                                  yscrollcommand=self.scrollbar.set)
        self.listbox.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        # Menghubungkan scrollbar dengan pergerakan vertikal listbox
        self.scrollbar.config(command=self.listbox.yview)

        # Membuat label status dinamis untuk memantau apakah data sudah siap, belum dimuat, atau sudah disorting
        self.lbl_status = ttk.Label(self.main_frame, text="Status: Data belum dimuat", style="Subtitle.TLabel")
        self.lbl_status.pack(anchor="w", pady=(0, 5))

        # Membuat tombol modern untuk memicu fungsi sort_data (Bubble Sort)
        self.btn_sort = ttk.Button(self.main_frame, text="⚡ Urutkan Data (Bubble Sort)", 
                                   command=self.sort_data, style="Sort.TButton", cursor="hand2")
        self.btn_sort.pack(fill="x", pady=(0, 15))

        # 4. BAGIAN PENCARIAN DATA (Langkah 3)
        # Membuat label petunjuk untuk pencarian data obat
        ttk.Label(self.main_frame, text="Langkah 3: Pencarian Cepat", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        
        # Membuat kolom input teks (Entry) modern dengan border tipis, warna teks hijau tua, dan ruang ketik yang lega
        self.entry_cari = tk.Entry(self.main_frame, font=("Segoe UI", 11), bg="#FFFFFF", 
                                   fg="#115E59", bd=1, relief="solid", insertbackground="#115E59")
        self.entry_cari.pack(fill="x", ipady=6, pady=(5, 5))
        
        # Menghubungkan tombol 'Enter' pada keyboard langsung ke fungsi search_data (Peningkatan UX yang cerdas)
        self.entry_cari.bind("<Return>", lambda event: self.search_data())

        # Membuat tombol modern untuk memicu fungsi search_data (Binary Search)
        self.btn_cari = ttk.Button(self.main_frame, text="🔍 Cari Nama Obat (Binary Search)", 
                                   command=self.search_data, style="Search.TButton", cursor="hand2")
        self.btn_cari.pack(fill="x", pady=(0, 10))

    def update_listbox(self):
        # Menghapus seluruh item lama yang ada di listbox
        self.listbox.delete(0, tk.END)
        # Memasukkan ulang semua data obat terbaru ke listbox dengan tambahan penomoran otomatis di depannya (1., 2., 3.)
        for idx, item in enumerate(self.data_obat, start=1):
            self.listbox.insert(tk.END, f" {idx}.  {item}")

    def load_file(self):
        # Membuka jendela dialog komputer untuk memilih file data
        file_path = filedialog.askopenfilename(
            filetypes=[("All files", "*.*"), ("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv"), ("Text files", "*.txt")],
        )
        
        # Memeriksa jika pengguna benar-benar memilih file
        if file_path:
            try:
                # Membaca data jika format file adalah CSV
                if file_path.endswith('.csv'):
                    df = pd.read_csv(file_path)
                # Membaca data jika format file adalah TXT
                elif file_path.endswith('.txt'):
                    df = pd.read_csv(file_path, delimiter='\t')
                # Menolak data jika format file adalah PDF dengan pesan error yang rapi
                elif file_path.endswith('.pdf'):
                    messagebox.showerror("Format Salah", "Format PDF tidak didukung untuk pengolahan tabel data!")
                    return
                # Membaca data jika format file adalah Excel (xlsx/xls)
                else:
                    df = pd.read_excel(file_path)
                
                # Mengambil kolom pertama, hapus baris kosong, ubah ke list string
                self.data_obat = df.iloc[:, 0].dropna().astype(str).tolist()
                
                # Memperbarui tampilan daftar obat di layar listbox
                self.update_listbox()
                
                # Mengubah teks label status menjadi warna oranye untuk mengingatkan pengguna bahwa data belum diurutkan
                self.lbl_status.config(text=f"Status: Berhasil memuat {len(self.data_obat)} data (Belum diurutkan)", foreground="#EA580C")
                
                # Menampilkan pesan sukses memuat data obat
                messagebox.showinfo("Sukses", f"Berhasil memuat {len(self.data_obat)} data obat.")
            except Exception as e:
                # Menangkap error jika file gagal atau rusak saat dibaca
                messagebox.showerror("Gagal Membaca", f"Sistem gagal membaca file:\n{e}")

    def sort_data(self):
        # Validasi mencegah error jika data di memori masih kosong
        if not self.data_obat:
            messagebox.showwarning("Peringatan", "Data masih kosong! Silakan pilih file terlebih dahulu pada Step 1.")
            return

        # Mengambil jumlah total data obat yang tersimpan
        n = len(self.data_obat)
        
        # Perulangan luar untuk mengontrol jumlah putaran Bubble Sort
        for i in range(n):
            # Perulangan dalam untuk membandingkan elemen bersebelahan
            for j in range(0, n - i - 1):
                # Membandingkan abjad nama obat secara case-insensitive (mengabaikan huruf besar/kecil)
                if self.data_obat[j].lower() > self.data_obat[j + 1].lower():
                    # Menukar posisi data jika data kiri alfabetnya lebih besar dari data kanan
                    self.data_obat[j], self.data_obat[j + 1] = self.data_obat[j + 1], self.data_obat[j]
        
        # Memperbarui urutan nama obat terbaru di layar listbox setelah disorting
        self.update_listbox()
        
        # Mengubah teks label status menjadi warna hijau sukses karena data sudah rapi dari A-Z
        self.lbl_status.config(text=f"Status: {len(self.data_obat)} Data telah diurutkan A-Z (Siap dicari)", foreground="#16A34A")
        
        # Menampilkan pesan sukses pengurutan data
        messagebox.showinfo("Sukses", "Data obat berhasil diurutkan secara Alfabet (A-Z)!")

    def search_data(self):
        # Mengambil teks nama obat yang diketik pengguna dan menghapus spasi di ujungnya (.strip)
        target = self.entry_cari.get().strip()
        
        # Validasi jika kolom input pencarian masih kosong
        if not target:
            messagebox.showwarning("Peringatan", "Silakan ketik nama obat yang ingin dicari!")
            return
        # Validasi mencegah error jika data obat di memori masih kosong
        if not self.data_obat:
            messagebox.showwarning("Peringatan", "Data obat belum tersedia di sistem. Muat file terlebih dahulu.")
            return

        # Inisialisasi pointer indeks awal (kiri) dan akhir (kanan) untuk algoritma Binary Search
        kiri, kanan = 0, len(self.data_obat) - 1
        
        # Variabel penanda status penemuan data obat
        ditemukan = False
        
        # Perulangan pencarian selama rentang batas kiri tidak melewati batas kanan
        while kiri <= kanan:
            # Menghitung letak titik tengah rentang pencarian aktif
            tengah = (kiri + kanan) // 2
            
            # Kondisi jika nama obat di titik tengah cocok dengan target kata kunci (case-insensitive)
            if self.data_obat[tengah].lower() == target.lower():
                # --- FITUR UX TAMBAHAN (OTOMATIS MENYOROT BARIS DATA DI SCREEN) ---
                # Menggulung otomatis listbox ke posisi data yang ditemukan
                self.listbox.see(tengah)
                # Menghapus seleksi biru lama di listbox
                self.listbox.selection_clear(0, tk.END)
                # Menyorot baris data obat yang ditemukan dengan warna seleksi biru baru
                self.listbox.selection_set(tengah)
                # Mengaktifkan fokus baris tersebut di listbox
                self.listbox.activate(tengah)
                
                # Menampilkan info sukses penemuan data beserta posisinya (indeks ditambah 1 agar mudah dibaca manusia)
                messagebox.showinfo("Hasil Pencarian", f"Obat '{target}' BERHASIL DITEMUKAN!\n\nBerada pada Urutan ke-{tengah + 1} di dalam tabel master.")
                ditemukan = True
                # Menghentikan paksa perulangan karena data sudah ketemu
                break
            
            # Kondisi jika abjad data tengah ternyata lebih kecil dari target kata kunci
            elif self.data_obat[tengah].lower() < target.lower():
                # Memperkecil wilayah pencarian dan fokus menggeser pointer ke area sebelah kanan
                kiri = tengah + 1
            
            # Kondisi jika abjad data tengah ternyata lebih besar dari target kata kunci
            else:
                # Memperkecil wilayah pencarian dan fokus menggeser pointer ke area sebelah kiri
                kanan = tengah - 1
        
        # Menampilkan pesan error jika sampai akhir perulangan data tidak ditemukan di list
        if not ditemukan:
            messagebox.showerror("Tidak Ditemukan", f"Obat dengan nama '{target}' tidak dapat ditemukan dalam basis data saat ini.\n\nPastikan Anda sudah menekan tombol 'Urutkan Data' sebelum mencari.")

if __name__ == "__main__":
    # Membuat instance utama aplikasi GUI tkinter
    root = tk.Tk()
    
    # Menghubungkan window dasar ke kelas logika program
    app = ApotekApp(root)
    
    # Menjalankan loop utama sistem agar window tetap terbuka di layar komputer
    root.mainloop()