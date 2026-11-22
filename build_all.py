# -*- coding: utf-8 -*-
"""
Script untuk membangun seluruh berkas buku "ATURAN 80:20 MEDIA SOSIAL":
1. Bab-bab Markdown terpisah di folder chapters/
2. Naskah lengkap di BUKU-LENGKAP.md
3. Pembaca Web iPad-Friendly di index.html
4. Dokumen Word (.docx) via python-docx
5. Dokumen PDF (.pdf) via reportlab dengan font besar yang nyaman untuk iPad
6. README.md untuk repositori GitHub
"""

import os
import sys
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

chapters_data = [
    {
        "id": "bab-1",
        "file": "01-paradoks-kreator.md",
        "title": "Bab 1: Paradoks Kreator — Mengapa Konsistensi Saja Tidak Cukup",
        "subtitle": "Menghentikan Kebiasaan Membunuh Format Konten yang Baru Saja Berhasil",
        "content": """## Realitas Pahit di Balik Tombol 'Posting'

Sebagian besar kreator dan pemilik bisnis digital beroperasi di bawah ilusi yang berbahaya: *"Jika saya memposting konten setiap hari secara konsisten, akun saya pada akhirnya pasti akan meledak."*

Kenyataannya tidak demikian.

Konsisten memposting konten yang tidak diinginkan oleh pasar hanya akan menghasilkan satu hal: Anda menjadi orang yang gagal dengan sangat disiplin.

Pasar media sosial tidak memberikan kompensasi kepada keringat, durasi rekaman, atau seberapa keras Anda berjuang. Pasar hanya memedulikan satu hal sederhana: **apakah konten ini memberikan nilai spesifik yang membuat orang bersedia menukarkan waktu dan perhatian mereka?**

---

## Penyakit 'Creative ADD': Membunuh Format Pemenang

Amati pola yang berulang pada 95% kreator pemula:

1. Mereka mencoba 10 topik yang berbeda secara acak: tips produktivitas, kutipan motivasi, cerita keseharian, tutorial teknis, dan bedah website.
2. Tiba-tiba, video nomor 5 tentang **"Bedah Website"** meledak dan mendapatkan 50.000 penayangan serta ratusan komentar.
3. Alih-alih membuat 50 video bedah website berikutnya, keesokan harinya kreator tersebut memposting video tentang: *"Rutinitas pagi saya minum kopi di kafe."*

Hasilnya? Video kopi tersebut hanya ditonton oleh 200 orang.

Ketika ditanya mengapa mereka tidak melanjutkan format bedah website, jawaban yang paling umum adalah:
> *"Saya takut audiens saya bosan kalau saya bahas website terus."*

Ini adalah kesalahan berpikir yang fatal. 

Audiens Anda tidak bosan. **Anda yang bosan.**

Audiens yang baru saja mengikuti akun Anda setelah menonton video bedah website mengharapkan lebih banyak bedah website. Mereka bahkan belum mengenal nama Anda dengan baik. Mereka datang untuk format dan nilai spesifik tersebut. 

Saat Anda mengganti topik secara acak, Anda memutuskan rantai momentum yang baru saja dibangun oleh algoritma.

---

## Media Sebagai Pengungkit Tanpa Izin

Dalam ekonomi digital modern, media adalah salah satu bentuk pengungkit (*leverage*) terbesar yang bisa diakses oleh siapa saja tanpa perlu meminta izin kepada siapa pun.

- Anda tidak perlu mempekerjakan ratusan staf untuk berbicara kepada 10.000 calon pembeli.
- Anda tidak perlu menyewa papan reklame di tengah kota seharga ratusan juta rupiah.
- Sebuah video atau tulisan yang tepat diunggah ke internet dapat bekerja untuk Anda selama 24 jam sehari, 7 hari seminggu, menjangkau orang-orang yang membutuhkan keahlian Anda saat Anda sedang tidur.

Namun, pengungkit ini hanya bekerja jika Anda memberikan **sinyal yang konsisten dan jelas** kepada dua pihak:
1. **Algoritma rekomendasi**, agar platform memahami kepada siapa konten Anda harus disalurkan.
2. **Penonton**, agar mereka memahami dalam kategori apa mereka harus mengingat nama dan keahlian Anda.

Solusinya bukan berhenti bereksperimen, melainkan mengelola eksperimen dengan batas matematika yang ketat. Inilah fondasi dari **Strategi 80:20 Media Sosial**."""
    },
    {
        "id": "bab-2",
        "file": "02-prinsip-dasar-80-20.md",
        "title": "Bab 2: Prinsip 80:20 — Eksploitasi Format Juara vs Eksplorasi Asimetris",
        "subtitle": "Arsitektur Distribusi Konten yang Menggabungkan Skala dan Inovasi",
        "content": """## Algoritma Adalah Mesin Pengenal Pola

Setiap platform media sosial modern—baik itu TikTok, Instagram, YouTube Shorts, LinkedIn, maupun X—bekerja dengan prinsip yang sama: **mencocokkan konten yang tepat dengan penonton yang tepat demi mempertahankan retensi pengguna.**

Algoritma bukan kurator seni yang mencari keberagaman estetika Anda. Algoritma adalah sistem pemrosesan data statistik:

1. Ketika video Anda tentang topik tertentu mencatatkan waktu tonton (*watch time*) yang tinggi dan interaksi aktif, sistem menandai akun Anda sebagai penyedia konten berkualitas dalam topik tersebut.
2. Jika konten berikutnya menyajikan topik dan format yang serupa, sistem memiliki tingkat kepastian (*confidence score*) yang tinggi untuk kembali mendistribusikannya ke kelompok penonton yang serupa.
3. Sebaliknya, jika Anda melompat dari topik desain website ke tips investasi saham lalu ke vlog kuliner, sistem kebingungan menentukan target audiens Anda, dan penyebaran konten Anda terhenti.

---

## Mengapa Harus 80% vs 20%?

Banyak praktisi terjebak di salah satu dari dua kutub ekstrem:
- **Kutub Kebosanan Total (100% Format Sama Tanpa Variasi):** Akun menjadi kaku, kreator jenuh, tidak ada ruang untuk menjual jasa, dan jika tren format tersebut meredup, akun ikut mati.
- **Kutub Kekacauan Kreatif (100% Eksperimen Acak):** Akun tidak pernah menemukan daya dorong karena tidak ada satu format pun yang didalami hingga matang.

Aturan 80:20 memecahkan kebuntuan ini dengan membagi beban kerja ke dalam dua pilar yang saling mendukung:

| Komponen | Porsi | Peran Utama | Fokus Eksekusi |
| :--- | :---: | :--- | :--- |
| **Winning Engine** | **80%** | Eksploitasi Format Pemenang | Menghasilkan jangkauan luas (*reach*), pengikut baru, dan domain authority. |
| **Asymmetric Sandbox** | **20%** | Eksplorasi & Monetisasi | Menghasilkan penjualan jasa, membangun kedekatan emosional, dan riset format baru. |

```
               KOMPOSISI KONTEN (100 KONTEN)
  ┌─────────────────────────────────────────────────┬───────────┐
  │                                                 │           │
  │           80% FORMAT PEMENANG TERBUKTI          │  20% LAB  │
  │      (Skala, Pengikut Baru, Domain Authority)   │ (Jual/Exp)│
  │                                                 │           │
  └─────────────────────────────────────────────────┴───────────┘
```

---

## Asimetri Risiko dan Peluang

Mengapa porsi eksperimen dibatasi hanya 20%?
Karena ini menciptakan struktur risiko asimetris:
- **Kerugian Terbatas (*Capped Downside*):** Jika sebuah konten eksperimen atau video penawaran jasa gagal mendapatkan penayangan tinggi, performa akun Anda secara keseluruhan tetap aman karena 80% konten lainnya terus memompa jangkauan dan pengikut baru.
- **Peluang Tak Terbatas (*Unlimited Upside*):** Jika satu eksperimen dari porsi 20% ternyata berhasil meledak melebihi format utama, Anda baru saja menemukan formula juara baru yang siap dipromosikan ke pilar 80% di masa depan.

Ini adalah cara kerja investor modal ventura yang diterapkan ke dalam jadwal publikasi media sosial."""
    },
    {
        "id": "bab-3",
        "file": "03-pilar-80-winning-engine.md",
        "title": "Bab 3: Fondasi 80% — Memeras Format Pemenang (Studi Kasus: Diskusi & Bedah Website)",
        "subtitle": "Cara Memproduksi 80 Konten Berkualitas Tinggi Tanpa Pernah Kehabisan Ide",
        "content": """## Anatomi Format Juara

Bayangkan Anda menemukan bahwa video bertema **"Diskusi & Bedah Website"** adalah format yang berhasil mencuri perhatian audiens. Video tersebut memiliki retensi tinggi, komentar berisi diskusi teknis, dan mendatangkan banyak pengikut baru.

Pertanyaannya: *Bagaimana cara membuat 80 konten berikutnya yang semuanya membahas bedah website tanpa terdengar mengulang hal yang sama persis?*

Jawabannya adalah **Metode Variasi Variabel Tunggal**:
> **Pertahankan hook dan struktur kerangka kerja tetap konsisten, ubah hanya subjek atau sudut pandang bedahannya.**

Mari kita pecah 80 konten tersebut menjadi 4 kelompok taktis, masing-masing berisi 20 konten.

---

## 4 Kelompok Taktis (20 Konten per Kelompok)

### 1. Authority Piggybacking: Bedah Website Brand / Figur Terkenal (20 Konten)
Manfaatkan rasa ingin tahu audiens terhadap nama-nama besar yang sudah mereka kenal.
- *Contoh 1-5:* Bedah psikologi tombol beli pada website Apple, Tesla, Nike, Tokopedia, dan Airbnb.
- *Contoh 6-10:* Analisis alasan mengapa landing page startup terkemuka terlihat sangat minimalis tapi menghasilkan konversi tinggi.
- *Contoh 11-20:* Bedah halaman pribadi figur publik dan kreator internasional teratas.
- **Nilai untuk Audiens:** Mempelajari strategi raksasa industri yang bisa diterapkan ke skala bisnis mereka sendiri.

### 2. Fatal Mistakes & Solusi Praktis (20 Konten)
Fokus pada rasa takut kehilangan uang (*loss aversion*), yaitu alasan utama mengapa pemilik bisnis mencari solusi website.
- *Contoh 21-25:* "3 Kesalahan fatal pada website UMKM yang membuat biaya iklan Anda terbuang percuma."
- *Contoh 26-30:* "Mengapa loading website lebih dari 3 detik membunuh 40% calon pembeli Anda (dan cara memperbaikinya hari ini)."
- *Contoh 31-40:* Bedah formulir checkout yang terlalu panjang, navigasi mobile yang membingungkan, dan tipografi yang tidak terbaca.
- **Nilai untuk Audiens:** Solusi cepat yang bisa langsung mereka perbaiki dalam hitungan menit.

### 3. Micro-Element Teardown: Analisis Komponen Spesifik (20 Konten)
Daripada membedah seluruh website, fokuskan perhatian pada satu elemen mikro secara mendalam.
- *Contoh 41-50:* Analisis headline di atas lipatan layar (*above the fold*). Mengapa kalimat pertama menentukan 80% keputusan pengunjung untuk tetap tinggal atau pergi.
- *Contoh 51-60:* Analisis warna tombol CTA, posisi testimoni, perbandingan harga (*pricing table*), dan tata letak footer.
- **Nilai untuk Audiens:** Wawasan teknis mendalam yang membuktikan bahwa Anda adalah seorang pakar di bidang detail.

### 4. Live Before & After Redesign (20 Konten)
Tunjukkan bukti keahlian nyata dalam bentuk transformasi visual langsung.
- *Contoh 61-70:* Mengambil website bisnis lokal yang berantakan, lalu mendesain ulang halaman depannya secara live dalam video 60 detik.
- *Contoh 71-80:* Transformasi landing page jasa konsultasi dari tampilan kaku menjadi tampilan modern berorientasi konversi tinggi.
- **Nilai untuk Audiens:** Bukti visual tanpa bantahan (*irrefutable proof*) bahwa Anda benar-benar mampu mengeksekusi apa yang Anda bicarakan.

---

## Template Hook Berdaya Pikat Tinggi untuk Format 80%

Gunakan struktur pembuka yang tajam pada 3 detik pertama:

1. **The Negative Anchor:**
   > *"Jika website Anda masih meletakkan tombol ini di bagian atas, Anda sedang membuang 50% calon pembeli setiap hari. Mari kita bedah alasannya."*
2. **The Contrast Teardown:**
   > *"Website seharga 100 juta ini melakukan satu kesalahan konyol yang bahkan tidak dilakukan oleh blog gratisan. Lihat bagian ini."*
3. **The Data-Driven Mystery:**
   > *"Website ini berhasil menaikkan penjualan 35% hanya dengan mengubah satu kalimat pada halaman utama. Ini bedahannya."*

Dengan formula ini, 80 konten dapat diproduksi secara terstruktur tanpa perlu menebak-nebak apa yang harus dibahas setiap pagi."""
    },
    {
        "id": "bab-4",
        "file": "04-pilar-20-asymmetric-sandbox.md",
        "title": "Bab 4: Laboratorium 20% — Monetisasi Jasa, Portofolio & Eksperimen Liar",
        "subtitle": "Mengubah Penonton Menjadi Pembeli dan Menemukan Terobosan Format Berikutnya",
        "content": """## Mengapa 80% Saja Bisa Membunuh Akun Anda?

Jika Anda hanya memposting 100% video bedah website selamanya tanpa pernah melakukan hal lain, dua masalah besar akan muncul:
1. **Audiens melihat Anda sebagai ensiklopedia gratis, bukan sebagai penyedia jasa profesional yang bisa mereka sewa.**
2. **Anda tidak pernah tahu apakah ada format lain yang berpotensi menghasilkan dampak 10 kali lebih besar bagi bisnis Anda.**

Porsi 20% adalah ruang perlindungan dan pertumbuhan strategis akun Anda. Dalam siklus 100 konten, 20 konten ini didistribusikan ke dalam 3 fungsi krusial:

```
                  PEMBAGIAN PORSI 20% (DARI 100 KONTEN)
  ┌────────────────────────┬────────────────────────┬──────────────────────┐
  │   8 KONTEN MONETISASI  │     6 KONTEN VLOG      │  6 KONTEN EKSPERIMEN │
  │   Penawaran Jasa &     │   Personal Brand &     │   Format Liar &      │
  │   Studi Kasus Klien    │   Filosofi Kerja       │   Uji Sudut Pandang  │
  └────────────────────────┴────────────────────────┴──────────────────────┘
```

---

## 1. Monetisasi Jasa & Studi Kasus (8 Konten)

Banyak kreator merasa bersalah atau malu saat harus menjual jasa mereka. Mereka merasa penonton akan marah jika disuguhi konten promosi.

Penonton hanya membenci promosi yang buruk: promosi yang memohon-mohon atau promosi yang tidak menawarkan nilai.

Cara terbaik menjual jasa di media sosial adalah **menjual melalui bukti hasil kerja (*results-based selling*)**:

### Formula Konten Penawaran Jasa:
1. **Masalah Nyata Klien:** Ceritakan klien yang datang dengan masalah website yang sepi pembeli meski traffic iklannya tinggi.
2. **Diagnosa Masalah:** Tunjukkan apa yang salah pada website lama mereka (menghubungkannya dengan format 80% yang biasa Anda bahas).
3. **Solusi & Hasil:** Tunjukkan perubahan desain yang Anda lakukan dan kenaikan metrik bisnis yang dicapai klien.
4. **Call to Action (CTA) Jelas:** *"Jika Anda pemilik bisnis dengan traffic di atas 1.000 pengunjung per bulan dan ingin website Anda dikonversi lebih tinggi, kirim pesan 'AUDIT' ke DM saya untuk mendapatkan sesi review gratis selama 15 menit."*

Konten seperti ini tidak terasa seperti iklan yang mengganggu; konten ini terasa seperti pelajaran berharga yang diakhiri dengan solusi bagi mereka yang membutuhkan.

---

## 2. Personal Brand, Vlog & Filosofi Kerja (6 Konten)

Orang bertransaksi dengan manusia, bukan dengan logo perusahaan atau akun tanpa wajah.

Gunakan 6 konten ini untuk membagikan:
- **Di Balik Layar Kerja:** Bagaimana setup meja kerja Anda, software yang Anda gunakan, dan alur kerja harian Anda dalam menyelesaikan proyek klien.
- **Pelajaran dari Kegagalan:** Cerita tentang kesalahan terbesar yang pernah Anda buat di masa lalu dan apa yang Anda pelajari darinya.
- **Prinsip Kerja:** Mengapa Anda menolak jenis proyek tertentu demi menjaga kualitas hasil kerja.

Konten ini membangun kepercayaan emosional mendalam (*deep trust*) yang tidak bisa ditiru oleh kompetitor.

---

## 3. Eksperimen Format Liar (6 Konten)

Gunakan 6 konten ini untuk mencoba format yang sama sekali baru:
- Mencoba format podcast monolog 3 menit dengan mikrofon klasik.
- Membuat konten karosel infografis visual beresolusi tinggi.
- Wawancara singkat dengan pemilik bisnis offline tentang tantangan digital mereka.

Sebagian besar eksperimen ini mungkin menghasilkan performa biasa saja. Tetapi jika salah satunya mencatatkan lonjakan metrik di luar dugaan, Anda baru saja menemukan calon format juara berikutnya."""
    },
    {
        "id": "bab-5",
        "file": "05-the-promotion-loop.md",
        "title": "Bab 5: Siklus Promosi — Mengubah Eksperimen 20% Menjadi Format Juara Baru",
        "subtitle": "Membangun Mesin Pertumbuhan Konten yang Tidak Pernah Usang",
        "content": """## Kurva Kejenuhan Konten (*Content Decay Curve*)

Tidak ada format konten yang bisa bertahan selamanya dengan performa puncak.

Bahkan format bedah website terbaik pada akhirnya akan mengalami penurunan efektivitas setelah 6 hingga 12 bulan. Penonton lama sudah terbiasa dengan gaya penyampaian Anda, dan algoritma mulai menuntut variasi baru untuk mempertahankan durasi tonton pengguna.

Kreator amatir panik ketika format andalan mereka mulai melambat. Mereka merasa algoritma sedang menghukum mereka.

Kreator profesional menyikapi ini dengan tenang karena mereka telah menjalankan porsi 20% secara konsisten sejak hari pertama.

---

## Protokol Promosi Format Konten

Kapan sebuah konten dari laboratorium 20% berhak dipromosikan menjadi pilar 80% yang baru?

Gunakan checklist metrik evaluasi objektif berikut:

1. **Metrik Retensi:** Apakah video eksperimen tersebut memiliki rasio penyelesaian (*completion rate*) di atas 60% pada 30 detik pertama?
2. **Metrik Pembagian (*Share Ratio*):** Apakah rasio pembagian konten terhadap penayangan mencapai minimal 2-3%? (Tanda bahwa konten tersebut memiliki nilai sosial tinggi).
3. **Pengikut Baru per Penayangan:** Apakah eksperimen tersebut menghasilkan laju konversi pengikut baru yang setara atau lebih tinggi dibandingkan format 80% yang sedang berjalan?

```
                        FLYWHEEL REGENERASI KONTEN
                               ┌───────────────┐
                               │  Jalankan 80% │
                               │ Format Juara  │
                               └───────┬───────┘
                                       │ (Mulai Jenuh / Plateau)
                                       ▼
    ┌─────────────────┐        ┌───────────────┐
    │ Promosikan Jadi │ ◄──────┤  Evaluasi 20% │
    │ 80% Format Baru │ (Lolos)│  Eksperimen   │
    └─────────────────┘        └───────────────┘
```

---

## Contoh Transisi Mulus

Katakanlah dari 6 eksperimen liar di porsi 20%, Anda mencoba format: **"Wawancara Pemilik Bisnis & Bedah Finansial Website Mereka."**

Format ini ternyata mendapatkan antusiasme luar biasa. Banyak pemilik bisnis lain yang berkomentar ingin diwawancarai.

Langkah strategis Anda pada siklus 100 konten berikutnya:
- **Format 80% Baru:** Wawancara Pemilik Bisnis & Bedah Finansial Website (kini mengambil porsi utama).
- **Format 20% Baru:** Bedah website versi cepat (yang sebelumnya 80%), penawaran jasa, dan vlog kehidupan.

Dengan cara ini, akun Anda berevolusi secara organik mengikuti data pasar riil, bukan berdasarkan tebakan emosional."""
    },
    {
        "id": "bab-6",
        "file": "06-matematika-konten-100-post.md",
        "title": "Bab 6: Cetak Biru Eksekusi 100 Konten Pertama",
        "subtitle": "Jadwal Mingguan, Sistem Batching, dan Alur Kerja Tanpa Burnout",
        "content": """## Jadwal Distribusi Mingguan Berbasis Rasio 4:1

Untuk mengeksekusi rasio 80:20 dalam kehidupan sehari-hari, cara paling praktis adalah membaginya ke dalam siklus mingguan dengan pola **4 Konten Pilar : 1 Konten Eksperimen/Monetisasi**.

Jika Anda memposting 5 konten per minggu:

| Hari | Kategori Konten | Tipe Konten | Tujuan Strategis |
| :--- | :--- | :--- | :--- |
| **Senin** | 80% Pilar Juara | Bedah Website Brand Terkenal | Menarik traffic baru di awal pekan. |
| **Selasa** | 80% Pilar Juara | 3 Kesalahan Fatal Website UMKM | Membangun kredibilitas dan edukasi teknis. |
| **Rabu** | 80% Pilar Juara | Analisis Mikro: Headline Penjual | Memperdalam otoritas bidang spesifik. |
| **Kamis** | 80% Pilar Juara | Live Redesign Before & After | Bukti visual keahlian praktis. |
| **Jumat** | 20% Eksperimen / Jual | Studi Kasus Klien & Penawaran Jasa | Mengonversi penonton menjadi inquiry/klien. |
| **Sabtu** | 20% Personal / Lab | Vlog Di Balik Layar / Format Baru | Memperkuat hubungan emosional personal. |
| **Minggu** | Evaluasi & Istirahat | Analisis Data & Review Metrik | Meninjau performa dan menyiapkan materi pekan depan. |

---

## Sistem Produksi Batch (Satu Sesi untuk Satu Pekan)

Jangan pernah bangun di pagi hari lalu bertanya: *"Hari ini saya mau bikin konten apa ya?"*

Itu adalah resep paling pasti untuk mengalami kelelahan mental (*decision fatigue*) dan akhirnya berhenti memposting.

Terapkan alur kerja batching:

1. **Hari Riset & Skrip (2 Jam):**
   - Buka 10 website target.
   - Catat 3 poin penting per website menggunakan template hook yang sudah terbukti.
   - Buat kerangka skrip poin-poin sederhana.
2. **Hari Rekam Layar & Kamera (2 Jam):**
   - Rekam seluruh 4 video bedah website secara berurutan dalam satu kali duduk.
   - Rekam 1 video penawaran jasa/studi kasus.
3. **Penyuntingan & Penjadwalan (2 Jam):**
   - Tambahkan teks penjelas (*captions*), sorotan visual, dan potongan layar penting.
   - Masukkan seluruh materi ke dalam sistem penjadwalan otomatis platform.

Dengan 6 jam kerja terfokus per minggu, kebutuhan konten Anda selama satu pekan penuh sudah selesai dieksekusi dengan standar kualitas yang konsisten."""
    },
    {
        "id": "bab-7",
        "file": "07-25-aforisme-hukum-besi.md",
        "title": "Bab 7: 25 Aksioma & Hukum Besi Media Sosial",
        "subtitle": "Prinsip-Prinsip Pokok Pengungkit Digital untuk Dibaca Berulang Kali",
        "content": """## Kumpulan Prinsip Aksiomatik Pertumbuhan Digital

Simpan halaman ini. Buka kembali saat Anda merasa kehilangan arah atau tergoda untuk berhenti memposting format yang sedang bekerja.

1. **Pasar tidak berhutang apa pun pada usaha Anda.** Pasar hanya merespons nilai nyata yang relevan bagi mereka.
2. **Kebosanan kreator adalah pembunuh momentum nomor satu.** Jangan membunuh format yang disukai penonton hanya karena Anda bosan membuatnya.
3. **Media adalah mesin pengganda tanpa izin.** Satu konten berkualitas dapat berbicara kepada ribuan orang sekaligus saat Anda tertidur lelap.
4. **Algoritma adalah cermin dari penonton.** Jika konten Anda tidak didistribusikan, jangan salahkan platform; perbaiki hook dan nilai konten Anda.
5. **Konsistensi tanpa arah hanyalah kesia-siaan yang disiplin.** Lebih baik memposting 3 konten terarah per minggu daripada 14 konten acak tanpa tujuan.
6. **Eksploitasi 80% membangun kerajaan jangkauan Anda.** Ini adalah mesin yang mendatangkan orang-orang baru setiap detik.
7. **Eksplorasi 20% melindungi masa depan kerajaan tersebut.** Ini adalah laboratorium yang memastikan Anda tidak punah saat tren bergeser.
8. **Jual hasil, bukan proses kerja yang rumit.** Klien tidak membeli baris kode atau desain website; mereka membeli peningkatan kepercayaan dan omzet bisnis.
9. **Spesialisasi mengalahkan generalisasi di fase awal.** Menjadi orang nomor satu di bidang "Bedah Website UMKM" jauh lebih menguntungkan daripada menjadi "Kreator Konten Digital Serbabisa".
10. **Tiga detik pertama adalah gerbang penentu.** Jika Anda gagal meyakinkan penonton dalam 3 detik pertama, 57 detik berikutnya tidak pernah ada di mata mereka.
11. **Bukti visual membungkam segala keraguan.** Menunjukkan perbaikan desain sebelum dan sesudah jauh lebih meyakinkan daripada menulis esai 1.000 kata tentang teori desain.
12. **Ubah subjeknya, pertahankan struktur pemenangnya.** Jika format bedah website berhasil, bedah 80 website berbeda dengan struktur yang sama.
13. **Konversi membutuhkan instruksi yang gamblang.** Jangan berasumsi penonton tahu cara menyewa jasa Anda; beri tahu mereka dengan tepat langkah apa yang harus diambil.
14. **Retensi adalah metrik kebenaran.** Jumlah klik bisa dipancing dengan judul sensasional, tapi waktu tonton hanya bisa diraih dengan substansi yang solid.
15. **Akun yang tidak pernah menjual jasa adalah hobi yang mahal.** Jangan takut menawarkan keahlian Anda kepada mereka yang membutuhkan solusi nyata.
16. **Emosi menghubungkan, data meyakinkan.** Gunakan cerita personal di porsi 20% untuk membangun kedekatan emosional, dan gunakan data di porsi 80% untuk membangun otoritas.
17. **Kerugian Anda di porsi 20% selalu terbatas, tapi potensi keuntungannya tidak terbatas.** Ini adalah definisi taruhan asimetris terbaik.
18. **Format yang lelah bukan akhir dari segalanya.** Format lama melambat agar format baru yang telah Anda uji di laboratorium bisa naik panggung utama.
19. **Jangan menunggu inspirasi datang untuk mulai merekam.** Inspirasi adalah hasil sampingan dari proses kerja yang terstruktur, bukan syarat awalnya.
20. **Audiens baru tidak membaca arsip masa lalu Anda.** Hal yang terasa berulang bagi Anda adalah wawasan segar bagi orang yang baru saja mengikuti akun Anda hari ini.
21. **Kemudahan konsumsi menentukan laju distribusi.** Buat ide yang rumit menjadi analogi visual yang bisa dipahami dalam waktu kurang dari satu menit.
22. **Kritik dari orang yang tidak pernah memproduksi apa pun bernilai nol.** Dengarkan data perilaku penonton, abaikan komentar negatif tanpa substansi.
23. **Skalakan apa yang berhasil sampai batas maksimal.** Sebelum Anda mencoba strategi baru, pastikan strategi lama sudah diperas seluruh potensinya.
24. **Kepercayaan membutuhkan waktu berbulan-bulan untuk dibangun, tapi hanya butuh satu kebohongan untuk hancur.** Selalu tulus dalam memberikan rekomendasi dan penilaian.
25. **Mainkan permainan jangka panjang dengan orang-orang jangka panjang.** Bangun reputasi digital Anda di atas dasar keahlian yang nyata dan nilai yang terbukti."""
    }
]

def build_markdown_files():
    print("Membangun berkas Markdown...")
    os.makedirs("chapters", exist_ok=True)
    
    full_book_md = []
    full_book_md.append("# MESIN KONTEN ASIMETRIS: STRATEGI 80:20 MEDIA SOSIAL")
    full_book_md.append("### *Panduan Taktis Mengembangkan Akun, Membangun Otoritas, dan Memonetisasi Keahlian Tanpa Kehabisan Ide*\n")
    full_book_md.append("> **Format Baca Optimal:** iPad & Tablet Reader (Font Besar, Tata Letak Lapang)\n")
    full_book_md.append("---\n")
    full_book_md.append("## DAFTAR ISI\n")
    
    for ch in chapters_data:
        full_book_md.append(f"- [{ch['title']}](#{ch['id']})")
    full_book_md.append("\n---\n")
    
    for ch in chapters_data:
        # Tulis file bab individu
        ch_path = os.path.join("chapters", ch["file"])
        with open(ch_path, "w", encoding="utf-8") as f:
            f.write(f"# {ch['title']}\n")
            f.write(f"*{ch['subtitle']}*\n\n")
            f.write(ch['content'])
            f.write("\n")
        print(f"  [OK] Ditulis: {ch_path}")
        
        # Tambahkan ke buku lengkap
        full_book_md.append(f"<a id='{ch['id']}'></a>\n")
        full_book_md.append(f"# {ch['title']}\n")
        full_book_md.append(f"*{ch['subtitle']}*\n\n")
        full_book_md.append(ch['content'])
        full_book_md.append("\n\n---\n")
        
    with open("BUKU-LENGKAP.md", "w", encoding="utf-8") as f:
        f.write("\n".join(full_book_md))
    print("  [OK] Ditulis: BUKU-LENGKAP.md")

def build_readme():
    print("Membangun README.md...")
    readme_content = """# 📚 Mesin Konten Asimetris: Strategi 80:20 Media Sosial

> **Panduan Taktis Mengembangkan Akun, Membangun Otoritas, dan Memonetisasi Keahlian Tanpa Kehabisan Ide.**  
> *Diformat khusus agar nyaman dan mudah dibaca di layar iPad dan Tablet.*

---

## 🎯 Inti Strategi 80:20

Banyak kreator dan profesional digital gagal bertumbuh bukan karena kekurangan ide, melainkan karena **mengidap sindrom bosan sendiri**: saat satu format konten berhasil meledak, mereka justru meninggalkannya dan mencoba hal acak lainnya.

Buku ini mengajarkan arsitektur konten dua pilar:
- **Pilar 80% (Winning Engine):** Mengeksploitasi format yang sudah terbukti berhasil dengan volume terstruktur. Misalnya jika format "Diskusi & Bedah Website" terbukti mendatangkan penonton, buat 80 konten berikutnya seputar bedah website dengan 4 variasi sudut pandang taktis.
- **Pilar 20% (Asymmetric Sandbox):** Mengalokasikan 20% ruang untuk menjual jasa/portofolio nyata klien, membagikan vlog personal di balik layar, serta menguji eksperimen format baru yang berpotensi menjadi format pemenang berikutnya.

---

## 📖 Daftar Isi Lengkap

| Bab | Judul | Topik Bahasan |
| :---: | :--- | :--- |
| **Bab 1** | [Paradoks Kreator](chapters/01-paradoks-kreator.md) | Jebakan 'konsistensi buta' dan kebiasaan membunuh format pemenang. |
| **Bab 2** | [Prinsip Dasar 80:20](chapters/02-prinsip-dasar-80-20.md) | Algoritma sebagai pengenal pola; eksplorasi vs eksploitasi terukur. |
| **Bab 3** | [Fondasi 80%: The Winning Engine](chapters/03-pilar-80-winning-engine.md) | Bedah studi kasus website: memproduksi 80 variasi konten terstruktur. |
| **Bab 4** | [Laboratorium 20%: Sandbox](chapters/04-pilar-20-asymmetric-sandbox.md) | Cara menjual jasa lewat studi kasus, vlog personal, dan eksperimen liar. |
| **Bab 5** | [Siklus Promosi](chapters/05-the-promotion-loop.md) | Kapan dan bagaimana mempromosikan eksperimen 20% menjadi format 80% baru. |
| **Bab 6** | [Cetak Biru 100 Konten](chapters/06-matematika-konten-100-post.md) | Kalender mingguan 4:1, sistem produksi batching 6 jam per pekan. |
| **Bab 7** | [25 Aksioma & Hukum Besi](chapters/07-25-aforisme-hukum-besi.md) | 25 prinsip ringkas dan tajam tentang pengungkit media digital. |

---

## 📱 Format yang Tersedia

1. **Web Reader Interaktif (`index.html`):**  
   Dirancang khusus untuk iPad & Tablet! Dilengkapi tombol pengatur ukuran font (A- / A+), tema kertas/sepia/gelap, daftar isi samping, dan indikator progres baca. Buka langsung di Safari atau Chrome iPad.
2. **Naskah Lengkap Markdown (`BUKU-LENGKAP.md`):**  
   Satu berkas lengkap yang nyaman dibaca langsung di browser GitHub.
3. **Dokumen Word (`buku-strategi-80-20.docx`):**  
   Diformat dengan tipografi 14pt yang lapang, siap diimpor ke aplikasi pengolah kata.
4. **Dokumen PDF (`buku-strategi-80-20.pdf`):**  
   Dihasilkan khusus dengan ukuran huruf besar untuk dibaca langsung di Apple Books / Files pada iPad tanpa perlu zoom.

---

## 🚀 Lisensi & Kontribusi
Dibuat untuk para kreator, agensi, dan profesional digital yang ingin membangun pengungkit media tanpa mengorbankan waktu dan kewarasan.
"""
    with open("README.md", "w", encoding="utf-8") as f:
        f.write(readme_content)
    print("  [OK] Ditulis: README.md")

def build_ipad_html():
    print("Membangun iPad Web Reader (index.html)...")
    
    # Format konten ke HTML
    chapters_html = []
    toc_links = []
    
    for ch in chapters_data:
        toc_links.append(f'<li><a href="#{ch["id"]}" onclick="closeNav()">{ch["title"]}</a></li>')
        
        # Konversi markdown sederhana ke HTML terstruktur
        body_lines = ch["content"].split("\n")
        parsed_body = []
        in_table = False
        table_rows = []
        in_code = False
        code_lines = []
        
        for line in body_lines:
            line_str = line.strip()
            
            # Code block
            if line_str.startswith("```"):
                if in_code:
                    in_code = False
                    parsed_body.append("<pre class='code-block'><code>" + "\n".join(code_lines) + "</code></pre>")
                    code_lines = []
                else:
                    in_code = True
                    code_lines = []
                continue
            if in_code:
                code_lines.append(line)
                continue
                
            # Table handling
            if line_str.startswith("|") and line_str.endswith("|"):
                if not in_table:
                    in_table = True
                    table_rows = []
                # Skip divider line like |:---|:---|
                if "---" in line_str:
                    continue
                cells = [c.strip() for c in line_str.strip("|").split("|")]
                table_rows.append(cells)
                continue
            else:
                if in_table:
                    in_table = False
                    # render table
                    t_html = "<div class='table-responsive'><table class='reader-table'>"
                    if len(table_rows) > 0:
                        t_html += "<thead><tr>" + "".join(f"<th>{c}</th>" for c in table_rows[0]) + "</tr></thead>"
                        if len(table_rows) > 1:
                            t_html += "<tbody>"
                            for r in table_rows[1:]:
                                t_html += "<tr>" + "".join(f"<td>{c}</td>" for c in r) + "</tr>"
                            t_html += "</tbody>"
                    t_html += "</table></div>"
                    parsed_body.append(t_html)
                    table_rows = []
                    
            if not line_str:
                continue
            elif line_str.startswith("## "):
                parsed_body.append(f"<h2 class='ch-h2'>{line_str[3:]}</h2>")
            elif line_str.startswith("### "):
                parsed_body.append(f"<h3 class='ch-h3'>{line_str[4:]}</h3>")
            elif line_str.startswith("> "):
                parsed_body.append(f"<blockquote class='callout'>{line_str[2:]}</blockquote>")
            elif line_str.startswith("- "):
                parsed_body.append(f"<li class='list-item'>{line_str[2:]}</li>")
            elif line_str.startswith("---"):
                parsed_body.append("<hr class='divider'/>")
            elif line_str[0].isdigit() and line_str[1:3] in [". ", ") "]:
                parsed_body.append(f"<li class='ordered-item'>{line_str[3:]}</li>")
            else:
                parsed_body.append(f"<p>{line_str}</p>")
                
        # Jika loop berakhir saat masih dalam table
        if in_table:
            t_html = "<div class='table-responsive'><table class='reader-table'>"
            if len(table_rows) > 0:
                t_html += "<thead><tr>" + "".join(f"<th>{c}</th>" for c in table_rows[0]) + "</tr></thead>"
                if len(table_rows) > 1:
                    t_html += "<tbody>"
                    for r in table_rows[1:]:
                        t_html += "<tr>" + "".join(f"<td>{c}</td>" for c in r) + "</tr>"
                    t_html += "</tbody>"
            t_html += "</table></div>"
            parsed_body.append(t_html)
            
        ch_section = f"""
        <section id="{ch['id']}" class="chapter-card">
            <header class="chapter-header">
                <span class="chapter-tag">BAGIAN BUKU</span>
                <h1 class="chapter-title">{ch['title']}</h1>
                <p class="chapter-subtitle">{ch['subtitle']}</p>
            </header>
            <div class="chapter-body">
                {"".join(parsed_body)}
            </div>
        </section>
        """
        chapters_html.append(ch_section)
        
    html_template = f"""<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=5.0, user-scalable=yes">
    <title>Mesin Konten Asimetris: Strategi 80:20 Media Sosial</title>
    <style>
        :root {{
            --bg-color: #FAF8F5;
            --surface-color: #FFFFFF;
            --text-color: #1F2328;
            --muted-color: #57606A;
            --border-color: #E5E0D8;
            --accent-color: #D9531E;
            --accent-bg: #FFF5EE;
            --callout-border: #D9531E;
            --font-size-base: 20px;
            --line-height-base: 1.8;
            --font-family-reader: 'Charter', 'Georgia', 'Cambria', serif;
            --font-family-ui: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
            --content-max-width: 780px;
        }}

        [data-theme="sepia"] {{
            --bg-color: #F4ECD8;
            --surface-color: #FCF7ED;
            --text-color: #382E24;
            --muted-color: #6D6054;
            --border-color: #E2D7C3;
            --accent-color: #8C4726;
            --accent-bg: #EFE4CC;
            --callout-border: #8C4726;
        }}

        [data-theme="dark"] {{
            --bg-color: #161B22;
            --surface-color: #0D1117;
            --text-color: #E6EDF3;
            --muted-color: #8B949E;
            --border-color: #30363D;
            --accent-color: #F78166;
            --accent-bg: #21262D;
            --callout-border: #F78166;
        }}

        [data-theme="oled"] {{
            --bg-color: #000000;
            --surface-color: #080808;
            --text-color: #D4D4D4;
            --muted-color: #737373;
            --border-color: #262626;
            --accent-color: #FF7B72;
            --accent-bg: #121212;
            --callout-border: #FF7B72;
        }}

        [data-font="sans"] {{
            --font-family-reader: -apple-system, BlinkMacSystemFont, "SF Pro Text", "Segoe UI", Roboto, sans-serif;
        }}

        * {{
            box-sizing: border-box;
            margin: 0;
            padding: 0;
            -webkit-tap-highlight-color: transparent;
        }}

        body {{
            background-color: var(--bg-color);
            color: var(--text-color);
            font-family: var(--font-family-reader);
            font-size: var(--font-size-base);
            line-height: var(--line-height-base);
            transition: background 0.3s ease, color 0.3s ease;
            text-rendering: optimizeLegibility;
            -webkit-font-smoothing: antialiased;
        }}

        /* Header & Toolbar iPad */
        .reader-header {{
            position: sticky;
            top: 0;
            z-index: 100;
            background: var(--surface-color);
            border-bottom: 1px solid var(--border-color);
            padding: 12px 20px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            box-shadow: 0 2px 8px rgba(0,0,0,0.03);
            font-family: var(--font-family-ui);
        }}

        .progress-bar-container {{
            position: absolute;
            bottom: 0;
            left: 0;
            width: 100%;
            height: 3px;
            background: transparent;
        }}

        .progress-bar {{
            height: 100%;
            width: 0%;
            background: var(--accent-color);
            transition: width 0.1s ease;
        }}

        .header-title-wrapper {{
            display: flex;
            align-items: center;
            gap: 12px;
        }}

        .btn-icon {{
            background: none;
            border: 1px solid var(--border-color);
            color: var(--text-color);
            padding: 8px 14px;
            border-radius: 8px;
            font-size: 15px;
            font-weight: 600;
            cursor: pointer;
            display: inline-flex;
            align-items: center;
            gap: 6px;
            transition: all 0.2s;
        }}

        .btn-icon:hover, .btn-icon:active {{
            background: var(--accent-bg);
            border-color: var(--accent-color);
        }}

        .header-book-title {{
            font-size: 16px;
            font-weight: 700;
            color: var(--text-color);
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
            max-width: 320px;
        }}

        .controls-group {{
            display: flex;
            align-items: center;
            gap: 8px;
        }}

        /* Sidebar Drawer */
        .drawer-overlay {{
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0,0,0,0.4);
            z-index: 200;
            display: none;
            backdrop-filter: blur(2px);
        }}

        .toc-drawer {{
            position: fixed;
            top: 0;
            left: -380px;
            width: 340px;
            height: 100%;
            background: var(--surface-color);
            z-index: 210;
            transition: left 0.3s cubic-bezier(0.1, 0.9, 0.2, 1);
            padding: 24px 20px;
            overflow-y: auto;
            border-right: 1px solid var(--border-color);
            font-family: var(--font-family-ui);
        }}

        .toc-drawer.open {{
            left: 0;
        }}

        .drawer-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 20px;
            padding-bottom: 12px;
            border-bottom: 1px solid var(--border-color);
        }}

        .drawer-title {{
            font-size: 18px;
            font-weight: 800;
        }}

        .toc-list {{
            list-style: none;
        }}

        .toc-list li {{
            margin-bottom: 14px;
        }}

        .toc-list a {{
            color: var(--text-color);
            text-decoration: none;
            font-size: 15px;
            line-height: 1.4;
            display: block;
            padding: 8px 10px;
            border-radius: 6px;
            transition: background 0.15s;
        }}

        .toc-list a:hover {{
            background: var(--accent-bg);
            color: var(--accent-color);
        }}

        /* Main Container iPad */
        .reader-container {{
            max-width: var(--content-max-width);
            margin: 0 auto;
            padding: 40px 24px 100px 24px;
        }}

        /* Book Title Hero */
        .book-hero {{
            text-align: center;
            padding: 40px 10px 50px 10px;
            margin-bottom: 40px;
            border-bottom: 2px solid var(--border-color);
        }}

        .book-badge {{
            display: inline-block;
            font-family: var(--font-family-ui);
            background: var(--accent-bg);
            color: var(--accent-color);
            font-weight: 800;
            font-size: 13px;
            letter-spacing: 1.5px;
            padding: 6px 14px;
            border-radius: 20px;
            margin-bottom: 16px;
            text-transform: uppercase;
        }}

        .book-main-title {{
            font-size: 2.2em;
            line-height: 1.25;
            font-weight: 900;
            margin-bottom: 14px;
            color: var(--text-color);
        }}

        .book-sub-title {{
            font-size: 1.15em;
            color: var(--muted-color);
            max-width: 620px;
            margin: 0 auto 20px auto;
            font-style: italic;
        }}

        .ipad-tip-box {{
            background: var(--accent-bg);
            border: 1px dashed var(--accent-color);
            border-radius: 12px;
            padding: 14px 18px;
            font-family: var(--font-family-ui);
            font-size: 14px;
            color: var(--text-color);
            max-width: 580px;
            margin: 20px auto 0 auto;
            text-align: left;
            display: flex;
            align-items: center;
            gap: 12px;
        }}

        /* Chapter Cards */
        .chapter-card {{
            background: var(--surface-color);
            border: 1px solid var(--border-color);
            border-radius: 16px;
            padding: 44px 36px;
            margin-bottom: 50px;
            box-shadow: 0 4px 16px rgba(0,0,0,0.02);
        }}

        .chapter-header {{
            margin-bottom: 30px;
            padding-bottom: 20px;
            border-bottom: 1px solid var(--border-color);
        }}

        .chapter-tag {{
            font-family: var(--font-family-ui);
            font-size: 12px;
            font-weight: 800;
            letter-spacing: 1px;
            color: var(--accent-color);
            text-transform: uppercase;
        }}

        .chapter-title {{
            font-size: 1.7em;
            line-height: 1.3;
            margin: 8px 0;
            font-weight: 800;
        }}

        .chapter-subtitle {{
            font-size: 1.05em;
            color: var(--muted-color);
            font-style: italic;
        }}

        .chapter-body p {{
            margin-bottom: 22px;
            text-align: justify;
        }}

        .ch-h2 {{
            font-size: 1.35em;
            margin: 36px 0 16px 0;
            font-weight: 700;
            color: var(--text-color);
            border-left: 4px solid var(--accent-color);
            padding-left: 12px;
        }}

        .ch-h3 {{
            font-size: 1.15em;
            margin: 24px 0 12px 0;
            font-weight: 700;
        }}

        .callout {{
            border-left: 4px solid var(--callout-border);
            background: var(--accent-bg);
            padding: 16px 20px;
            margin: 26px 0;
            border-radius: 0 10px 10px 0;
            font-style: italic;
            color: var(--text-color);
        }}

        .list-item, .ordered-item {{
            margin-left: 28px;
            margin-bottom: 12px;
        }}

        .divider {{
            border: none;
            height: 1px;
            background: var(--border-color);
            margin: 34px 0;
        }}

        .code-block {{
            background: var(--bg-color);
            border: 1px solid var(--border-color);
            border-radius: 8px;
            padding: 16px;
            font-family: "SF Mono", Consolas, Monaco, monospace;
            font-size: 14px;
            line-height: 1.5;
            overflow-x: auto;
            margin: 24px 0;
        }}

        /* Table Responsive */
        .table-responsive {{
            overflow-x: auto;
            margin: 26px 0;
            border-radius: 8px;
            border: 1px solid var(--border-color);
        }}

        .reader-table {{
            width: 100%;
            border-collapse: collapse;
            font-family: var(--font-family-ui);
            font-size: 15px;
            line-height: 1.5;
        }}

        .reader-table th, .reader-table td {{
            padding: 12px 16px;
            text-align: left;
            border-bottom: 1px solid var(--border-color);
        }}

        .reader-table th {{
            background: var(--accent-bg);
            color: var(--accent-color);
            font-weight: 700;
        }}

        .reader-table tr:last-child td {{
            border-bottom: none;
        }}

        /* Quick Floating Actions on iPad */
        .ipad-nav-footer {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-top: 40px;
            padding-top: 20px;
            border-top: 1px solid var(--border-color);
            font-family: var(--font-family-ui);
        }}

        @media (max-width: 768px) {{
            .chapter-card {{
                padding: 28px 20px;
            }}
            .header-book-title {{
                display: none;
            }}
            .book-main-title {{
                font-size: 1.8em;
            }}
        }}
    </style>
</head>
<body data-theme="warm" data-font="serif">

    <!-- Header & Progress -->
    <header class="reader-header">
        <div class="progress-bar-container">
            <div id="progressBar" class="progress-bar"></div>
        </div>

        <div class="header-title-wrapper">
            <button class="btn-icon" onclick="openNav()" title="Daftar Isi">
                <span>☰</span> <span>Daftar Isi</span>
            </button>
            <span class="header-book-title">Aturan 80:20 Media Sosial</span>
        </div>

        <div class="controls-group">
            <button class="btn-icon" onclick="toggleTheme()" id="themeBtn" title="Ganti Tema">☀️ Mode</button>
            <button class="btn-icon" onclick="toggleFontFamily()" id="fontBtn" title="Ganti Gaya Font">Aa Font</button>
            <button class="btn-icon" onclick="adjustFontSize(-2)" title="Perkecil Font">A-</button>
            <button class="btn-icon" onclick="adjustFontSize(2)" title="Perbesar Font">A+</button>
        </div>
    </header>

    <!-- Drawer Daftar Isi -->
    <div id="drawerOverlay" class="drawer-overlay" onclick="closeNav()"></div>
    <aside id="tocDrawer" class="toc-drawer">
        <div class="drawer-header">
            <div class="drawer-title">Daftar Bab Buku</div>
            <button class="btn-icon" onclick="closeNav()">✕</button>
        </div>
        <ul class="toc-list">
            {"".join(toc_links)}
        </ul>
    </aside>

    <!-- Main Container -->
    <main class="reader-container">
        
        <!-- Hero Section -->
        <section class="book-hero">
            <span class="book-badge">Edisi Pembaca iPad & Tablet</span>
            <h1 class="book-main-title">ATURAN 80:20 MEDIA SOSIAL</h1>
            <p class="book-sub-title">Mesin Konten Asimetris: 80% Eksploitasi Format Pemenang, 20% Eksperimen & Penjualan Jasa.</p>
            
            <div class="ipad-tip-box">
                <span style="font-size: 24px;">💡</span>
                <div>
                    <strong>Tips Membaca di iPad:</strong> Gunakan tombol <strong>A+ / A-</strong> di atas jika ingin tulisan lebih besar. Klik <strong>Mode</strong> untuk berganti ke tampilan Kertas, Sepia, atau Gelap.
                </div>
            </div>
        </section>

        <!-- Chapter Contents -->
        {"".join(chapters_html)}

        <!-- Footer -->
        <footer style="text-align: center; margin-top: 60px; color: var(--muted-color); font-family: var(--font-family-ui); font-size: 14px;">
            <p><strong>Buku Strategi 80:20 Media Sosial</strong> &copy; 2026. Hak Cipta Dilindungi.</p>
            <p>Dibuat untuk eksekusi nyata, bukan sekadar teori.</p>
        </footer>

    </main>

    <script>
        // Reading Progress Indicator
        window.addEventListener('scroll', () => {{
            const winScroll = document.body.scrollTop || document.documentElement.scrollTop;
            const height = document.documentElement.scrollHeight - document.documentElement.clientHeight;
            const scrolled = (winScroll / height) * 100;
            document.getElementById("progressBar").style.width = scrolled + "%";
        }});

        // TOC Drawer Controls
        function openNav() {{
            document.getElementById("tocDrawer").classList.add("open");
            document.getElementById("drawerOverlay").style.display = "block";
        }}
        function closeNav() {{
            document.getElementById("tocDrawer").classList.remove("open");
            document.getElementById("drawerOverlay").style.display = "none";
        }}

        // Font Size Adjuster (Comfortable for iPad)
        let currentFontSize = 20;
        function adjustFontSize(delta) {{
            currentFontSize = Math.min(28, Math.max(16, currentFontSize + delta));
            document.documentElement.style.setProperty('--font-size-base', currentFontSize + 'px');
            localStorage.setItem('ipad_reader_font_size', currentFontSize);
        }}

        // Theme Switcher (Warm, Sepia, Dark, OLED)
        const themes = ['warm', 'sepia', 'dark', 'oled'];
        let currentThemeIndex = 0;
        function toggleTheme() {{
            currentThemeIndex = (currentThemeIndex + 1) % themes.length;
            const theme = themes[currentThemeIndex];
            document.body.setAttribute('data-theme', theme);
            const btn = document.getElementById('themeBtn');
            if (theme === 'warm') btn.innerHTML = '☀️ Terang';
            else if (theme === 'sepia') btn.innerHTML = '📜 Sepia';
            else if (theme === 'dark') btn.innerHTML = '🌙 Gelap';
            else if (theme === 'oled') btn.innerHTML = '🌑 OLED';
            localStorage.setItem('ipad_reader_theme', theme);
        }}

        // Font Family Toggle (Serif vs Sans)
        let isSerif = true;
        function toggleFontFamily() {{
            isSerif = !isSerif;
            document.body.setAttribute('data-font', isSerif ? 'serif' : 'sans');
            document.getElementById('fontBtn').innerText = isSerif ? 'Aa Serif' : 'Aa Sans';
            localStorage.setItem('ipad_reader_font_type', isSerif ? 'serif' : 'sans');
        }}

        // Restore settings
        window.addEventListener('DOMContentLoaded', () => {{
            const savedSize = localStorage.getItem('ipad_reader_font_size');
            if (savedSize) {{
                currentFontSize = parseInt(savedSize);
                document.documentElement.style.setProperty('--font-size-base', currentFontSize + 'px');
            }}
            const savedTheme = localStorage.getItem('ipad_reader_theme');
            if (savedTheme) {{
                document.body.setAttribute('data-theme', savedTheme);
                currentThemeIndex = themes.indexOf(savedTheme);
                if (currentThemeIndex === -1) currentThemeIndex = 0;
            }}
            const savedFontType = localStorage.getItem('ipad_reader_font_type');
            if (savedFontType === 'sans') {{
                isSerif = false;
                document.body.setAttribute('data-font', 'sans');
                document.getElementById('fontBtn').innerText = 'Aa Sans';
            }}
        }});
    </script>
</body>
</html>
"""
    with open("index.html", "w", encoding="utf-8") as f:
        f.write(html_template)
    print("  [OK] Ditulis: index.html")

def build_docx():
    print("Membangun Dokumen Word (buku-strategi-80-20.docx)...")
    doc = Document()
    
    # Atur margin yang lapang untuk tablet / cetak
    sections = doc.sections
    for s in sections:
        s.top_margin = Inches(1.0)
        s.bottom_margin = Inches(1.0)
        s.left_margin = Inches(1.0)
        s.right_margin = Inches(1.0)
        
    # Styling judul buku
    title_p = doc.add_paragraph()
    title_run = title_p.add_run("ATURAN 80:20 MEDIA SOSIAL")
    title_run.font.size = Pt(26)
    title_run.font.bold = True
    title_run.font.name = "Georgia"
    title_run.font.color.rgb = RGBColor(217, 83, 30)
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    subtitle_p = doc.add_paragraph()
    sub_run = subtitle_p.add_run("Mesin Konten Asimetris: 80% Eksploitasi Format Pemenang, 20% Eksperimen & Penjualan Jasa\nEdisi Tablet & iPad (Font Besar)")
    sub_run.font.size = Pt(14)
    sub_run.font.italic = True
    sub_run.font.name = "Georgia"
    subtitle_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph("\n" + "="*50 + "\n")
    
    # Isi Bab
    for ch in chapters_data:
        # Heading Bab
        h1 = doc.add_heading(level=1)
        h1_run = h1.add_run(ch["title"])
        h1_run.font.size = Pt(18)
        h1_run.font.bold = True
        h1_run.font.name = "Georgia"
        h1_run.font.color.rgb = RGBColor(31, 35, 40)
        
        # Subheading
        sub_ch = doc.add_paragraph()
        sub_ch_run = sub_ch.add_run(ch["subtitle"])
        sub_ch_run.font.size = Pt(12)
        sub_ch_run.font.italic = True
        sub_ch_run.font.color.rgb = RGBColor(100, 100, 100)
        
        # Body content
        lines = ch["content"].split("\n")
        for line in lines:
            line_str = line.strip()
            if not line_str or line_str.startswith("```"):
                continue
            elif line_str.startswith("## "):
                h2 = doc.add_heading(level=2)
                h2_run = h2.add_run(line_str[3:])
                h2_run.font.size = Pt(15)
                h2_run.font.bold = True
                h2_run.font.name = "Georgia"
                h2_run.font.color.rgb = RGBColor(217, 83, 30)
            elif line_str.startswith("### "):
                h3 = doc.add_heading(level=3)
                h3_run = h3.add_run(line_str[4:])
                h3_run.font.size = Pt(13)
                h3_run.font.bold = True
                h3_run.font.name = "Georgia"
            elif line_str.startswith("> "):
                callout_p = doc.add_paragraph()
                c_run = callout_p.add_run(line_str[2:])
                c_run.font.size = Pt(13)
                c_run.font.italic = True
                c_run.font.color.rgb = RGBColor(140, 71, 38)
            elif line_str.startswith("- ") or (line_str[0].isdigit() and line_str[1:3] in [". ", ") "]):
                list_p = doc.add_paragraph(style='List Bullet')
                item_text = line_str[2:] if line_str.startswith("- ") else line_str[3:]
                l_run = list_p.add_run(item_text)
                l_run.font.size = Pt(13)
                l_run.font.name = "Georgia"
            elif line_str.startswith("|"):
                continue # Hindari format tabel mentah di paragraf biasa
            elif line_str.startswith("---"):
                continue
            else:
                p = doc.add_paragraph()
                p_run = p.add_run(line_str)
                p_run.font.size = Pt(13)
                p_run.font.name = "Georgia"
                p.paragraph_format.line_spacing = 1.3
                p.paragraph_format.space_after = Pt(8)
                
        doc.add_page_break()
        
    doc.save("buku-strategi-80-20.docx")
    print("  [OK] Ditulis: buku-strategi-80-20.docx")

def build_pdf():
    print("Membangun Dokumen PDF (buku-strategi-80-20.pdf)...")
    pdf_filename = "buku-strategi-80-20.pdf"
    
    # Setup document dengan margin nyaman untuk iPad
    doc = SimpleDocTemplate(
        pdf_filename,
        pagesize=letter,
        leftMargin=54,
        rightMargin=54,
        topMargin=54,
        bottomMargin=54
    )
    
    styles = getSampleStyleSheet()
    
    # Buat style khusus font besar ramah iPad
    title_style = ParagraphStyle(
        'BookTitle',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=24,
        leading=30,
        textColor=colors.HexColor("#D9531E"),
        alignment=1, # Center
        spaceAfter=12
    )
    
    subtitle_style = ParagraphStyle(
        'BookSubtitle',
        parent=styles['Normal'],
        fontName='Helvetica-Oblique',
        fontSize=13,
        leading=18,
        textColor=colors.HexColor("#555555"),
        alignment=1, # Center
        spaceAfter=24
    )
    
    ch_title_style = ParagraphStyle(
        'ChapterTitle',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=17,
        leading=22,
        textColor=colors.HexColor("#1F2328"),
        spaceBefore=18,
        spaceAfter=6
    )
    
    ch_subtitle_style = ParagraphStyle(
        'ChapterSubtitle',
        parent=styles['Normal'],
        fontName='Helvetica-Oblique',
        fontSize=11,
        leading=15,
        textColor=colors.HexColor("#777777"),
        spaceAfter=14
    )
    
    h2_style = ParagraphStyle(
        'H2Style',
        parent=styles['Heading2'],
        fontName='Helvetica-Bold',
        fontSize=14,
        leading=18,
        textColor=colors.HexColor("#D9531E"),
        spaceBefore=14,
        spaceAfter=8
    )
    
    body_style = ParagraphStyle(
        'BodyLarge',
        parent=styles['Normal'],
        fontName='Helvetica',
        fontSize=12,
        leading=17,
        textColor=colors.HexColor("#222222"),
        spaceAfter=10
    )
    
    callout_style = ParagraphStyle(
        'CalloutStyle',
        parent=styles['Normal'],
        fontName='Helvetica-Oblique',
        fontSize=12,
        leading=16,
        textColor=colors.HexColor("#8C4726"),
        leftIndent=15,
        rightIndent=15,
        spaceBefore=10,
        spaceAfter=12
    )
    
    story = []
    
    # Cover / Judul
    story.append(Spacer(1, 40))
    story.append(Paragraph("ATURAN 80:20 MEDIA SOSIAL", title_style))
    story.append(Paragraph("Mesin Konten Asimetris: 80% Eksploitasi Format Pemenang, 20% Eksperimen & Jualan Jasa<br/><b>Edisi Tablet & iPad (Font Besar)</b>", subtitle_style))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor("#E5E0D8"), spaceAfter=30))
    
    for ch in chapters_data:
        story.append(Paragraph(ch["title"], ch_title_style))
        story.append(Paragraph(ch["subtitle"], ch_subtitle_style))
        story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#E5E0D8"), spaceAfter=14))
        
        lines = ch["content"].split("\n")
        for line in lines:
            line_str = line.strip()
            if not line_str or line_str.startswith("```") or line_str.startswith("|") or line_str.startswith("---"):
                continue
            elif line_str.startswith("## "):
                story.append(Paragraph(line_str[3:], h2_style))
            elif line_str.startswith("### "):
                story.append(Paragraph(f"<b>{line_str[4:]}</b>", body_style))
            elif line_str.startswith("> "):
                clean_c = line_str[2:].replace("<", "&lt;").replace(">", "&gt;")
                story.append(Paragraph(f"<i>&ldquo;{clean_c}&rdquo;</i>", callout_style))
            elif line_str.startswith("- "):
                clean_l = line_str[2:].replace("<", "&lt;").replace(">", "&gt;")
                story.append(Paragraph(f"&bull; {clean_l}", body_style))
            else:
                clean_p = line_str.replace("<", "&lt;").replace(">", "&gt;")
                story.append(Paragraph(clean_p, body_style))
                
        story.append(Spacer(1, 25))
        
    doc.build(story)
    print("  [OK] Ditulis: buku-strategi-80-20.pdf")

if __name__ == "__main__":
    print("Mulai pembuatan paket buku...")
    build_markdown_files()
    build_readme()
    build_ipad_html()
    build_docx()
    build_pdf()
    print("Semua berkas buku berhasil dibangun dengan sempurna!")
