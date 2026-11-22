# -*- coding: utf-8 -*-
"""
Builder v2.1: Membangun buku Word (.docx) lengkap dengan parsing Markdown penuh (Bold & Italic asli),
foto & ilustrasi di setiap bab, serta atribusi penulis 'By Zen'.
Sinkronisasi ke Web Reader iPad (index.html), PDF, dan Markdown.
"""

import os
import re
import shutil
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# Data bab buku lengkap dengan metadata ilustrasi
chapters_data = [
    {
        "id": "bab-1",
        "file": "01-paradoks-kreator.md",
        "title": "Bab 1: Paradoks Kreator — Mengapa Konsistensi Saja Tidak Cukup",
        "subtitle": "Menghentikan Kebiasaan Membunuh Format Konten yang Baru Saja Berhasil",
        "image": "assets/ilustrasi_bab1_paradoks.jpg",
        "caption": "Ilustrasi 1.1: Paradoks Kreator — Kurva Pertumbuhan Eksponensial Terputus oleh Manuver Topik Acak",
        "content": """## Realitas Pahit di Balik Tombol 'Posting'

Sebagian besar kreator dan pemilik bisnis digital beroperasi di bawah ilusi yang berbahaya: *"Jika saya memposting konten setiap hari secara konsisten, akun saya pada akhirnya pasti akan meledak."*

Kenyataannya tidak demikian.

Konsisten memposting konten yang tidak diinginkan oleh pasar hanya akan menghasilkan satu hal: Anda menjadi orang yang gagal dengan sangat disiplin.

Pasar media sosial tidak memberikan kompensasi kepada keringat, durasi rekaman, atau seberapa keras Anda berjuang. Pasar hanya memedulikan satu hal sederhana: **apakah konten ini memberikan nilai spesifik yang membuat orang bersedia menukarkan waktu dan perhatian mereka?**

---

## Penyakit 'Creative ADD': Membunuh Format Pemenang

Amati pola yang berulang pada 95% kreator pemula:

1. Mereka mencoba 10 topik yang berbeda secara acak: tips produktivitas, kutipan motivasi, cerita keseharian, tutorial teknis, dan bedah website.
2. Tiba-tiba, video nomor 5 tentang **"Bedah Website"** meledak dan mendapatkan 50.000 penayangan serta ratusan komentar antusias.
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
        "image": "assets/diagram_bab2_arsitektur.png",
        "caption": "Diagram 2.1: Arsitektur Konten 80:20 — Pembagian Beban Antara Skala Eksploitasi & Eksplorasi Asimetris",
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
        "image": "assets/ilustrasi_bab3_website.jpg",
        "caption": "Ilustrasi 3.1: Cetak Biru Bedah Website — 80 Variasi Taktis dari Satu Format Pemenang",
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
        "image": "assets/ilustrasi_bab4_sandbox.jpg",
        "caption": "Ilustrasi 4.1: Ruang Kerja Studio 20% — Mengonversi Audiens Menjadi Klien Lewat Bukti Nyata",
        "content": """## Mengapa 80% Saja Bisa Membunuh Akun Anda?

Jika Anda hanya memposting 100% video bedah website selamanya tanpa pernah melakukan hal lain, dua masalah besar akan muncul:
1. **Audiens melihat Anda sebagai ensiklopedia gratis, bukan sebagai penyedia jasa profesional yang bisa mereka sewa.**
2. **Anda tidak pernah tahu apakah ada format lain yang berpotensi menghasilkan dampak 10 kali lebih besar bagi bisnis Anda.**

Porsi 20% adalah ruang perlindungan dan pertumbuhan strategis akun Anda. Dalam siklus 100 konten, 20 konten ini didistribusikan ke dalam 3 fungsi krusial:
- **8 Konten Penawaran Jasa & Portofolio Klien** (Monetisasi Langsung)
- **6 Konten Personal Brand, Vlog & Filosofi Kerja** (Kedekatan Emosional)
- **6 Konten Eksperimen Format Liar** (R&D Format Masa Depan)

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
        "image": "assets/ilustrasi_bab5_flywheel.jpg",
        "caption": "Ilustrasi 5.1: Flywheel Regenerasi Konten — Siklus Promosi Eksperimen Menjadi Format Juara Baru",
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
        "image": "assets/diagram_bab6_kalender.png",
        "caption": "Diagram 6.1: Matriks Distribusi Mingguan 4:1 & Sistem Batching Tanpa Burnout",
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
        "image": "assets/ilustrasi_bab7_aksioma.jpg",
        "caption": "Ilustrasi 7.1: Monolit Prinsip Digital Abadi — Bertahan Menghadapi Pasang Surut Algoritma",
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

def parse_markdown_to_docx_runs(paragraph, text, base_font_size=13.5, font_name="Georgia", default_color=None, inherit_italic=False, inherit_bold=False):
    """
    Mengonversi string yang mengandung format Markdown inline (**bold**, *italic*, ***bolditalic***, `code`)
    menjadi runs Word asli pada objek paragraph python-docx.
    Menghilangkan tanda asterisk (*) dan backtick (`) sehingga menghasilkan teks Word yang bersih dan terformat.
    """
    pattern = re.compile(r'(\*\*\*.*?\*\*\*|\*\*.*?\*\*|\*[^\*]+?\*|`[^`]+?`)')
    tokens = pattern.split(text)
    
    for token in tokens:
        if not token:
            continue
        
        # Bold & Italic: ***text***
        if token.startswith('***') and token.endswith('***'):
            inner = token[3:-3]
            r = paragraph.add_run(inner)
            r.font.name = font_name
            r.font.size = Pt(base_font_size)
            r.bold = True
            r.italic = True
            if default_color:
                r.font.color.rgb = default_color
                
        # Bold: **text** (bisa ada *italic* di dalamnya)
        elif token.startswith('**') and token.endswith('**'):
            inner = token[2:-2]
            sub_pattern = re.compile(r'(\*[^\*]+?\*)')
            sub_tokens = sub_pattern.split(inner)
            for st in sub_tokens:
                if not st:
                    continue
                if st.startswith('*') and st.endswith('*'):
                    r = paragraph.add_run(st[1:-1])
                    r.font.name = font_name
                    r.font.size = Pt(base_font_size)
                    r.bold = True
                    r.italic = True
                else:
                    r = paragraph.add_run(st)
                    r.font.name = font_name
                    r.font.size = Pt(base_font_size)
                    r.bold = True
                    r.italic = inherit_italic
                if default_color:
                    r.font.color.rgb = default_color
                    
        # Italic: *text* (bisa ada **bold** di dalamnya)
        elif token.startswith('*') and token.endswith('*'):
            inner = token[1:-1]
            sub_pattern = re.compile(r'(\*\*.*?\*\*)')
            sub_tokens = sub_pattern.split(inner)
            for st in sub_tokens:
                if not st:
                    continue
                if st.startswith('**') and st.endswith('**'):
                    r = paragraph.add_run(st[2:-2])
                    r.font.name = font_name
                    r.font.size = Pt(base_font_size)
                    r.bold = True
                    r.italic = True
                else:
                    r = paragraph.add_run(st)
                    r.font.name = font_name
                    r.font.size = Pt(base_font_size)
                    r.bold = inherit_bold
                    r.italic = True
                if default_color:
                    r.font.color.rgb = default_color
                    
        # Code: `code`
        elif token.startswith('`') and token.endswith('`'):
            inner = token[1:-1]
            r = paragraph.add_run(inner)
            r.font.name = "Consolas"
            r.font.size = Pt(base_font_size - 1)
            r.font.color.rgb = RGBColor(180, 50, 50) if not default_color else default_color
            
        # Regular text
        else:
            r = paragraph.add_run(token)
            r.font.name = font_name
            r.font.size = Pt(base_font_size)
            r.bold = inherit_bold
            r.italic = inherit_italic
            if default_color:
                r.font.color.rgb = default_color

def set_cell_background(cell, fill_hex):
    """Set warna latar sel tabel Word"""
    tcPr = cell._element.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def build_docx_rich():
    print("Membangun Dokumen Word Kaya Format (WordCraft) -> buku-strategi-80-20.docx...")
    doc = Document()
    
    # Atur margin ramah iPad
    for s in doc.sections:
        s.top_margin = Inches(0.9)
        s.bottom_margin = Inches(0.9)
        s.left_margin = Inches(0.9)
        s.right_margin = Inches(0.9)
        
    # --- HALAMAN COVER ---
    if os.path.exists("assets/cover.jpg"):
        cover_p = doc.add_paragraph()
        cover_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        c_run = cover_p.add_run()
        c_run.add_picture("assets/cover.jpg", width=Inches(5.0))
        
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(14)
    p_title.paragraph_format.space_after = Pt(4)
    run_t = p_title.add_run("ATURAN 80:20 MEDIA SOSIAL")
    run_t.font.name = "Georgia"
    run_t.font.size = Pt(24)
    run_t.font.bold = True
    run_t.font.color.rgb = RGBColor(217, 83, 30)
    
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(10)
    run_s = p_sub.add_run("Mesin Konten Asimetris: 80% Eksploitasi Format Pemenang, 20% Eksperimen & Penjualan Jasa\nEdisi Eksekutif iPad & Tablet (Font Besar & Berilustrasi)")
    run_s.font.name = "Georgia"
    run_s.font.size = Pt(13)
    run_s.font.italic = True
    run_s.font.color.rgb = RGBColor(87, 96, 106)

    # Penulis: By Zen
    p_author = doc.add_paragraph()
    p_author.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_author.paragraph_format.space_after = Pt(20)
    run_auth = p_author.add_run("Karya: Zen")
    run_auth.font.name = "Georgia"
    run_auth.font.size = Pt(13.5)
    run_auth.font.bold = True
    run_auth.font.color.rgb = RGBColor(217, 83, 30)

    doc.add_page_break()

    # --- DAFTAR ISI ---
    toc_h = doc.add_heading(level=1)
    toc_h_run = toc_h.add_run("DAFTAR ISI")
    toc_h_run.font.name = "Georgia"
    toc_h_run.font.size = Pt(18)
    toc_h_run.font.bold = True
    toc_h_run.font.color.rgb = RGBColor(31, 35, 40)
    
    for ch in chapters_data:
        p_toc = doc.add_paragraph(style='List Bullet')
        p_toc.paragraph_format.space_after = Pt(6)
        r_title = p_toc.add_run(ch["title"])
        r_title.font.name = "Georgia"
        r_title.font.size = Pt(12.5)
        r_title.font.bold = True
        r_title.font.color.rgb = RGBColor(217, 83, 30)
        
        r_sub = p_toc.add_run(f"\n   {ch['subtitle']}")
        r_sub.font.name = "Georgia"
        r_sub.font.size = Pt(11)
        r_sub.font.italic = True
        r_sub.font.color.rgb = RGBColor(100, 100, 100)

    doc.add_page_break()

    # --- BAB DEMI BAB ---
    for ch in chapters_data:
        # Header Bab
        h1 = doc.add_heading(level=1)
        h1.paragraph_format.space_before = Pt(10)
        h1.paragraph_format.space_after = Pt(4)
        h1_run = h1.add_run(ch["title"])
        h1_run.font.size = Pt(18)
        h1_run.font.bold = True
        h1_run.font.name = "Georgia"
        h1_run.font.color.rgb = RGBColor(31, 35, 40)
        
        # Subtitle
        sub_p = doc.add_paragraph()
        sub_p.paragraph_format.space_after = Pt(14)
        sub_run = sub_p.add_run(ch["subtitle"])
        sub_run.font.size = Pt(12.5)
        sub_run.font.italic = True
        sub_run.font.name = "Georgia"
        sub_run.font.color.rgb = RGBColor(100, 100, 100)
        
        # Embed Gambar Ilustrasi Bab
        if os.path.exists(ch["image"]):
            img_p = doc.add_paragraph()
            img_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            img_p.paragraph_format.space_before = Pt(8)
            img_p.paragraph_format.space_after = Pt(4)
            r_img = img_p.add_run()
            r_img.add_picture(ch["image"], width=Inches(5.2))
            
            # Caption Gambar
            cap_p = doc.add_paragraph()
            cap_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cap_p.paragraph_format.space_after = Pt(18)
            cap_run = cap_p.add_run(ch["caption"])
            cap_run.font.name = "Georgia"
            cap_run.font.size = Pt(10.5)
            cap_run.font.italic = True
            cap_run.font.color.rgb = RGBColor(110, 110, 110)

        # Konten Teks Bab dengan Parser Markdown Asli
        lines = ch["content"].split("\n")
        in_table = False
        table_rows = []
        
        for line in lines:
            line_str = line.strip()
            
            # Tangani Tabel
            if line_str.startswith("|") and line_str.endswith("|"):
                if not in_table:
                    in_table = True
                    table_rows = []
                if "---" in line_str:
                    continue
                cells = [c.strip() for c in line_str.strip("|").split("|")]
                table_rows.append(cells)
                continue
            else:
                if in_table:
                    in_table = False
                    if len(table_rows) > 0:
                        tbl = doc.add_table(rows=len(table_rows), cols=len(table_rows[0]))
                        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
                        for row_idx, r_data in enumerate(table_rows):
                            row = tbl.rows[row_idx]
                            for col_idx, cell_data in enumerate(r_data):
                                cell = row.cells[col_idx]
                                p = cell.paragraphs[0]
                                p.text = "" # Bersihkan teks default
                                is_head = (row_idx == 0)
                                parse_markdown_to_docx_runs(
                                    p, cell_data,
                                    base_font_size=11,
                                    font_name="Georgia",
                                    default_color=RGBColor(217, 83, 30) if is_head else RGBColor(31, 35, 40),
                                    inherit_bold=is_head
                                )
                                if is_head:
                                    set_cell_background(cell, "FFF5EE")
                                else:
                                    set_cell_background(cell, "FAFAFA" if row_idx % 2 == 1 else "FFFFFF")
                        doc.add_paragraph().paragraph_format.space_after = Pt(8)
                    table_rows = []
                    
            if not line_str or line_str.startswith("```"):
                continue
            elif line_str.startswith("## "):
                h2 = doc.add_heading(level=2)
                h2.paragraph_format.space_before = Pt(16)
                h2.paragraph_format.space_after = Pt(6)
                parse_markdown_to_docx_runs(
                    h2, line_str[3:],
                    base_font_size=14.5,
                    font_name="Georgia",
                    default_color=RGBColor(217, 83, 30),
                    inherit_bold=True
                )
            elif line_str.startswith("### "):
                h3 = doc.add_heading(level=3)
                h3.paragraph_format.space_before = Pt(12)
                h3.paragraph_format.space_after = Pt(4)
                parse_markdown_to_docx_runs(
                    h3, line_str[4:],
                    base_font_size=13.0,
                    font_name="Georgia",
                    default_color=RGBColor(31, 35, 40),
                    inherit_bold=True
                )
            elif line_str.startswith("> "):
                callout_tbl = doc.add_table(rows=1, cols=1)
                callout_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
                cell = callout_tbl.cell(0, 0)
                set_cell_background(cell, "FFF5EE")
                c_p = cell.paragraphs[0]
                c_p.paragraph_format.left_indent = Inches(0.15)
                c_p.paragraph_format.right_indent = Inches(0.15)
                c_p.paragraph_format.space_before = Pt(6)
                c_p.paragraph_format.space_after = Pt(6)
                
                # Hilangkan '> ' dan kutip
                quote_text = line_str[2:].strip()
                parse_markdown_to_docx_runs(
                    c_p, f"“{quote_text}”",
                    base_font_size=12.5,
                    font_name="Georgia",
                    default_color=RGBColor(140, 71, 38),
                    inherit_italic=True
                )
                doc.add_paragraph().paragraph_format.space_after = Pt(6)
            elif line_str.startswith("- ") or (line_str[0].isdigit() and line_str[1:3] in [". ", ") "]):
                list_p = doc.add_paragraph(style='List Bullet')
                item_text = line_str[2:] if line_str.startswith("- ") else line_str[3:]
                parse_markdown_to_docx_runs(
                    list_p, item_text,
                    base_font_size=13.0,
                    font_name="Georgia"
                )
                list_p.paragraph_format.space_after = Pt(4)
            elif line_str.startswith("---"):
                continue
            else:
                p = doc.add_paragraph()
                parse_markdown_to_docx_runs(
                    p, line_str,
                    base_font_size=13.5,
                    font_name="Georgia"
                )
                p.paragraph_format.line_spacing = 1.35
                p.paragraph_format.space_after = Pt(10)
                
        doc.add_page_break()

    doc.save("buku-strategi-80-20.docx")
    print("  [SUKSES] Dokumen Word WordCraft berhasil dibuat: buku-strategi-80-20.docx")

def build_markdown_with_images():
    print("Membangun Berkas Markdown dengan atribusi 'By Zen'...")
    full_book_md = []
    full_book_md.append("# 📚 MESIN KONTEN ASIMETRIS: STRATEGI 80:20 MEDIA SOSIAL")
    full_book_md.append("### *Panduan Taktis Mengembangkan Akun, Membangun Otoritas, dan Memonetisasi Keahlian Tanpa Kehabisan Ide*\n")
    full_book_md.append("**Karya: Zen**\n")
    full_book_md.append("![Cover Buku](assets/cover.jpg)\n")
    full_book_md.append("> **Format Baca Optimal:** iPad & Tablet (Font Besar, Berilustrasi Penuh)\n")
    full_book_md.append("---\n")
    full_book_md.append("## DAFTAR ISI\n")
    for ch in chapters_data:
        full_book_md.append(f"- [{ch['title']}](#{ch['id']})")
    full_book_md.append("\n---\n")

    for ch in chapters_data:
        ch_path = os.path.join("chapters", ch["file"])
        with open(ch_path, "w", encoding="utf-8") as f:
            f.write(f"# {ch['title']}\n")
            f.write(f"*{ch['subtitle']}*\n\n")
            f.write(f"![{ch['caption']}](../{ch['image']})\n")
            f.write(f"*{ch['caption']}*\n\n")
            f.write(ch['content'])
            f.write("\n")
            
        full_book_md.append(f"<a id='{ch['id']}'></a>\n")
        full_book_md.append(f"# {ch['title']}\n")
        full_book_md.append(f"*{ch['subtitle']}*\n\n")
        full_book_md.append(f"![{ch['caption']}]({ch['image']})\n")
        full_book_md.append(f"*{ch['caption']}*\n\n")
        full_book_md.append(ch['content'])
        full_book_md.append("\n\n---\n")

    with open("BUKU-LENGKAP.md", "w", encoding="utf-8") as f:
        f.write("\n".join(full_book_md))
    print("  [SUKSES] BUKU-LENGKAP.md & chapters/ diperbarui!")

def build_ipad_html_rich():
    print("Membangun Web Reader iPad (index.html)...")
    chapters_html = []
    toc_links = []
    
    for ch in chapters_data:
        toc_links.append(f'<li><a href="#{ch["id"]}" onclick="closeNav()">{ch["title"]}</a></li>')
        
        body_lines = ch["content"].split("\n")
        parsed_body = []
        in_table = False
        table_rows = []
        in_code = False
        code_lines = []
        
        img_tag = f"""
        <figure class="chapter-illustration">
            <img src="{ch['image']}" alt="{ch['caption']}" loading="lazy" class="responsive-img"/>
            <figcaption>{ch['caption']}</figcaption>
        </figure>
        """
        parsed_body.append(img_tag)
        
        for line in body_lines:
            line_str = line.strip()
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
                
            if line_str.startswith("|") and line_str.endswith("|"):
                if not in_table:
                    in_table = True
                    table_rows = []
                if "---" in line_str:
                    continue
                cells = [c.strip() for c in line_str.strip("|").split("|")]
                table_rows.append(cells)
                continue
            else:
                if in_table:
                    in_table = False
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
                # Format bold & italic untuk HTML
                html_line = line_str
                html_line = re.sub(r'\*\*\*(.*?)\*\*\*', r'<strong><em>\1</em></strong>', html_line)
                html_line = re.sub(r'\*\*(.*?)\*\*', r'<strong>\1</strong>', html_line)
                html_line = re.sub(r'\*(.*?)\*', r'<em>\1</em>', html_line)
                html_line = re.sub(r'`(.*?)`', r'<code>\1</code>', html_line)
                parsed_body.append(f"<p>{html_line}</p>")
                
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
    <title>Mesin Konten Asimetris: Strategi 80:20 Media Sosial - by Zen</title>
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
            --content-max-width: 800px;
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

        .reader-container {{
            max-width: var(--content-max-width);
            margin: 0 auto;
            padding: 40px 24px 100px 24px;
        }}

        .book-hero {{
            text-align: center;
            padding: 30px 10px 40px 10px;
            margin-bottom: 40px;
            border-bottom: 2px solid var(--border-color);
        }}

        .hero-cover-img {{
            max-width: 340px;
            width: 100%;
            border-radius: 14px;
            box-shadow: 0 10px 30px rgba(0,0,0,0.15);
            margin-bottom: 24px;
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
            margin-bottom: 10px;
            color: var(--text-color);
        }}

        .book-author {{
            font-size: 1.15em;
            font-weight: 700;
            color: var(--accent-color);
            margin-bottom: 14px;
            font-family: var(--font-family-ui);
        }}

        .book-sub-title {{
            font-size: 1.15em;
            color: var(--muted-color);
            max-width: 620px;
            margin: 0 auto 20px auto;
            font-style: italic;
        }}

        .chapter-card {{
            background: var(--surface-color);
            border: 1px solid var(--border-color);
            border-radius: 16px;
            padding: 44px 36px;
            margin-bottom: 50px;
            box-shadow: 0 4px 16px rgba(0,0,0,0.02);
        }}

        .chapter-header {{
            margin-bottom: 24px;
            padding-bottom: 16px;
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

        .chapter-illustration {{
            margin: 24px 0 34px 0;
            text-align: center;
        }}

        .responsive-img {{
            max-width: 100%;
            height: auto;
            border-radius: 12px;
            box-shadow: 0 6px 20px rgba(0,0,0,0.08);
            border: 1px solid var(--border-color);
        }}

        figcaption {{
            font-size: 14px;
            font-style: italic;
            color: var(--muted-color);
            margin-top: 10px;
            font-family: var(--font-family-ui);
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

        @media (max-width: 768px) {{
            .chapter-card {{
                padding: 26px 18px;
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

    <main class="reader-container">
        
        <section class="book-hero">
            <img src="assets/cover.jpg" alt="Cover Buku Aturan 80:20 Media Sosial by Zen" class="hero-cover-img"/>
            <br/>
            <span class="book-badge">Edisi Pembaca iPad & Tablet</span>
            <h1 class="book-main-title">ATURAN 80:20 MEDIA SOSIAL</h1>
            <p class="book-author">Karya: Zen</p>
            <p class="book-sub-title">Mesin Konten Asimetris: 80% Eksploitasi Format Pemenang, 20% Eksperimen & Penjualan Jasa.</p>
        </section>

        {"".join(chapters_html)}

        <footer style="text-align: center; margin-top: 60px; color: var(--muted-color); font-family: var(--font-family-ui); font-size: 14px;">
            <p><strong>Buku Strategi 80:20 Media Sosial</strong> &copy; 2026 Zen. Hak Cipta Dilindungi.</p>
        </footer>

    </main>

    <script>
        window.addEventListener('scroll', () => {{
            const winScroll = document.body.scrollTop || document.documentElement.scrollTop;
            const height = document.documentElement.scrollHeight - document.documentElement.clientHeight;
            const scrolled = (winScroll / height) * 100;
            document.getElementById("progressBar").style.width = scrolled + "%";
        }});

        function openNav() {{
            document.getElementById("tocDrawer").classList.add("open");
            document.getElementById("drawerOverlay").style.display = "block";
        }}
        function closeNav() {{
            document.getElementById("tocDrawer").classList.remove("open");
            document.getElementById("drawerOverlay").style.display = "none";
        }}

        let currentFontSize = 20;
        function adjustFontSize(delta) {{
            currentFontSize = Math.min(28, Math.max(16, currentFontSize + delta));
            document.documentElement.style.setProperty('--font-size-base', currentFontSize + 'px');
            localStorage.setItem('ipad_reader_font_size', currentFontSize);
        }}

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

        let isSerif = true;
        function toggleFontFamily() {{
            isSerif = !isSerif;
            document.body.setAttribute('data-font', isSerif ? 'serif' : 'sans');
            document.getElementById('fontBtn').innerText = isSerif ? 'Aa Serif' : 'Aa Sans';
            localStorage.setItem('ipad_reader_font_type', isSerif ? 'serif' : 'sans');
        }}

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
    print("  [SUKSES] index.html diperbarui dengan atribusi 'by Zen'!")

def build_pdf_rich():
    print("Membangun Dokumen PDF dengan atribusi 'by Zen' (buku-strategi-80-20.pdf)...")
    pdf_filename = "buku-strategi-80-20.pdf"
    
    doc = SimpleDocTemplate(
        pdf_filename,
        pagesize=letter,
        leftMargin=50,
        rightMargin=50,
        topMargin=50,
        bottomMargin=50
    )
    
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(
        'BookTitle',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=24,
        leading=30,
        textColor=colors.HexColor("#D9531E"),
        alignment=1,
        spaceAfter=6
    )

    author_style = ParagraphStyle(
        'BookAuthor',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        fontSize=13,
        leading=17,
        textColor=colors.HexColor("#D9531E"),
        alignment=1,
        spaceAfter=10
    )
    
    subtitle_style = ParagraphStyle(
        'BookSubtitle',
        parent=styles['Normal'],
        fontName='Helvetica-Oblique',
        fontSize=12,
        leading=16,
        textColor=colors.HexColor("#555555"),
        alignment=1,
        spaceAfter=20
    )
    
    ch_title_style = ParagraphStyle(
        'ChapterTitle',
        parent=styles['Heading1'],
        fontName='Helvetica-Bold',
        fontSize=17,
        leading=22,
        textColor=colors.HexColor("#1F2328"),
        spaceBefore=16,
        spaceAfter=4
    )
    
    ch_subtitle_style = ParagraphStyle(
        'ChapterSubtitle',
        parent=styles['Normal'],
        fontName='Helvetica-Oblique',
        fontSize=11,
        leading=15,
        textColor=colors.HexColor("#777777"),
        spaceAfter=12
    )
    
    caption_style = ParagraphStyle(
        'CaptionStyle',
        parent=styles['Normal'],
        fontName='Helvetica-Oblique',
        fontSize=10,
        leading=14,
        textColor=colors.HexColor("#666666"),
        alignment=1,
        spaceBefore=6,
        spaceAfter=16
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
        fontSize=11.5,
        leading=16.5,
        textColor=colors.HexColor("#222222"),
        spaceAfter=10
    )
    
    callout_style = ParagraphStyle(
        'CalloutStyle',
        parent=styles['Normal'],
        fontName='Helvetica-Oblique',
        fontSize=11.5,
        leading=16,
        textColor=colors.HexColor("#8C4726"),
        leftIndent=15,
        rightIndent=15,
        spaceBefore=8,
        spaceAfter=12
    )
    
    story = []
    
    if os.path.exists("assets/cover.jpg"):
        story.append(Spacer(1, 20))
        story.append(RLImage("assets/cover.jpg", width=280, height=280))
        story.append(Spacer(1, 15))
        
    story.append(Paragraph("ATURAN 80:20 MEDIA SOSIAL", title_style))
    story.append(Paragraph("<b>Karya: Zen</b>", author_style))
    story.append(Paragraph("Mesin Konten Asimetris: 80% Eksploitasi Format Pemenang, 20% Eksperimen & Jualan Jasa<br/><b>Edisi Tablet & iPad (Font Besar & Berilustrasi)</b>", subtitle_style))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor("#E5E0D8"), spaceAfter=24))
    
    for ch in chapters_data:
        story.append(Paragraph(ch["title"], ch_title_style))
        story.append(Paragraph(ch["subtitle"], ch_subtitle_style))
        
        if os.path.exists(ch["image"]):
            try:
                story.append(RLImage(ch["image"], width=460, height=260))
                story.append(Paragraph(ch["caption"], caption_style))
            except Exception as e:
                print(f"Gagal menyematkan gambar di PDF: {e}")
                
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
                clean_c = re.sub(r'\*\*(.*?)\*\*', r'<b>\1</b>', clean_c)
                clean_c = re.sub(r'\*(.*?)\*', r'<i>\1</i>', clean_c)
                story.append(Paragraph(f"&ldquo;{clean_c}&rdquo;", callout_style))
            elif line_str.startswith("- "):
                clean_l = line_str[2:].replace("<", "&lt;").replace(">", "&gt;")
                clean_l = re.sub(r'\*\*(.*?)\*\*', r'<b>\1</b>', clean_l)
                clean_l = re.sub(r'\*(.*?)\*', r'<i>\1</i>', clean_l)
                story.append(Paragraph(f"&bull; {clean_l}", body_style))
            else:
                clean_p = line_str.replace("<", "&lt;").replace(">", "&gt;")
                clean_p = re.sub(r'\*\*(.*?)\*\*', r'<b>\1</b>', clean_p)
                clean_p = re.sub(r'\*(.*?)\*', r'<i>\1</i>', clean_p)
                story.append(Paragraph(clean_p, body_style))
                
        story.append(Spacer(1, 20))
        
    doc.build(story)
    print("  [SUKSES] Dokumen PDF berhasil dibuat: buku-strategi-80-20.pdf")

if __name__ == "__main__":
    build_docx_rich()
    build_markdown_with_images()
    build_ipad_html_rich()
    build_pdf_rich()
    print("Semua format buku dengan pembaruan WordCraft (Bold & Italic terkonversi) dan 'by Zen' selesai dibangun!")
