# 🏆 OlimpMath — Sistem Seleksi Olimpiade Matematika

> Sistem cerdas berbasis Machine Learning untuk seleksi tahap awal calon peserta Olimpiade Matematika (Olimpmath)

![Python](https://img.shields.io/badge/Python-3.10+-3776AB?style=for-the-badge&logo=python&logoColor=white)
![Streamlit](https://img.shields.io/badge/Streamlit-1.32-FF4B4B?style=for-the-badge&logo=streamlit&logoColor=white)
![scikit-learn](https://img.shields.io/badge/scikit--learn-1.4-F7931E?style=for-the-badge&logo=scikit-learn&logoColor=white)
![Accuracy](https://img.shields.io/badge/Akurasi-97%25-48BB78?style=for-the-badge)

---

## 📋 Deskripsi

**OlimpMath** adalah sistem seleksi tahap awal calon peserta Olimpiade Matematika yang dibangun menggunakan metode klasifikasi **Random Forest**. Sistem ini menganalisis 6 dimensi kemampuan matematika siswa dan menghasilkan:

- **Status Kesiapan**: Siap Olimpiade / Potensial / Tidak Siap
- **Review** personal berdasarkan profil skor
- **Rekomendasi** tindak lanjut yang spesifik
- **Saran** pengembangan kemampuan

---

## ✨ Fitur Utama

| Fitur | Deskripsi |
|-------|-----------|
| 👤 **Input Satu Siswa** | Analisis individual dengan slider interaktif |
| 📂 **Upload Massal** | Upload Excel berisi banyak siswa sekaligus |
| 📥 **Export Excel** | Download hasil analisis dalam format .xlsx yang rapi |
| 📊 **Visualisasi** | Profil skor dengan bar chart berwarna |
| 🤖 **AI-Powered** | Random Forest dengan akurasi 97% |

---

## 📐 Dimensi yang Dianalisis

| Komponen | Keterangan |
|----------|-----------|
| **Numerasi Aljabar** | Kemampuan aljabar & persamaan |
| **Numerasi Geometri** | Kemampuan geometri & ruang |
| **Numerasi Bilangan** | Kemampuan bilangan & aritmetika |
| **Data & Ketidakpastian** | Statistika & probabilitas |
| **Skor Menalar** | Penalaran logis & analitis |
| **Skor Literasi** | Literasi & pemecahan masalah |

---

## 🎯 Kategori Status

| Status | Kriteria | Warna |
|--------|----------|-------|
| 🏆 **Siap Olimpiade** | Rata-rata ≥ 80 DAN semua skor ≥ 60 | Hijau |
| ⭐ **Potensial** | Rata-rata ≥ 65 DAN skor min ≥ 45 | Kuning |
| 📋 **Tidak Siap** | Selain kondisi di atas | Merah |

---

## 🚀 Cara Menjalankan

### Prasyarat
- Python 3.10+
- pip

### Langkah 1: Clone Repository

```bash
git clone https://github.com/USERNAME/olimpmath.git
cd olimpmath
```

### Langkah 2: Install Dependencies

```bash
cd streamlit_app
pip install -r requirements.txt
```

### Langkah 3: Siapkan File

Pastikan struktur folder seperti ini:
```
streamlit_app/
├── app.py
├── requirements.txt
├── .streamlit/
│   └── config.toml
├── data/
│   ├── dataset.xlsx          ← File dataset training
│   └── template_upload.xlsx  ← Template upload siswa
└── models/
    └── model_olimpmath.pkl   ← Model (dibuat otomatis jika tidak ada)
```

> **Catatan:** Model akan dibuat otomatis saat pertama kali dijalankan jika `dataset.xlsx` tersedia.

### Langkah 4: Jalankan Aplikasi

```bash
streamlit run app.py
```

Buka browser di `http://localhost:8501`

---

## ☁️ Deploy ke Streamlit Cloud

1. **Fork** repository ini ke akun GitHub kamu
2. Buka [share.streamlit.io](https://share.streamlit.io)
3. Klik **"New app"**
4. Pilih repository ini, branch `main`, file `streamlit_app/app.py`
5. Klik **"Deploy!"**

> ⚠️ **Penting untuk deploy:** Upload file `dataset.xlsx` ke folder `streamlit_app/data/` sebelum push ke GitHub, agar model dapat dilatih saat deploy.

---

## 📓 Google Colab

Gunakan notebook `notebooks/OlimpMath_Training.ipynb` untuk:
- Eksplorasi dan visualisasi data
- Training & evaluasi model
- Export model ke file `.pkl`

[![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/USERNAME/olimpmath/blob/main/notebooks/OlimpMath_Training.ipynb)

---

## 📁 Struktur Proyek

```
olimpmath/
├── 📓 notebooks/
│   └── OlimpMath_Training.ipynb    # Google Colab notebook
├── 🖥️ streamlit_app/
│   ├── app.py                       # Main Streamlit app
│   ├── requirements.txt             # Python dependencies
│   ├── .streamlit/
│   │   └── config.toml             # Streamlit theme config
│   ├── data/
│   │   ├── dataset.xlsx            # Dataset training (30.000 data)
│   │   └── template_upload.xlsx    # Template untuk upload massal
│   └── models/
│       └── model_olimpmath.pkl     # Trained model (auto-generated)
└── README.md
```

---

## 🤖 Detail Model

| Parameter | Nilai |
|-----------|-------|
| Algoritma | Random Forest Classifier |
| n_estimators | 200 |
| max_depth | 10 |
| min_samples_leaf | 5 |
| Data Training | 24.000 sampel (80%) |
| Data Testing | 6.000 sampel (20%) |
| Akurasi | **97%** |
| Cross-validation | 5-Fold Stratified |

---

## 📦 Dependencies

```
streamlit==1.32.0
pandas==2.2.1
numpy==1.26.4
scikit-learn==1.4.1
openpyxl==3.1.2
xlrd==2.0.1
```

---

## 👨‍💻 Pengembang

Dibuat sebagai tugas mata kuliah **Data Mining**  
Program Studi: **Pendidikan Matematika**

---

## 📄 Lisensi

MIT License — bebas digunakan untuk keperluan pendidikan.
