# Data Mining: Amaliy Mashg'ulotlar

Ushbu repozitoriy **Data Mining (Ma'lumotlarni intellektual tahlili)** fani bo'yicha amaliy mashg'ulotlar, Python dasturlash kodlari, datasetlar va tayyor hisobotlarni o'z ichiga oladi.

---

## 📁 Repozitoriy tuzilishi

```
├── 1-amaliyot/
│   ├── datasets/                         # 12 ta soha bo'yicha generatsiya qilingan CSV datasetlar
│   │   ├── topshiriq_01_sogliqni_saqlash_xarajatlar.csv
│   │   ├── topshiriq_02_qurilish_loyihalari.csv
│   │   ├── topshiriq_03_talim_talabalar_xavfi.csv
│   │   ├── topshiriq_04_bemor_reabilitatsiyasi.csv
│   │   ├── topshiriq_05_oqitish_usullari_samaradorligi.csv
│   │   ├── topshiriq_06_epidemiologiya_atrof_muhit.csv
│   │   ├── topshiriq_07_futbol_klublar_statistikasi.csv
│   │   ├── topshiriq_08_jamoat_transporti_samaradorligi.csv
│   │   ├── topshiriq_09_demografiya_saylov_tendensiyasi.csv
│   │   ├── topshiriq_10_sportchilar_jaroxatlari.csv
│   │   ├── topshiriq_11_transport_yolovchilar_qoniqishi.csv
│   │   └── topshiriq_12_urbanizatsiya_tendensiyalari.csv
│   ├── 1-amaliy mashg'ulot.docx          # Topshiriq sharti
│   ├── 1-amaliy_mashg'ulot_hisoboti.docx # To'liq tayyor Word hisoboti
│   ├── 1-amaliy_mashgulot.ipynb          # Jupyter Notebook tahlili
│   └── 1_amaliy_bajarish.py              # Barcha datasetlarni yaratuvchi Python kodi
│
├── 2-amaliyot/                           # 2-amaliyot materiallari uchun
├── 3-amaliyot/
│   └── +3-amaliyot topshirig'i.docx      # 3-amaliyot topshiriq sharti
└── 4-amaliyot/
    └── +4-amaliyot topshirig'i.docx      # 4-amaliyot topshiriq sharti
```

---

## 🛠 Ishlatilgan texnologiyalar
- **Python 3.12**
- **Pandas** — ma'lumotlar jadvallari (DataFrames) bilan ishlash
- **NumPy** — matematik va statistik hisob-kitoblar, tasodifiy modellashtirish
- **python-docx** — hisobot hujjatlarini avtomatik generatsiya qilish
- **Jupyter Notebook** — interaktiv tahlil

---

## 🚀 Ishga tushirish
Kutubxonalarni o'rnatish:
```bash
pip install pandas numpy python-docx matplotlib
```

1-amaliyot kodini ishga tushirish:
```bash
cd 1-amaliyot
python 1_amaliy_bajarish.py
```
