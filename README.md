# 📊 Data Mining: Amaliy Mashg'ulotlar To'plami

Ushbu repozitoriy **Data Mining (Ma'lumotlarni intellektual tahlili)** fani bo'yicha 1-, 2-, 3-, 4- va 5-amaliy mashg'ulotlarning to'liq yechimlari, Python tahlil dasturlari, CSV datasetlar, interaktiv **Jupyter Notebook** (`.ipynb`) tahlillari hamda akademik talablarga to'liq mos keluvchi rasmiy **Word hisobotlari**ni (`.docx`) o'z ichiga oladi.

---

## 📁 Repozitoriy Tuzilishi

```text
├── 1-amaliyot/                           # 1-amaliyot: Ma'lumotlarni yig'ish va datasetlar yaratish
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
│   ├── 1-amaliy mashg'ulot.docx          # Asl topshiriq sharti
│   ├── 1-amaliy_mashg'ulot_hisoboti.docx # Barcha 12 ta topshiriqning to'liq Word hisoboti
│   ├── 1-amaliy_mashgulot.ipynb          # Jupyter Notebook tahlili
│   ├── 1_amaliy_bajarish.py              # Barcha datasetlarni yaratuvchi Python kodi
│   └── generate_full_word_report.py      # Word hisobot generatori
│
├── 2-amaliyot/                           # 2-amaliyot: Ma'lumotlar to'plami va uning atributlari
│   ├── datasets/                         # Tanlangan 2 ta topshiriq CSV fayllari
│   │   ├── topshiriq_07_futbol_klublar_statistikasi.csv
│   │   └── topshiriq_03_talim_talabalar_xavfi.csv
│   ├── 2-amaliy mashg'ulot.docx          # Asl topshiriq sharti
│   ├── 2-amaliy_mashg'ulot_hisoboti.docx # To'liq tayyor Word hisoboti
│   ├── 2-amaliy_mashgulot.ipynb          # Jupyter Notebook tahlili
│   ├── 2_amaliy_dastur.py                # Zamonaviy GUI (Tkinter) va CLI interfeysli asosiy dastur
│   └── generate_report_and_notebook.py   # Word hisoboti va Notebook generatori
│
├── 3-amaliyot/                           # 3-amaliyot: Atribut turlarini aniqlash va sifat tahlili
│   ├── talabalar.csv                     # 10 ta talaba bo'yicha xom ma'lumotlar to'plami
│   ├── +3-amaliyot topshirig'i.docx      # Asl topshiriq sharti
│   ├── 3-amaliy_mashg'ulot_hisoboti.docx # 4 ta topshiriq va 5 ta nazorat savoli Word hisoboti
│   ├── 3-amaliy_mashgulot.ipynb          # Jupyter Notebook tahlili
│   ├── 3_amaliy_tahlil.py                # Barcha 4 ta vazifani bajaruvchi mustaqil Python dasturi
│   └── generate_3_amaliyot_report.py     # Word hisoboti va Notebook generatori
│
├── 4-amaliyot/                           # 4-amaliyot: Ma'lumotlarni tozalash va qayta ishlash
│   ├── mijozlar.csv                      # 10 ta mijoz bo'yicha xom ma'lumotlar to'plami
│   ├── +4-amaliyot topshirig'i.docx      # Asl topshiriq sharti
│   ├── 4-amaliy_mashg'ulot_hisoboti.docx # To'liq tayyor Word hisoboti
│   ├── 4-amaliy_mashgulot.ipynb          # Jupyter Notebook tahlili
│   ├── 4_amaliy_tahlil.py                # Ma'lumotlarni tozalovchi mustaqil Python dasturi
│   └── generate_4_amaliyot_report.py     # Word hisoboti va Notebook generatori
│
│
├── 5-amaliyot/                           # 5-amaliyot: Normallashtirish va standartlashtirish
│   ├── +5-Normallashtirish_va_standartlashtirish_amaliyot.docx # Asl topshiriq sharti
│   ├── 5-amaliy_mashg'ulot_hisoboti.docx # To'liq tayyor Word hisoboti
│   ├── 5-amaliy_mashgulot.ipynb          # Jupyter Notebook tahlili
│   ├── 5_amaliy_tahlil.py                # Normallashtirish va ML modellari tahlili kodi
│   └── generate_5_amaliyot_report.py     # Word hisoboti va Notebook generatori
│
└── .gitignore                            # Vaqtinchalik fayllar uchun filtr
```

---

## 📌 Amaliy Mashg'ulotlar Qisqacha Tavsifi

### 1-amaliyot: Data Miningda ma'lumotlarni yig'ish va datasetlar yaratish
- **Mavzu:** Python (`pandas`, `numpy`) yordamida 12 ta real soha (sog'liqni saqlash, qurilish, ta'lim, reabilitatsiya, o'qitish usullari, epidemiologiya, futbol statistikasi, transport, demografiya, sport jarohatlari, yo'lovchilar qoniqishi, urbanizatsiya) bo'yicha sintetik datasetlar yaratish, tozalash va CSV formatida saqlash.

### 2-amaliyot: Ma'lumotlar to'plami va uning atributlari. Metama'lumotlar
- **Mavzu:** Tanlangan 2 ta topshiriq (Futbol klublari va Talabalar xavfi) bo'yicha 7 ta amaliy vazifa:
  1. Atributlar ro'yxati va turlari;
  2. Tasodifiy 10 ta obyekt (Namuna);
  3. O'zgaruvchilar tavsifi;
  4. Umumiy populyatsiya (Bosh to'plam) tahlili;
  5. 3 tadan gipotezalar berish va ularni ma'lumotlar (korrelyatsiya) orqali tekshirish;
  6. Diskret va uzluksiz ma'lumotlarga ajratish;
  7. Shkalalar (*nominal, ordinal, interval, ratio*) bo'yicha tasniflash.
- **Dastur:** Qulay grafik oynali (Tkinter GUI) hamda terminal (CLI) menyusiga ega universal dastur.

### 3-amaliyot: Ma’lumotlar to‘plami, atributlar va ma’lumotlar turlarini aniqlash
- **Mavzu:** «Talabalar» ma'lumotlar to'plami (10 ta talaba) tahlili:
  - 1-topshiriq: To'plamni tavsiflash (obyekt, atribut, ID roli, amaliy savollar);
  - 2-topshiriq: Atribut turlarini aniqlash jadvali (binar, ordinal, nominal, interval, nisbat);
  - 3-topshiriq: Statistik tavsif (Min, Max, Diapazon, Mean, Median, Mode);
  - 4-topshiriq: Python (pandas) orqali tekshirish, bo'sh qiymatlarni to'ldirish (imputation), 120 yoshli outlierni aniqlash va bartaraf etish;
  - 5 ta nazorat savollariga to'liq javoblar.

### 4-amaliyot: Ma’lumotlarni tozalash, yetishmayotgan qiymatlarni aniqlash va qayta ishlash
- **Mavzu:** «Do‘kon mijozlari» ma'lumotlar to'plamidagi 4 ta sifat muammolarini bartaraf etish:
  - Yetishmayotgan qiymatlarni to'ldirish (Yoshi=30, Shahar='Samarqand', Oylik_xarid=505, Xaridlar_soni=5);
  - Takroriy yozuvlarni (Dublikat 7-satr) aniqlash va o'chirish;
  - Registr nomuvofiqligini (`samarqand` -> `Samarqand`) standartlashtirish;
  - G'ayrioddiy qiymatni (5000 outlier) medianaga almashtirish;
  - Excel va Python (pandas) tahlillarini solishtirish hamda 5 ta nazorat savoli.

---


### 5-amaliyot: Ma’lumotlarni normallashtirish va standartlashtirish usullarini qo‘llash
- **Mavzu:** Belgilarni bir xil masshtabga keltirish (Min-Max, Z-score, O'nlik masshtablash) usullarini qo'lda va scikit-learn orqali tadqiq qilish hamda ularning mashinaviy o'rganish modellariga ta'siri:
  - 1-topshiriq: Xodimlar yoshi bo'yicha Min-Max [0, 1] va [-1, 1] oraliqlariga normallashtirish hamda o'nlik masshtablash;
  - 2-topshiriq: Talabalar test ballari bo'yicha Z-score standartlashtirish (o'rtacha=0, std=1 ekanligini isbotlash);
  - 3-topshiriq: Python'da MinMaxScaler va StandardScaler (Evklid masofasi va outlier ta'sirini tahlil qilish);
  - 4-topshiriq: Wine to'plamida k-NN va Qarorlar daraxti (Decision Tree) modellarining masshtablashga sezgirligi tahlili (k-NN aniqligi 72.2% dan 96.3% ga oshishi, Pipeline va Data Leakage mexanizmlari);
  - Nazorat savollariga to'liq tahliliy javoblar.

## 🛠 Texnologiyalar va Kutubxonalar

- **Python 3.12+**
- **Tkinter** — foydalanuvchi grafik interfeysi (GUI)
- **Pandas** — ma'lumotlar jadvallari (DataFrames) bilan ishlash, tozalash va to'ldirish
- **NumPy** — matematik hisob-kitoblar, korrelyatsiya va tasodifiy modellashtirish
- **python-docx** — professional formatlangan Word hisobotlarini avtomatik generatsiya qilish
- **Jupyter Notebook** — interaktiv laboratoriya tahlillari

---

## 🚀 O'rnatish va Ishga Tushirish

### 1. Kutubxonalarni o'rnatish:
```bash
pip install pandas numpy python-docx matplotlib
```

### 2. 1-amaliyotni ishga tushirish:
```bash
cd 1-amaliyot
python 1_amaliy_bajarish.py
```

### 3. 2-amaliyotni ishga tushirish:
```bash
cd 2-amaliyot

# Grafik oynada (GUI) ochish:
python 2_amaliy_dastur.py

# Yoki Konsol menyusida (CLI) ochish:
python 2_amaliy_dastur.py --cli
```

### 4. 3-amaliyotni ishga tushirish:
```bash
cd 3-amaliyot
python 3_amaliy_tahlil.py
```

### 5. 4-amaliyotni ishga tushirish:
```bash
cd 4-amaliyot
python 4_amaliy_tahlil.py
```
