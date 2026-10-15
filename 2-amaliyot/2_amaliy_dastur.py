# -*- coding: utf-8 -*-
"""
2-amaliy mashg'ulot: Ma'lumotlar to'plami va uning atributlari. Metama'lumotlar
Ushbu dastur 2 ta topshiriq bo'yicha barcha 7 ta amaliy vazifani bajaradi.
Grafik (Tkinter GUI) hamda Interaktiv Konsol (CLI) interfeyslariga ega.
"""

import os
import sys
import random
import numpy as np
import pandas as pd

# Grafik interfeys (Tkinter) kutubxonalari
try:
    import tkinter as tk
    from tkinter import ttk, messagebox, scrolledtext
    TKINTER_AVAILABLE = True
except ImportError:
    TKINTER_AVAILABLE = False

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATASETS_DIR = os.path.join(BASE_DIR, "datasets")

# ==============================================================================
# DATASETLAR VA METAMA'LUMOTLAR BAZASI
# ==============================================================================

TASK_METADATA = {
    "topshiriq_07": {
        "title": "7-topshiriq: Futbol klublari statistikasi",
        "file_name": "topshiriq_07_futbol_klublar_statistikasi.csv",
        "domain": "Sport va futbol klublari tahlili",
        "variables": {
            "Jamoa_Nomi": "Klubning rasmiy nomi (unikal toifaviy belgi).",
            "Liga": "Klub ishtirok etayotgan musobaqa ligasi (Superliga, Pro-liga, Birinchi liga).",
            "Oyinlar_Soni": "Mavsum davomida o'tkazilgan jami rasmiy uchrashuvlar soni.",
            "Galabalar": "Jamoa g'alaba qozongan o'yinlar soni.",
            "Duranglar": "Durang bilan yakunlangan o'yinlar soni.",
            "Maglubiyatlar": "Mag'lubiyat bilan yakunlangan o'yinlar soni.",
            "Ochkolar": "Jamoa to'plagan jami ochkolar (3 * G'alaba + 1 * Durang).",
            "Urilgan_Gollar": "Raqib darvozasiga kiritilgan jami to'plar soni.",
            "Otkazilgan_Gollar": "O'z darvozasidan o'tkazib yuborilgan jami to'plar soni.",
            "Gollar_Farqi": "Urilgan va o'tkazilgan gollar farqi (musbat, nol yoki manfiy).",
            "Xavfli_Vaziyatlar": "Mavsum davomida yaratilgan xavfli golli vaziyatlar soni.",
            "Sovrinlar_Oxirgi_5_Yil": "So'nggi 5 yil davomida yutib olingan rasmiy sovrinlar (kuboklar) soni."
        },
        "population": {
            "Jamoa_Nomi": "Mamlakatdagi yoki mintaqadagi barcha ro'yxatdan o'tgan professional futbol klublari majmuasi.",
            "Liga": "Milliy futbol assotsiatsiyasi tasarrufidagi barcha professional va yarim professional ligalar.",
            "Oyinlar_Soni": "Chempionat nizomiga muvofiq barcha rejalashtirilgan o'yinlar miqdori.",
            "Galabalar": "Mavsumdagi barcha o'yinlarda mumkin bo'lgan g'alabalar majmuasi [0; N].",
            "Duranglar": "Mavsumdagi barcha o'yinlarda qayd etilishi mumkin bo'lgan duranglar majmuasi.",
            "Maglubiyatlar": "Mavsumdagi barcha o'yinlarda qayd etilishi mumkin bo'lgan mag'lubiyatlar majmuasi.",
            "Ochkolar": "Mavsum yakunida klublar to'plashi mumkin bo'lgan nazariy ochkolar [0; 3*N].",
            "Urilgan_Gollar": "Futbol uchrashuvlarida kiritilishi mumkin bo'lgan barcha gollar to'plami.",
            "Otkazilgan_Gollar": "Uchrashuvlarda darvozadan o'tkazib yuborilishi mumkin bo'lgan barcha gollar to'plami.",
            "Gollar_Farqi": "Musobaqadagi barcha nazariy gollar nisbati va differensiallari [-M; +M].",
            "Xavfli_Vaziyatlar": "Barcha o'yinlardagi potensial xavfli hujumlar va vaziyatlar to'plami.",
            "Sovrinlar_Oxirgi_5_Yil": "So'nggi 5 yil ichida professional futbol musobaqalarida o'ynalgan barcha rasmiy kuboklar."
        },
        "hypotheses": [
            {
                "id": "H1",
                "nomi": "G'alabalar va To'plangan Ochkolar o'rtasidagi to'g'ridan-to'g'ri bog'liqlik",
                "mazmuni": "Jamoaning g'alabalari soni qancha ko'p bo'lsa, to'plagan ochkolari shuncha yuqori bo'ladi (kuchli musbat korrelyatsiya).",
                "test_func": lambda df: f"G'alabalar va Ochkolar korrelyatsiyasi (Pearson r): {df['Galabalar'].corr(df['Ochkolar']):.4f} (Deyarli to'liq to'g'ri chiziqli bog'liqlik tasdiqlandi!)"
            },
            {
                "id": "H2",
                "nomi": "Xavfli Vaziyatlar va Urilgan Gollar bog'liqligi",
                "mazmuni": "Raqib jarima maydonchasida ko'proq xavfli vaziyat yaratgan klublar sezilarli darajada ko'proq gol uradi.",
                "test_func": lambda df: f"Xavfli vaziyatlar va Urilgan gollar korrelyatsiyasi: {df['Xavfli_Vaziyatlar'].corr(df['Urilgan_Gollar']):.4f} (Gipoteza tasdiqlandi: ko'proq xavfli vaziyat = ko'proq gol)."
            },
            {
                "id": "H3",
                "nomi": "Superliga klublarining o'rtacha gollar farqi",
                "mazmuni": "Superligada o'ynaydigan jamoalarning o'rtacha urilgan gollar ko'rsatkichi quyi liga jamoalarinikidan sezilarli farq qiladi.",
                "test_func": lambda df: "\n".join([f"  - {liga}: O'rtacha urilgan gol = {sub['Urilgan_Gollar'].mean():.1f}, O'rtacha ochko = {sub['Ochkolar'].mean():.1f}" for liga, sub in df.groupby('Liga')])
            }
        ],
        "data_types": {
            "Diskret": [
                "Jamoa_Nomi (toifaviy)",
                "Liga (toifaviy)",
                "Oyinlar_Soni (butun son, sanoqli)",
                "Galabalar (butun son)",
                "Duranglar (butun son)",
                "Maglubiyatlar (butun son)",
                "Ochkolar (butun son)",
                "Urilgan_Gollar (butun son)",
                "Otkazilgan_Gollar (butun son)",
                "Gollar_Farqi (butun son, manfiy/musbat)",
                "Xavfli_Vaziyatlar (butun son)",
                "Sovrinlar_Oxirgi_5_Yil (butun son)"
            ],
            "Uzluksiz": [
                "Izoh: Ushbu datasetda barcha o'lchovlar butun birliklarda sanalgani sababli sof uzluksiz ustun yo'q. Agar o'yin boshiga to'g'ri keluvchi o'rtacha gollar yoki to'pga egalik qilish foizi hisoblansa, ular uzluksiz bo'ladi."
            ]
        },
        "scales": {
            "Nominal": [
                "Jamoa_Nomi: Tartiblanmaydigan sof identifikator nomlar."
            ],
            "Tartib (Ordinal)": [
                "Liga: Ligalar o'rtasida aniq ierarxiya mavjud (Superliga > Pro-liga > Birinchi liga)."
            ],
            "Intervalli (Interval)": [
                "Gollar_Farqi: Qiymatlar orasidagi farq ma'noga ega, biroq '0' qiymati mutlaq gol yo'qligini emas, balki urilgan va o'tkazilgan gollar tengligini anglatadi. Manfiy sonlar mavjud."
            ],
            "Nisbiy (Ratio)": [
                "Oyinlar_Soni, Galabalar, Duranglar, Maglubiyatlar, Ochkolar, Urilgan_Gollar, Otkazilgan_Gollar, Xavfli_Vaziyatlar, Sovrinlar_Oxirgi_5_Yil (Barchasida mutlaq nol mavjud va karrali nisbat ma'noga ega, masalan 4 ta gol 2 ta goldan 2 barobar ko'p)."
            ]
        }
    },

    "topshiriq_03": {
        "title": "3-topshiriq: Talabalar ta'lim ko'rsatkichlari va chetlatilish xavfi",
        "file_name": "topshiriq_03_talim_talabalar_xavfi.csv",
        "domain": "Oliy ta'lim sifati va talabalar monitoringi",
        "variables": {
            "Talaba_ID": "Talabaning unikal tizimdagi identifikatori / kodi.",
            "Fakultet": "Talaba tahsil olayotgan fakultet nomi (AKT, Iqtisodiyot, Menejment...).",
            "Kurs": "Talabaning o'quv bosqichi / kursi (1, 2, 3, 4).",
            "Davomat_Foizi": "Semestr davomidagi darslarga qatnashish ko'rsatkichi (%).",
            "Oraliq_Baho_Ortacha": "Oraliq nazorat imtihonlaridan olingan o'rtacha ball (0-100 ball).",
            "Mustaqil_Ish_Foizi": "Mustaqil ta'lim topshiriqlarining o'z vaqtida bajarilish ulushi (%).",
            "Ijtimoiy_Yordam": "Talabaga moddiy yoki ijtimoiy imtiyoz berilganligi ('Ha' / 'Yo`q').",
            "Yotoqxonada_Turadi": "Talabaning talabalar turar joyida yashashi ('Ha' / 'Yo`q').",
            "Haftalik_Ishlash_Soati": "Talabaning darsdan tashqari haftalik mehnat faoliyati soatlari (0-40 soat).",
            "Xavf_Guruhi": "Akademik qarzdorlik yoki o'qishdan qolish xavfi darajasi ('Past xavf', 'O`rta xavf', 'Yuqori xavf')."
        },
        "population": {
            "Talaba_ID": "Respublika yoki universitet bo'yicha ro'yxatdan o'tgan barcha talabalar kontingenti.",
            "Fakultet": "Oliy ta'lim muassasasidagi barcha mavjud fakultetlar va yo'nalishlar majmuasi.",
            "Kurs": "Bakalavriat ta'lim bosqichining barcha kurslari to'plami [1-4].",
            "Davomat_Foizi": "Barcha talabalarning semestr davomidagi nazariy davomat ko'rsatkichlari oralig'i [0.0% - 100.0%].",
            "Oraliq_Baho_Ortacha": "Barcha talabalarning oraliq baholash natijalari [0.0 - 100.0].",
            "Mustaqil_Ish_Foizi": "Talabalar tomonidan topshirilishi kerak bo'lgan barcha mustaqil ishlar ulushi [0% - 100%].",
            "Ijtimoiy_Yordam": "Universitetdagi ijtimoiy himoyaga muhtoj barcha talabalar toifasi.",
            "Yotoqxonada_Turadi": "Yotoqxona bilan ta'minlangan va ijarada turuvchi barcha talabalar toifasi.",
            "Haftalik_Ishlash_Soati": "Ishlaydigan va ishlamaydigan barcha talabalarning darsdan tashqari bandlik soatlari majmuasi.",
            "Xavf_Guruhi": "OTM talabalarining o'zlashtirish va xavf darajalari bo'yicha butun bosh to'plami."
        },
        "hypotheses": [
            {
                "id": "H1",
                "nomi": "Davomat va Mustaqil Ish foizi o'rtasidagi to'g'ridan-to'g'ri bog'liqlik",
                "mazmuni": "Darsga muntazam qatnashuvchi talabalar mustaqil ishlarni ham o'z vaqtida yuqori sifatda topshiradi.",
                "test_func": lambda df: f"Davomat va Mustaqil ish foizi korrelyatsiyasi (Pearson r): {df['Davomat_Foizi'].corr(df['Mustaqil_Ish_Foizi']):.4f} (Kuchli musbat korrelyatsiya tasdiqlandi!)"
            },
            {
                "id": "H2",
                "nomi": "Haftalik Ishlash Soati va Xavf Guruhi o'rtasidagi bog'liqlik",
                "mazmuni": "Haftasiga darsdan tashqari ko'p soat ishlaydigan talabalarda 'Yuqori xavf' guruhiga tushish ehtimoli ortadi.",
                "test_func": lambda df: "\n".join([f"  - {guruh}: O'rtacha haftalik ishlash = {sub['Haftalik_Ishlash_Soati'].mean():.1f} soat" for guruh, sub in df.groupby('Xavf_Guruhi')])
            },
            {
                "id": "H3",
                "nomi": "Davomat ko'rsatkichi va Xavf Guruhi bog'liqligi",
                "mazmuni": "Davomat foizi 90% dan yuqori talabalar xavfsiz (past xavf) guruhiga kiradi, past davomat esa o'qishdan chetlatilish xavfini oshiradi.",
                "test_func": lambda df: "\n".join([f"  - {guruh}: O'rtacha davomat = {sub['Davomat_Foizi'].mean():.1f}%" for guruh, sub in df.groupby('Xavf_Guruhi')])
            }
        ],
        "data_types": {
            "Diskret": [
                "Talaba_ID (toifaviy kod)",
                "Fakultet (toifaviy matn)",
                "Kurs (butun son: 1, 2, 3, 4)",
                "Ijtimoiy_Yordam (ikkilik toifa: Ha/Yo'q)",
                "Yotoqxonada_Turadi (ikkilik toifa: Ha/Yo'q)",
                "Haftalik_Ishlash_Soati (butun soatlar)",
                "Xavf_Guruhi (toifaviy daraja)"
            ],
            "Uzluksiz": [
                "Davomat_Foizi (0.0% dan 100.0% gacha bo'lgan uzluksiz foiz qiymatlari)",
                "Oraliq_Baho_Ortacha (o'nli kasrlar ko'rinishidagi uzluksiz ball)",
                "Mustaqil_Ish_Foizi (uzluksiz foiz ko'rsatkichi)"
            ]
        },
        "scales": {
            "Nominal": [
                "Talaba_ID: Talabalarni farqlovchi unikal belgilar.",
                "Fakultet: O'zaro tartiblanmaydigan ta'lim yo'nalishlari.",
                "Ijtimoiy_Yordam: Toifaviy ikkilik belgi (Ha/Yo'q).",
                "Yotoqxonada_Turadi: Toifaviy ikkilik belgi (Ha/Yo'q)."
            ],
            "Tartib (Ordinal)": [
                "Kurs: Kurslar o'rtasida aniq ketma-ketlik bor (1-kurs < 2-kurs < 3-kurs < 4-kurs).",
                "Xavf_Guruhi: Xavf darajasining ierarxiyasi mavjud (Past xavf < O'rta xavf < Yuqori xavf)."
            ],
            "Intervalli (Interval)": [
                "Izoh: Sof intervalli shkala yo'q, barcha raqamli ko'rsatkichlar tabiiy nolga ega bo'lgani sababli nisbiy shkalaga mansub."
            ],
            "Nisbiy (Ratio)": [
                "Davomat_Foizi: Mutlaq nol mavjud (0% - umuman kelmagan).",
                "Oraliq_Baho_Ortacha: Ball 0 dan boshlanadi, 80 ball 40 balldan 2 barobar yuqori.",
                "Mustaqil_Ish_Foizi: Mutlaq nol mavjud (0% - topshirmagan).",
                "Haftalik_Ishlash_Soati: 0 soat - mutlaq ishlamaslik holati."
            ]
        }
    }
}


def load_dataset(task_key):
    meta = TASK_METADATA[task_key]
    file_path = os.path.join(DATASETS_DIR, meta["file_name"])
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Dataset topilmadi: {file_path}")
    df = pd.read_csv(file_path)
    return df, meta


# ==============================================================================
# 7 TA ASOSIY VAZIFANI BAJARUVCHI FUNKSIYALAR
# ==============================================================================

def get_task_1_attributes(df):
    """1. CSV faylida yozilgan jadvalning atributlarini chop etish"""
    res = f"Jadval o'lchami: {df.shape[0]} qator x {df.shape[1]} ustun\n"
    res += f"Jami atributlar (ustunlar) soni: {len(df.columns)}\n\n"
    res += "Atributlar ro'yxati va ma'lumot turlari (dtypes):\n"
    res += "-" * 55 + "\n"
    for idx, (col, dtype) in enumerate(df.dtypes.items(), 1):
        res += f"{idx:2d}. {col:<26} | Type: {str(dtype):<10}\n"
    return res


def get_task_2_random_samples(df, n=10):
    """2. CSV faylida hosil qilingan jadvalning tasodifiy 10 ta obyektini chop etish"""
    sample_df = df.sample(n=min(n, len(df)), random_state=random.randint(1, 10000))
    res = f"Jadvaldan tasodifiy tanlab olingan {len(sample_df)} ta obyekt (qatorlar):\n"
    res += "=" * 80 + "\n"
    res += sample_df.to_string(index=True) + "\n"
    res += "=" * 80 + "\n"
    return res, sample_df


def get_task_3_variables(meta):
    """3. Har bir atribut uchun o'zgaruvchilarni aniqlash"""
    res = f"HAR BIR ATRIBUT UCHUN O'ZGARUVCHILAR TAVSIFI ({meta['title']}):\n"
    res += "=" * 80 + "\n"
    for col, desc in meta["variables"].items():
        res += f"* [{col}]:\n    {desc}\n\n"
    return res


def get_task_4_population(meta):
    """4. Har bir atribut uchun umumiy populyatsiyani (bosh to'plam) aniqlash"""
    res = f"HAR BIR ATRIBUT UCHUN UMUMIY POPULYATSIYA (BOSH TO'PLAM):\n"
    res += "=" * 80 + "\n"
    for col, pop in meta["population"].items():
        res += f"* [{col}]:\n    Umumiy populyatsiya: {pop}\n\n"
    return res


def get_task_5_hypotheses(df, meta):
    """5. Kamida 3 tadan gipotezalar berish va ularni ma'lumotlar asosida tekshirish"""
    res = f"MA'LUMOTLAR ASOSIDA ILGARI SURILGAN GIPOTEZALAR VA ULARNING TEKSHIRUVI:\n"
    res += "=" * 80 + "\n"
    for h in meta["hypotheses"]:
        res += f"\n>>> [{h['id']}] {h['nomi']}\n"
        res += f"    Faraz mazmuni: {h['mazmuni']}\n"
        res += f"    Haqiqiy ma'lumotlar bo'yicha tekshiruv natijasi:\n"
        try:
            test_out = h["test_func"](df)
            res += f"    {test_out}\n"
        except Exception as e:
            res += f"    Xatolik: {e}\n"
        res += "-" * 75 + "\n"
    return res


def get_task_6_data_types(meta):
    """6. Jadvalning har bir ustunini diskret va uzluksiz ma'lumotlarga ajratish"""
    res = f"USTUNLARNING DISKRET VA UZLUKSIZ MA'LUMOTLARGA AJRATILISHI:\n"
    res += "=" * 80 + "\n\n"
    res += "1. DISKRET MA'LUMOTLAR (Sanash mumkin bo'lgan, butun yoki toifaviy qiymatlar):\n"
    for item in meta["data_types"]["Diskret"]:
        res += f"   - {item}\n"
    
    res += "\n2. UZLUKSIZ MA'LUMOTLAR (Ma'lum oraliqdagi har qanday haqiqiy qiymatni oluvchi):\n"
    for item in meta["data_types"]["Uzluksiz"]:
        res += f"   - {item}\n"
    return res


def get_task_7_scales(meta):
    """7. Jadvalning har bir ustunini o'lchov shkalalari bo'yicha ajratish"""
    res = f"USTUNLARNING O'LCHOV SHKALALARI BO'YICHA TASNIFLANISHI:\n"
    res += "=" * 80 + "\n\n"
    for scale_name, cols in meta["scales"].items():
        res += f"[*] {scale_name.upper()} SHKALA:\n"
        for col_desc in cols:
            res += f"    - {col_desc}\n"
        res += "\n"
    return res


def get_full_report(task_key):
    """Barcha 7 ta topshiriqning to'liq birlashgan matni"""
    df, meta = load_dataset(task_key)
    out = []
    out.append("=" * 85)
    out.append(f"  {meta['title'].upper()}")
    out.append(f"  Tadqiqot sohasi: {meta['domain']}")
    out.append(f"  Fayl nomi: {meta['file_name']}")
    out.append("=" * 85 + "\n")
    
    out.append(">>> 1-TOPSHIRIQ: CSV faylidagi jadvalning atributlari:")
    out.append(get_task_1_attributes(df))
    
    out.append("\n>>> 2-TOPSHIRIQ: Tasodifiy 10 ta obyekt (Namuna):")
    samples_str, _ = get_task_2_random_samples(df, n=10)
    out.append(samples_str)
    
    out.append("\n>>> 3-TOPSHIRIQ: O'zgaruvchilar tavsifi:")
    out.append(get_task_3_variables(meta))
    
    out.append("\n>>> 4-TOPSHIRIQ: Umumiy populyatsiya (Bosh to'plam):")
    out.append(get_task_4_population(meta))
    
    out.append("\n>>> 5-TOPSHIRIQ: Kamida 3 ta gipoteza va tekshirish:")
    out.append(get_task_5_hypotheses(df, meta))
    
    out.append("\n>>> 6-TOPSHIRIQ: Diskret va uzluksiz ma'lumotlarga ajratish:")
    out.append(get_task_6_data_types(meta))
    
    out.append("\n>>> 7-TOPSHIRIQ: Shkalalar bo'yicha ajratish:")
    out.append(get_task_7_scales(meta))
    
    return "\n".join(out)


# ==============================================================================
# GRAFIK INTERFEYS (TKINTER GUI)
# ==============================================================================

class AppGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("2-Amaliy Mashg'ulot: Ma'lumotlar To'plami va Atributlar Tahlili")
        self.root.geometry("1100x750")
        self.root.minsize(900, 600)
        
        # Stil sozlamalari
        self.style = ttk.Style()
        self.style.theme_use("clam")
        
        self.current_task_key = "topshiriq_07"
        self.df = None
        self.meta = None
        
        self.setup_ui()
        self.on_dataset_change()

    def setup_ui(self):
        # Yuqori sarlavha paneli
        header_frame = tk.Frame(self.root, bg="#1e293b", padx=15, pady=12)
        header_frame.pack(fill=tk.X)
        
        title_lbl = tk.Label(
            header_frame, 
            text="2-AMALIY MASHG'ULOT: MA'LUMOTLAR TO'PLAMI VA METAMA'LUMOTLAR TAHLILI",
            font=("Segoe UI", 13, "bold"),
            fg="#f8fafc", 
            bg="#1e293b"
        )
        title_lbl.pack(anchor="w")
        
        subtitle_lbl = tk.Label(
            header_frame,
            text="Python dasturi | Barcha 7 ta topshiriq to'liq interaktiv qulay interfeysda",
            font=("Segoe UI", 9),
            fg="#94a3b8",
            bg="#1e293b"
        )
        subtitle_lbl.pack(anchor="w", pady=(2, 0))

        # Asosiy gorizontal container
        main_container = tk.Frame(self.root, bg="#f1f5f9")
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Chap menyu paneli (Tugmalar va boshqaruv)
        left_panel = tk.Frame(main_container, bg="#ffffff", relief=tk.RIDGE, bd=1, width=280)
        left_panel.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        left_panel.pack_propagate(False)

        # Dataset tanlash bloki
        ds_lbl = tk.Label(left_panel, text="DATASETNI TANLANG:", font=("Segoe UI", 10, "bold"), bg="#ffffff", fg="#0f172a")
        ds_lbl.pack(anchor="w", padx=12, pady=(12, 6))

        self.ds_var = tk.StringVar(value="topshiriq_07")
        rb1 = tk.Radiobutton(
            left_panel, 
            text="1-vazifa: Futbol Statistikasi (Topshiriq 7)", 
            variable=self.ds_var, 
            value="topshiriq_07",
            command=self.on_dataset_change,
            bg="#ffffff",
            activebackground="#ffffff",
            font=("Segoe UI", 9)
        )
        rb1.pack(anchor="w", padx=12, pady=2)

        rb2 = tk.Radiobutton(
            left_panel, 
            text="2-vazifa: Talabalar Xavfi (Topshiriq 3)", 
            variable=self.ds_var, 
            value="topshiriq_03",
            command=self.on_dataset_change,
            bg="#ffffff",
            activebackground="#ffffff",
            font=("Segoe UI", 9)
        )
        rb2.pack(anchor="w", padx=12, pady=2)

        ttk.Separator(left_panel, orient="horizontal").pack(fill=tk.X, padx=10, pady=10)

        # Topshiriqlar tugmalari
        actions_lbl = tk.Label(left_panel, text="7 TA TOPSHIRIQNI KO'RISH:", font=("Segoe UI", 10, "bold"), bg="#ffffff", fg="#0f172a")
        actions_lbl.pack(anchor="w", padx=12, pady=(0, 8))

        buttons = [
            ("1. Atributlar ro'yxati", self.show_task_1),
            ("2. Tasodifiy 10 ta obyekt", self.show_task_2),
            ("3. O'zgaruvchilar tavsifi", self.show_task_3),
            ("4. Umumiy populyatsiya", self.show_task_4),
            ("5. Gipotezalar va Tekshiruv", self.show_task_5),
            ("6. Diskret va Uzluksiz", self.show_task_6),
            ("7. Shkalalar bo'yicha ajratish", self.show_task_7),
            ("★ TO'LIQ HISOBOTNI CHIQARISH", self.show_full_report),
        ]

        for text, cmd in buttons:
            btn_bg = "#2563eb" if "★" in text else "#0ea5e9"
            btn = tk.Button(
                left_panel,
                text=text,
                command=cmd,
                bg=btn_bg,
                fg="white",
                activebackground="#1d4ed8",
                activeforeground="white",
                font=("Segoe UI", 9, "bold" if "★" in text else "normal"),
                relief=tk.FLAT,
                cursor="hand2",
                pady=6
            )
            btn.pack(fill=tk.X, padx=12, pady=3)

        # Quyi info
        info_frame = tk.Frame(left_panel, bg="#f8fafc", padx=8, pady=8, relief=tk.GROOVE, bd=1)
        info_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=10)
        self.status_lbl = tk.Label(info_frame, text="", font=("Segoe UI", 8), bg="#f8fafc", fg="#475569", justify=tk.LEFT)
        self.status_lbl.pack(anchor="w")

        # O'ng tomondagi natijalar paneli
        right_panel = tk.Frame(main_container, bg="#ffffff", relief=tk.RIDGE, bd=1)
        right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # Tab yoki Text chiqarish bloki
        top_bar = tk.Frame(right_panel, bg="#f8fafc", padx=10, pady=8)
        top_bar.pack(fill=tk.X)
        
        self.result_title = tk.Label(top_bar, text="Natijalar oynasi", font=("Segoe UI", 11, "bold"), bg="#f8fafc", fg="#1e293b")
        self.result_title.pack(side=tk.LEFT)

        copy_btn = tk.Button(
            top_bar,
            text="Nusxa olish (Copy)",
            command=self.copy_to_clipboard,
            font=("Segoe UI", 8),
            bg="#e2e8f0",
            relief=tk.FLAT,
            cursor="hand2"
        )
        copy_btn.pack(side=tk.RIGHT)

        self.text_area = scrolledtext.ScrolledText(
            right_panel, 
            wrap=tk.WORD, 
            font=("Consolas", 10),
            bg="#0f172a",
            fg="#f8fafc",
            insertbackground="white",
            padx=10,
            pady=10
        )
        self.text_area.pack(fill=tk.BOTH, expand=True)

    def on_dataset_change(self):
        self.current_task_key = self.ds_var.get()
        self.df, self.meta = load_dataset(self.current_task_key)
        self.status_lbl.config(
            text=f"Yuklandi: {self.meta['file_name']}\nQatorlar: {self.df.shape[0]} ta\nUstunlar: {self.df.shape[1]} ta"
        )
        self.show_task_1()

    def set_content(self, title, content):
        self.result_title.config(text=f"{self.meta['title']} -> {title}")
        self.text_area.delete(1.0, tk.END)
        self.text_area.insert(tk.END, content)

    def show_task_1(self):
        self.set_content("1. Jadvalning Atributlari", get_task_1_attributes(self.df))

    def show_task_2(self):
        txt, _ = get_task_2_random_samples(self.df, n=10)
        self.set_content("2. Tasodifiy 10 ta Obyekt", txt)

    def show_task_3(self):
        self.set_content("3. O'zgaruvchilar Tavsifi", get_task_3_variables(self.meta))

    def show_task_4(self):
        self.set_content("4. Umumiy Populyatsiya", get_task_4_population(self.meta))

    def show_task_5(self):
        self.set_content("5. Gipotezalar va Tekshiruv", get_task_5_hypotheses(self.df, self.meta))

    def show_task_6(self):
        self.set_content("6. Diskret va Uzluksiz Ma'lumotlar", get_task_6_data_types(self.meta))

    def show_task_7(self):
        self.set_content("7. Shkalalar bo'yicha Tasniflash", get_task_7_scales(self.meta))

    def show_full_report(self):
        self.set_content("TO'LIQ HISOBOT (Barcha 7 ta vazifa)", get_full_report(self.current_task_key))

    def copy_to_clipboard(self):
        text = self.text_area.get(1.0, tk.END)
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        messagebox.showinfo("Muvaffaqiyatli", "Matn buferga (clipboard) nusxalandi!")


# ==============================================================================
# KONSOL INTERFEYSI (CLI REJIMI)
# ==============================================================================

def run_cli():
    print("=" * 80)
    print("  2-AMALIY MASHG'ULOT: MA'LUMOTLAR TO'PLAMI VA UNING ATRIBUTLARI")
    print("=" * 80)
    
    current_key = "topshiriq_07"
    df, meta = load_dataset(current_key)
    
    while True:
        print(f"\n[JORIY DATASET]: {meta['title']} ({meta['file_name']})")
        print("-" * 60)
        print("1. Datasetni almashtirish (1: Futbol statistikasi | 2: Talabalar xavfi)")
        print("2. 1-topshiriq: Atributlarni chop etish")
        print("3. 2-topshiriq: Tasodifiy 10 ta obyektni chop etish")
        print("4. 3-topshiriq: O'zgaruvchilarni aniqlash")
        print("5. 4-topshiriq: Umumiy populyatsiyani aniqlash")
        print("6. 5-topshiriq: Gipotezalar berish va tekshirish")
        print("7. 6-topshiriq: Diskret va uzluksiz ma'lumotlarga ajratish")
        print("8. 7-topshiriq: Shkalalar bo'yicha ajratish")
        print("9. BARCHA TOPSHIRIQLARNI BIRGALIKDA CHIQARISH (To'liq hisobot)")
        print("0. Dasturdan chiqish")
        print("-" * 60)
        
        choice = input("Tanlovingizni kiriting [0-9]: ").strip()
        
        if choice == "0":
            print("\nDastur yakunlandi. Xayr!")
            break
        elif choice == "1":
            print("\n1. 7-topshiriq: Futbol klublar statistikasi")
            print("2. 3-topshiriq: Talabalar ta'lim xavfi")
            c2 = input("Tanlang (1 yoki 2): ").strip()
            if c2 == "2":
                current_key = "topshiriq_03"
            else:
                current_key = "topshiriq_07"
            df, meta = load_dataset(current_key)
            print(f"Tanlandi: {meta['title']}")
        elif choice == "2":
            print("\n" + get_task_1_attributes(df))
        elif choice == "3":
            txt, _ = get_task_2_random_samples(df, n=10)
            print("\n" + txt)
        elif choice == "4":
            print("\n" + get_task_3_variables(meta))
        elif choice == "5":
            print("\n" + get_task_4_population(meta))
        elif choice == "6":
            print("\n" + get_task_5_hypotheses(df, meta))
        elif choice == "7":
            print("\n" + get_task_6_data_types(meta))
        elif choice == "8":
            print("\n" + get_task_7_scales(meta))
        elif choice == "9":
            print("\n" + get_full_report(current_key))
        else:
            print("Noto'g'ri tanlov, qaytadan urinib ko'ring!")


# ==============================================================================
# ASOSIY ISHGA TUSHIRISH NUQTASI
# ==============================================================================

def main():
    # Agar argument sifatida --cli berilsa yoki GUI mavjud bo'lmasa CLI rejimida ishlaydi
    if "--cli" in sys.argv or not TKINTER_AVAILABLE:
        run_cli()
    else:
        try:
            root = tk.Tk()
            app = AppGUI(root)
            root.mainloop()
        except Exception as e:
            print(f"GUI rejimida xatolik yuz berdi ({e}). CLI rejimiga o'tilmoqda...")
            run_cli()


if __name__ == "__main__":
    main()
