# -*- coding: utf-8 -*-
"""
3-amaliy mashg'ulot: Ma’lumotlar to‘plami, atributlar va ma’lumotlar turlarini aniqlash hamda tahlil qilish
Ushbu skript topshiriqdagi barcha 4 ta asosiy topshiriqni va nazorat savollarini to'liq bajaradi.
"""

import os
import numpy as np
import pandas as pd

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
CSV_PATH = os.path.join(CURRENT_DIR, "talabalar.csv")

def main():
    print("=" * 80)
    print("  3-AMALIY MASHG'ULOT: MA'LUMOTLAR TO'PLAMI VA ATRIBUTLAR TAHLILI")
    print("=" * 80)

    # -------------------------------------------------------------------------
    # 1-TOPSHIRIQ. MA'LUMOTLAR TO'PLAMINI TAVSIFLASH
    # -------------------------------------------------------------------------
    print("\n" + "=" * 80)
    print("1-TOPSHIRIQ. MA'LUMOTLAR TO'PLAMINI TAVSIFLASH")
    print("=" * 80)
    
    df = pd.read_csv(CSV_PATH)
    
    print(f"1. Obyektlar va atributlar soni:")
    print(f"   - Obyektlar (satrlar) soni: {df.shape[0]} ta")
    print(f"   - Atributlar (ustunlar) soni: {df.shape[1]} ta")
    print(f"\n2. Bu to'plamda obyekt nimani ifodalaydi?")
    print(f"   Javob: Har bir satr (obyekt) alohida bir talabani, uning shaxsiy, demografik va akademik ko'rsatkichlarini ifodalaydi.")
    print(f"\n3. Qaysi atribut faqat identifikatsiyalash uchun xizmat qiladi va tahlilda qatnashmaydi? Nima uchun?")
    print(f"   Javob: 'ID' atributi. Chunki u faqat tizimda talabani unikal farqlash (sun'iy tartib raqami) vazifasini bajaradi.")
    print(f"   Unda hech qanday tahliliy, statistik yoki sabab-oqibat bog'liqligi (semantik ma'no) mavjud emas.")
    print(f"\n4. Ushbu to'plam yordamida javob topish mumkin bo'lgan 2 ta amaliy savol:")
    print(f"   Savol 1: Talabaning akademik o'zlashtirishi (GPA) uning stipendiya olishiga qanchalik ta'sir qiladi?")
    print(f"   Savol 2: Qaysi ta'lim yo'nalishi (Sun'iy intellekt, Dasturiy injiniring, Kompyuter injiniringi) talabalari eng yuqori o'rtacha GPA ga ega?")

    # -------------------------------------------------------------------------
    # 2-TOPSHIRIQ. ATRIBUT TURLARINI ANIQLASH
    # -------------------------------------------------------------------------
    print("\n" + "=" * 80)
    print("2-TOPSHIRIQ. ATRIBUT TURLARINI ANIQLASH JADVALI")
    print("=" * 80)
    
    attr_table = [
        {"Atribut": "Jinsi", "Turi": "Binar (simmetrik)", "Diskret/Uzluksiz": "Diskret", "Asoslash": "Faqat 2 ta teng huquqli toifa (Erkak / Ayol) mavjud."},
        {"Atribut": "Yoshi", "Turi": "Nisbat (Ratio)", "Diskret/Uzluksiz": "Diskret", "Asoslash": "Haqiqiy mutlaq nol mavjud (tug'ilish nuqtasi), karrali nisbat ma'noga ega (20 yosh 10 yoshdan 2 marta katta)."},
        {"Atribut": "Kurs", "Turi": "Ordinal (Tartib)", "Diskret/Uzluksiz": "Diskret", "Asoslash": "Kurslar orasida ketma-ketlik va ierarxiya mavjud (1 < 2 < 3 < 4), lekin bilim farqi son bilan o'lchanmaydi."},
        {"Atribut": "Yo‘nalish", "Turi": "Nominal", "Diskret/Uzluksiz": "Diskret", "Asoslash": "Mutaxassislik nomlari bo'lib, ular orasida tabiiy tartib yoki matematik munosabat yo'q."},
        {"Atribut": "GPA", "Turi": "Nisbat (Ratio)", "Diskret/Uzluksiz": "Uzluksiz", "Asoslash": "Haqiqiy mutlaq nol (0.0) mavjud, istalgan o'nli kasr qiymatni oladi, 4.0 ball 2.0 dan 2 barobar yuqori."},
        {"Atribut": "Stipendiya", "Turi": "Binar (asimmetrik)", "Diskret/Uzluksiz": "Diskret", "Asoslash": "Faqat 2 ta mumkin bo'lgan holat mavjud: 'Ha' (oladi) yoki 'Yo‘q' (olmaydi)."},
        {"Atribut": "Baho darajasi", "Turi": "Ordinal (Tartib)", "Diskret/Uzluksiz": "Diskret", "Asoslash": "Darajalar bo'yicha tabiiy tartib mavjud: Qoniqarli < Yaxshi < A’lo."},
        {"Atribut": "Tug‘ilgan yil", "Turi": "Interval", "Diskret/Uzluksiz": "Diskret", "Asoslash": "Taqvimiy shkala bo'lib, mutlaq nol mavjud emas (0-yil vaqt boshlanishi emas). Yillar farqi ma'noli, ammo nisbati ma'nosiz."}
    ]
    df_attr = pd.DataFrame(attr_table)
    print(df_attr.to_string(index=False))

    print("\nQo'shimcha savollarga javoblar:")
    print("• 'Kurs' atributini nima uchun nominal emas, balki ordinal deb hisoblaymiz?")
    print("  Javob: Chunki kurslar o'rtasida aniq ta'lim bosqichlari ketma-ketligi va tartibi mavjud (1-kurs < 2-kurs < 3-kurs < 4-kurs).")
    print("• 'Tug‘ilgan yil' va 'Yoshi' atributlarining turlari nima bilan farq qiladi?")
    print("  Javob: 'Tug‘ilgan yil' - Interval shkalada (mutlaq nol yo'q, 2000-yil 1000-yildan 2 barobar katta deb bo'lmaydi).")
    print("         'Yoshi' - Nisbat (Ratio) shkalasida (haqiqiy mutlaq nol - tug'ilish payti bor, 20 yosh 10 yoshdan 2 barobar katta deb nisbat olinadi).")

    # -------------------------------------------------------------------------
    # 3-TOPSHIRIQ. ATRIBUTLARNING STATISTIK TAVSIFI
    # -------------------------------------------------------------------------
    print("\n" + "=" * 80)
    print("3-TOPSHIRIQ. ATRIBUTLARNING STATISTIK TAVSIFI (HISOB-KITOBLAR)")
    print("=" * 80)
    
    # 1. Yoshi va GPA statistikasi
    # Yoshi (barcha 10 ta, jumladan 120 anomaliyasi bilan)
    yosh_all = df['Yoshi'].dropna()
    yosh_clean = df[df['Yoshi'] < 100]['Yoshi'].dropna()
    gpa_vals = df['GPA'].dropna()

    stats_num = [
        {"Ko'rsatkich": "Minimum (Min)", "Yoshi (Asl 120 bilan)": yosh_all.min(), "Yoshi (Outliersiz)": yosh_clean.min(), "GPA (9 ta talaba)": gpa_vals.min()},
        {"Ko'rsatkich": "Maksimum (Max)", "Yoshi (Asl 120 bilan)": yosh_all.max(), "Yoshi (Outliersiz)": yosh_clean.max(), "GPA (9 ta talaba)": gpa_vals.max()},
        {"Ko'rsatkich": "Diapazon (Max - Min)", "Yoshi (Asl 120 bilan)": yosh_all.max() - yosh_all.min(), "Yoshi (Outliersiz)": yosh_clean.max() - yosh_clean.min(), "GPA (9 ta talaba)": round(gpa_vals.max() - gpa_vals.min(), 2)},
        {"Ko'rsatkich": "O'rtacha qiymat (Mean)", "Yoshi (Asl 120 bilan)": round(yosh_all.mean(), 2), "Yoshi (Outliersiz)": round(yosh_clean.mean(), 2), "GPA (9 ta talaba)": round(gpa_vals.mean(), 2)},
        {"Ko'rsatkich": "Mediana (Median)", "Yoshi (Asl 120 bilan)": yosh_all.median(), "Yoshi (Outliersiz)": yosh_clean.median(), "GPA (9 ta talaba)": gpa_vals.median()},
        {"Ko'rsatkich": "Moda (Mode)", "Yoshi (Asl 120 bilan)": f"{yosh_all.mode()[0]} (3 marta)", "Yoshi (Outliersiz)": f"{yosh_clean.mode()[0]} (3 marta)", "GPA (9 ta talaba)": "Moda yo'q (barchasi 1 martadan)"}
    ]
    df_stats_num = pd.DataFrame(stats_num)
    print("\nSonli atributlar (Yoshi va GPA) tavsifi:")
    print(df_stats_num.to_string(index=False))

    # 2. Toifaviy atributlar: Yo'nalish, Stipendiya, Baho darajasi
    print("\nToifaviy atributlarning chastotalari va modasi:")
    print("-" * 55)
    print("• Yo'nalish (9 ta ma'lumot mavjud, 1 ta bo'sh):")
    y_counts = df['Yo‘nalish'].value_counts()
    for cat, cnt in y_counts.items():
        print(f"  - {cat}: {cnt} ta ({cnt/9*100:.1f}%)")
    print(f"  >>> Modasi: {df['Yo‘nalish'].mode()[0]} ({y_counts.max()} marta)")

    print("\n• Stipendiya (10 ta talaba):")
    st_counts = df['Stipendiya'].value_counts()
    for cat, cnt in st_counts.items():
        print(f"  - {cat}: {cnt} ta ({cnt/10*100:.1f}%)")
    print(f"  >>> Modasi: {df['Stipendiya'].mode()[0]} va {df['Stipendiya'].mode()[1] if len(df['Stipendiya'].mode()) > 1 else ''} (teng 5 tadan, bimodallik)")

    print("\n• Baho darajasi (10 ta talaba):")
    b_counts = df['Baho darajasi'].value_counts()
    for cat, cnt in b_counts.items():
        print(f"  - {cat}: {cnt} ta ({cnt/10*100:.1f}%)")
    print(f"  >>> Modasi: {df['Baho darajasi'].mode()[0]} ({b_counts.max()} marta)")

    # 3. Baho darajasi medianasi va o'rtacha qiymat haqida
    print("\n• 'Baho darajasi' atributining medianasi:")
    print("  Darajalar tartibi: Qoniqarli (2 ta), Yaxshi (5 ta), A’lo (3 ta)")
    print("  Ketma-ketlik: [Qoniqarli, Qoniqarli, Yaxshi, Yaxshi, *Yaxshi*, *Yaxshi*, Yaxshi, A’lo, A’lo, A’lo]")
    print("  Javob: Mediana = 'Yaxshi' (5 va 6-o'rindagi qiymat).")
    print("• Nima uchun 'Baho darajasi' atributi uchun o'rtacha qiymatni hisoblab bo'lmaydi?")
    print("  Javob: Chunki bu matnli toifaviy (ordinal) atribut. Unda sonli qiymatlar bo'lmagani sababli qo'shish va bo'lish arifmetik amallarini bajarish matematik jihatdan ma'noga ega emas.")

    # -------------------------------------------------------------------------
    # 4-TOPSHIRIQ. PYTHON YORDAMIDA TAHLIL VA MA'LUMOTLAR SIFATINI TEKSHIRISH
    # -------------------------------------------------------------------------
    print("\n" + "=" * 80)
    print("4-TOPSHIRIQ. PYTHON (PANDAS) ASOSIDAGI TAHLIL VA MA'LUMOTLAR SIFATI")
    print("=" * 80)
    
    print("1. df.shape (O'lcham):", df.shape)
    print("\n2. df.dtypes (Pandas ma'lumot turlari):\n", df.dtypes)
    print("\n3. df.describe() (Statistik ko'rsatkichlar):\n", df.describe())
    print("\n4. df.isnull().sum() (Bo'sh qiymatlar soni):\n", df.isnull().sum())
    print("\n5. df['Yo‘nalish'].value_counts():\n", df["Yo‘nalish"].value_counts())

    print("\n" + "-" * 80)
    print("PANDAS NATIJALARI ASOSIDAGI XULOSA VA TAHLILLAR:")
    print("-" * 80)
    print("a) Pandas tiplari bilan nazariy turlarni solishtirish:")
    print("   - ID, Yoshi, Kurs, Tug‘ilgan yil -> Pandas 'int64' deb aniqlagan (Ammo Kurs va Tug‘ilgan yil faqat oddiy son emas, ordinal va interval shkalaga ega).")
    print("   - GPA -> 'float64' (Uzluksiz nisbat shkalasi).")
    print("   - Ism, Jinsi, Yo‘nalish, Stipendiya, Baho darajasi -> Pandas 'object' (matn) deb ko'radi. Biroq nazariy jihatdan ular nominal, binar va ordinal kabi turlicha xossalarga ega.")

    print("\nb) Bo'sh qiymatlarni to'ldirish (Imputation) taklifi:")
    print("   - Bekzod (5-satr) GPA bo'sh -> uni talabalarning o'rtacha GPA si (3.77) yoki 'Yaxshi' baholilar o'rtacha GPA si (3.68) bilan to'ldirish mantiqiy.")
    print("   - Madina (10-satr) Yo'nalishi bo'sh -> eng ko'p uchraydigan moda qiymati ('Sun’iy intellekt') bilan to'ldirish mumkin.")

    print("\nc) Shubhali (Outlier) qiymat tahlili:")
    print("   - Otabek (9-satr) Yoshi = 120 ko'rsatilgan. Bu yaqqol xatolik (inson 120 yoshda 2-kurs talabasi bo'lolmaydi, ehtimol 20 yosh bo'lgan).")
    print(f"   - 120 yosh bilan o'rtacha yosh: {yosh_all.mean():.2f} yosh.")
    print(f"   - 120 yoshni olib tashlagandagi o'rtacha yosh: {yosh_clean.mean():.2f} yosh.")
    print(f"   - O'zgarish: Outlier o'rtacha yoshni birdaniga {yosh_all.mean() - yosh_clean.mean():.2f} yoshga sun'iy ravishda ko'tarib yuborgan!")

    # Ma'lumotlarni tozalash namoyishi
    df_clean = df.copy()
    # 1. Bekzod GPA sini Yaxshi baholilar o'rtachasi bilan to'ldirish
    yaxshi_gpa_mean = df_clean[df_clean['Baho darajasi'] == 'Yaxshi']['GPA'].mean()
    df_clean['GPA'] = df_clean['GPA'].fillna(round(yaxshi_gpa_mean, 2))
    # 2. Madina Yo'nalishini moda bilan to'ldirish
    df_clean['Yo‘nalish'] = df_clean['Yo‘nalish'].fillna(df_clean['Yo‘nalish'].mode()[0])
    # 3. Otabek yoshini to'g'irlash (2005-yilda tug'ilgan, demak 20 yosh)
    df_clean.loc[df_clean['ID'] == 9, 'Yoshi'] = 2025 - 2005 # 20 yosh

    print("\nTozalangan va to'ldirilgan jadval (Cleaned DataFrame):")
    print(df_clean.to_string(index=False))

    # -------------------------------------------------------------------------
    # 5. NAZORAT SAVOLLARI VA JAVOBLARI
    # -------------------------------------------------------------------------
    print("\n" + "=" * 80)
    print("5. NAZORAT SAVOLLARI VA ULARNING JAVOBLARI")
    print("=" * 80)
    qa = [
        ("1. Obyekt va atribut tushunchalarini misollar bilan tushuntiring.",
         "Obyekt - o'rganilayotgan hodisa, jarayon yoki birlik (masalan: aniq bir talaba, bemor, avtomobil). "
         "Atribut - obyektning alohida bir xususiyati yoki ko'rsatkichi (masalan: talabaning yoshi, GPA si, fakulteti)."),
        
        ("2. Interval va nisbat atributlarining asosiy farqi nimada?",
         "Nisbat (Ratio) shkalasida haqiqiy mutlaq nol mavjud va qiymatlar nisbati ma'noga ega (masalan, 20 yosh 10 yoshdan 2 barobar katta). "
         "Interval shkalada esa mutlaq nol yo'q (masalan, 0°C harorat issiqlik yo'qligini bildirmaydi, 20°C 10°C dan 2 barobar issiq deyilmaydi, faqat 10°C ga farq qiladi)."),
        
        ("3. Simmetrik va asimmetrik binar atributlarga misol keltiring.",
         "Simmetrik binar: Har ikkala holat ham teng ahamiyatga ega (masalan: Jinsi - Erkak/Ayol). "
         "Asimmetrik binar: Bitta holat (odatda 1/Ha) ikkinchisidan ancha muhimroq (masalan: Tibbiy test natijasi - Kasal/Sog'lom, firibgarlik holati - Bor/Yo'q)."),
        
        ("4. Qaysi statistik ko'rsatkich g'ayrioddiy qiymatlarga (outlier) kamroq sezgir: o'rtacha yoki mediana? Nima uchun?",
         "Mediana outlierlarga ancha chidamli (kamroq sezgir). Chunki o'rtacha qiymat barcha sonlarni qo'shib hisoblaydi va bitta anomal katta son (masalan, 120 yosh) "
         "o'rtachani birdaniga buzadi. Mediana esa qatordagi faqat o'rtadagi elementning o'rniga qaraydi va chetki qiymatlarning hajmiga bog'liq emas."),
        
        ("5. Ma'lumotlar sifatini pasaytiruvchi asosiy omillarni sanab bering.",
         "1) Yetishmayotgan (bo'sh / Missing values) qiymatlar;\n"
         "2) Shovqinli va anomaliya (Outliers / Noise) ma'lumotlar (masalan 120 yosh);\n"
         "3) Nomuvofiq va qarama-qarshi (Inconsistent) ma'lumotlar;\n"
         "4) Takrorlanuvchi (Duplicate) yozuvlar;\n"
         "5) Formatlash va kiritishdagi insoniy xatolar.")
    ]
    for q, a in qa:
        print(f"\n{q}\n>>> {a}")

if __name__ == "__main__":
    main()
