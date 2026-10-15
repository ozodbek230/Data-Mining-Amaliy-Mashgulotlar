# -*- coding: utf-8 -*-
"""
4-amaliy mashg'ulot: Ma’lumotlarni tozalash, yetishmayotgan qiymatlarni aniqlash va qayta ishlash
Ushbu skript topshiriqdagi barcha 4 ta vazifani, hisob-kitoblarni, pandas orqali tozalashni va nazorat savollarini bajaradi.
"""

import os
import numpy as np
import pandas as pd

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
CSV_PATH = os.path.join(CURRENT_DIR, "mijozlar.csv")

def main():
    print("=" * 80)
    print("  4-AMALIY MASHG'ULOT: MA'LUMOTLARNI TOZALASH VA YETISHMAYOTGAN QIYMATLAR TAHLILI")
    print("=" * 80)

    df_raw = pd.read_csv(CSV_PATH)

    # -------------------------------------------------------------------------
    # 1-TOPSHIRIQ. MA'LUMOTLAR SIFATINI TEKSHIRISH
    # -------------------------------------------------------------------------
    print("\n" + "=" * 80)
    print("1-TOPSHIRIQ. MA'LUMOTLAR SIFATINI TEKSHIRISH VA MUAMMOLAR JADVALI")
    print("=" * 80)

    quality_problems = [
        {"Muammo turi": "Yetishmayotgan qiymat", "Qaysi atribut(lar)da?": "Yoshi, Shahar, Oylik_xarid, Xaridlar_soni", "Qaysi satr(lar)da (ID)?": "ID: 2, 4, 5, 6, 9", "Soni": 5},
        {"Muammo turi": "Takroriy yozuv", "Qaysi atribut(lar)da?": "Barcha ustunlar (Rustam mijoz)", "Qaysi satr(lar)da (ID)?": "ID: 7 (ID 3 ning dublikati)", "Soni": 1},
        {"Muammo turi": "Nomuvofiq format", "Qaysi atribut(lar)da?": "Jinsi, Shahar (kichik harflar)", "Qaysi satr(lar)da (ID)?": "ID: 3 ('erkak', 'samarqand')", "Soni": 2},
        {"Muammo turi": "G‘ayrioddiy qiymat (Outlier)", "Qaysi atribut(lar)da?": "Oylik_xarid (5000 ming so'm)", "Qaysi satr(lar)da (ID)?": "ID: 9 (Ulug‘bek)", "Soni": 1}
    ]
    df_prob = pd.DataFrame(quality_problems)
    print(df_prob.to_string(index=False))

    print("\nSavol: 7-satr 3-satrning takrori ekanini kompyuter dasturi nima uchun darhol aniqlay olmasligi mumkin?")
    print("Javob: Chunki kompyuter satrlarni harf-baharf (katta-kichik harf registrini hisobga olgan holda) taqqoslaydi.")
    print("       3-satrda: Jinsi='erkak' (kichik harfda), Shahar='samarqand' (kichik harfda).")
    print("       7-satrda: Jinsi='Erkak' (bosh harfda), Shahar='Samarqand' (bosh harfda).")
    print("       Dastur 'samarqand' == 'Samarqand' ifodasini False (teng emas) deb baholaydi va dastlab dublikat deb tanimaydi.")
    print("       Faqat registrlar standartlashtirilgandan (capitalize qilingandan) so'nggina dublikat aniqlanadi.")

    # -------------------------------------------------------------------------
    # 2-TOPSHIRIQ. YETISHMAYOTGAN QIYMATLARNI QAYTA ISHLASH USULINI TANLASH
    # -------------------------------------------------------------------------
    print("\n" + "=" * 80)
    print("2-TOPSHIRIQ. YETISHMAYOTGAN QIYMATLARNI QAYTA ISHLASH USULINI TANLASH")
    print("=" * 80)

    # 7-satrni olib tashlash (dublikat)
    df_nodup = df_raw.drop(index=6).copy()
    
    print("Avval takroriy 7-satr o'chirildi. Qolgan 9 ta qatordagi yetishmayotgan qiymatlar tahlili:\n")
    
    impute_table = [
        {"Atribut": "Yoshi (ID 2, 9)", "Tanlangan usul": "Mediana", "Hisoblash": "Mavjud yoshlar: [25, 27, 28, 30, 33, 34, 41] -> Mediana = 30", "To'ldiriladigan qiymat": "30 yosh", "Asoslash": "Sonli atribut, simmetrik taqsimotga yaqin, butun son."},
        {"Atribut": "Shahar (ID 4)", "Tanlangan usul": "Moda", "Hisoblash": "Samarqand: 4 marta, Toshkent: 2 marta, Buxoro: 2 marta", "To'ldiriladigan qiymat": "Samarqand", "Asoslash": "Kategorial (nominal) atribut bo'lgani uchun o'rtacha olib bo'lmaydi, eng ko'p uchragan shahar olinadi."},
        {"Atribut": "Oylik_xarid (ID 5)", "Tanlangan usul": "Mediana", "Hisoblash": "Mavjud xaridlar: [380, 450, 470, 490, 520, 560, 610, 5000] -> Mediana = 505", "To'ldiriladigan qiymat": "505 ming so'm", "Asoslash": "9-satrda 5000 outlier mavjud bo'lgani sababli o'rtacha qiymat xolis emas, mediana esa o'ta chidamli."},
        {"Atribut": "Xaridlar_soni (ID 6)", "Tanlangan usul": "Mediana / Moda", "Hisoblash": "Mavjud xaridlar soni: [3, 4, 5, 5, 5, 6, 6, 7] -> Mediana = 5, Moda = 5", "To'ldiriladigan qiymat": "5 marta", "Asoslash": "Diskret sonli ko'rsatkich, eng tipik mijoz xaridlari soni 5 ta."}
    ]
    df_imp = pd.DataFrame(impute_table)
    print(df_imp.to_string(index=False))

    # Oylik_xarid o'rtacha va medianasi taqqoslashi
    xarid_vals = df_nodup["Oylik_xarid"].dropna()
    mean_xarid = xarid_vals.mean()
    med_xarid = xarid_vals.median()
    print(f"\n• 'Oylik_xarid' atributi uchun ko'rsatkichlar (8 ta mavjud qiymat):")
    print(f"  - Arifmetik o'rtacha (Mean): {mean_xarid:.2f} ming so'm")
    print(f"  - Mediana (Median): {med_xarid:.2f} ming so'm")
    print(f"  - Farq: {mean_xarid - med_xarid:.2f} ming so'm!")
    print("  Katta farqning sababi: 9-satrda 5000 ming so'm qiymatli g'ayrioddiy anomaliya (outlier) mavjudligi.")
    print("  O'rtacha qiymat ushbu bitta katta songa o'ta sezgir bo'lib, sun'iy ravishda 1060 ming so'mga ko'tarilib ketgan.")
    print("  Mediana esa chetki xatolarga chidamli bo'lib, haqiqiy o'rtacha mijoz xarajatini (505 ming so'm) xolis ko'rsatadi.")

    # Bo'sh satrlarni shunchaki o'chirib tashlash haqida
    df_dropna = df_nodup.dropna()
    print(f"\n• Agar yetishmayotgan qiymatli satrlar shunchaki o'chirilsa:")
    print(f"  - Qolgan satrlar soni: {len(df_dropna)} ta (dastlabki 10 tadan atigi 5 tasi qoladi!)")
    print(f"  - Kamchiligi: Ma'lumotlarning 50% i (teng yarmi) yo'qotiladi! Natijada tanlanma reprezentativligini yo'qotadi va kichik to'plamlarda tahlil butunlay yaroqsiz holga keladi.")

    # -------------------------------------------------------------------------
    # 3-TOPSHIRIQ. EXCEL'DA MA'LUMOTLARNI TOZALASH BOSQICHLARI
    # -------------------------------------------------------------------------
    print("\n" + "=" * 80)
    print("3-TOPSHIRIQ. EXCEL VOSITASIDA BOSQICHMA-BOSQICH TOZALASH")
    print("=" * 80)
    print("1. Standartlashtirish: Jinsi va Shahar ustunlari uchun =PROPER() funksiyasi qo'llanilib, 'Erkak' va 'Samarqand' holatiga keltirildi.")
    print("2. Dublikatni o'chirish: Data -> Remove Duplicates menyusi orqali ID ustunisiz taqqoslanib, 7-satrdagi takror o'chirildi (9 ta satr qoldi).")
    print("3. Bo'sh kataklarni to'ldirish: Yoshi=30, Shahar='Samarqand', Oylik_xarid=505, Xaridlar_soni=5 qiymatlari bilan to'ldirildi.")
    print("4. Outlierni to'g'rilash: 9-satrdagi 5000 qiymati mediana (505) bilan almashtirildi.")
    print("Natija: Tozalangan jadvalda 9 ta satr qoldi va 0 ta bo'sh katak qoldi.")

    # -------------------------------------------------------------------------
    # 4-TOPSHIRIQ. PYTHON (PANDAS) YORDAMIDA TOZALASH
    # -------------------------------------------------------------------------
    print("\n" + "=" * 80)
    print("4-TOPSHIRIQ. PYTHON (PANDAS) YORDAMIDA MA'LUMOTLARNI TOZALASH")
    print("=" * 80)

    df = pd.read_csv(CSV_PATH)
    print("Boshlang'ich bo'sh qiymatlar soni:\n", df.isnull().sum())

    # 1) Formatni bir xil qilish
    for col in ["Jinsi", "Shahar"]:
        df[col] = df[col].str.strip().str.capitalize()

    # 2) Takrorlarni topish va o'chirish (ID ustunisiz)
    ustunlar = df.columns[1:]
    takrorlar_soni = df.duplicated(subset=ustunlar).sum()
    print(f"\n1-qadamdan keyin aniqlangan takrorlar soni: {takrorlar_soni} ta")
    df = df.drop_duplicates(subset=ustunlar)

    # 3) Yetishmayotgan qiymatlarni to'ldirish
    df["Yoshi"] = df["Yoshi"].fillna(df["Yoshi"].median())
    df["Shahar"] = df["Shahar"].fillna(df["Shahar"].mode()[0])
    med = df["Oylik_xarid"].median()
    df.loc[df["Oylik_xarid"] > 2000, "Oylik_xarid"] = med
    df["Oylik_xarid"] = df["Oylik_xarid"].fillna(med)
    df["Xaridlar_soni"] = df["Xaridlar_soni"].fillna(df["Xaridlar_soni"].median())

    print("\nTozalangandan keyingi bo'sh qiymatlar soni:\n", df.isnull().sum())
    print(f"\nTozalangan jadval o'lchami (df.shape): {df.shape}")
    print("\nYakuniy tozalangan jadval (Cleaned DataFrame):")
    print(df.to_string(index=False))

    print("\nQo'shimcha savollarga javoblar:")
    print("• Agar 1-qadam (formatni bir xil qilish) bajarilmasa, takrorlar soni nechta chiqadi?")
    print("  Javob: Takrorlar soni 0 chiqadi! Chunki 'samarqand' != 'Samarqand' bo'lib, kompyuter ularni turli xil qiymat deb o'ylaydi.")
    print("• Python natijasi Excel natijasi bilan mos keladimi?")
    print("  Javob: Ha, to'liq 100% mos keladi (9 ta satr, 0 ta bo'sh katak, Yoshi=30, Shahar=Samarqand, Oylik_xarid=505, Xaridlar=5).")

    # -------------------------------------------------------------------------
    # 5. NAZORAT SAVOLLARI VA JAVOBLARI
    # -------------------------------------------------------------------------
    print("\n" + "=" * 80)
    print("5. NAZORAT SAVOLLARI VA JAVOBLARI")
    print("=" * 80)
    qa = [
        ("1. Ma’lumotlarni tozalash nima uchun Data mining jarayonining muhim bosqichi hisoblanadi?",
         "Chunki real ma'lumotlar shovqinli, to'liq emas va xatoliklarga to'la bo'ladi. 'Garbage in — garbage out' (axlat kirsa — axlat chiqadi) "
         "tamoyiliga ko'ra, xato ma'lumotlarga asoslangan har qanday mashinali o'qitish modeli yoki statistik tahlil noto'g'ri xulosa beradi."),
        
        ("2. Yetishmayotgan qiymatlar paydo bo‘lishining qanday sabablarini bilasiz?",
         "1) Foydalanuvchining ma'lumotni kiritishni istamasligi (masalan, yoshi yoki daromadini yashirishi);\n"
         "2) Dasturiy yoki texnik uzilishlar (server xatosi, sensor nosozligi);\n"
         "3) Tizimlar integratsiyasidagi ma'lumotlar yo'qolishi;\n"
         "4) Ayrim savollarning hamma uchun ham tegishli emasligi."),
        
        ("3. Qachon yetishmayotgan qiymatli satrni o‘chirish maqsadga muvofiq, qachon esa to‘ldirish kerak?",
         "O'chirish maqsadga muvofiq: Agar dataset juda katta bo'lsa va yetishmayotgan satrlar ulushi 1-3% dan kam bo'lsa, yoki aynan nishon (target/label) ustuni bo'sh bo'lsa.\n"
         "To'ldirish (imputation) kerak: Agar dataset kichik bo'lsa (bizning holatdek 10 ta qator) yoki satr o'chirilsa boshqa qimmatli ustunlar yo'qotilishi xavfi bo'lsa."),
        
        ("4. Nima uchun nominal atributni o‘rtacha qiymat bilan to‘ldirib bo‘lmaydi?",
         "Chunki nominal atributlar sifat ko'rsatkichlari (masalan: Shahar, Kasb, Rangi) bo'lib, matnli toifalardan iborat. "
         "Ularni qo'shish yoki bo'lish arifmetik amallarini bajarib bo'lmaydi (masalan, Samarqand va Toshkentning 'o'rtachasi' degan tushuncha yo'q). Shuning uchun faqat eng ko'p uchraydigan qiymat — Moda ishlatiladi."),
        
        ("5. G‘ayrioddiy qiymat (outlier) har doim xato hisoblanadimi? Misol bilan tushuntiring.",
         "Yo'q, har doim xato hisoblanmaydi! U ikki xil bo'lishi mumkin:\n"
         "a) Xatolik (Noise/Error): Masalan, talabaning yoshi 120 deb kiritilishi (texnik yoki insoniy xato);\n"
         "b) Noyob haqiqiy holat (Genuine Outlier): Masalan, bank tizimida o'ta boy mijozning milliardlab tranzaksiyasi yoki kibertahdid (hujum) signali. "
         "Bunday haqiqiy anomaliyalarni o'chirib bo'lmaydi, aksincha anomaliyalarni aniqlash (Fraud detection) maqsadida chuqur o'rganiladi.")
    ]
    for q, a in qa:
        print(f"\n{q}\n>>> {a}")

if __name__ == "__main__":
    main()
