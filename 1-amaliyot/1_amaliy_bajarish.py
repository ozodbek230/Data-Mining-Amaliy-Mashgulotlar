# -*- coding: utf-8 -*-
"""
1-amaliy mashg'ulot: Data Mining - Ma'lumotlarni yig'ish va qayta ishlash
Ushbu skript hujjatdagi barcha 12 ta topshiriq uchun datasetlarni generatsiya qiladi,
CSV fayllarga saqlaydi hamda to'liq hisobot DOCX faylini va Jupyter Notebook faylini yaratadi.
"""

import os
import random
import numpy as np
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
import json

# Tasodifiy sonlar generatori uchun seed
np.random.seed(42)
random.seed(42)

DATASET_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "datasets")
os.makedirs(DATASET_DIR, exist_ok=True)

print("Datasetlar katalogi:", DATASET_DIR)

# ==============================================================================
# 1-TOPSHIRIQ: Sog'liqni saqlash (Shifoxona xarajatlari va tibbiy xizmat sifati)
# ==============================================================================
def generate_task_1(n=200):
    shifoxonalar = [f"Shifoxona_{i:03d}" for i in range(1, n + 1)]
    viloyatlar = np.random.choice(["Toshkent", "Samarqand", "Farg'ona", "Buxoro", "Andijon", "Qashqadaryo", "Xorazm"], n)
    ovqatlanish_xarajati_mln = np.round(np.random.uniform(50, 300, n), 2)
    dorilar_xarajati_mln = np.round(np.random.uniform(120, 800, n), 2)
    tibbiy_xizmat_xarajati_mln = np.round(np.random.uniform(200, 1200, n), 2)
    boshqa_xarajatlar_mln = np.round(np.random.uniform(30, 250, n), 2)
    
    jami_xarajat_mln = np.round(ovqatlanish_xarajati_mln + dorilar_xarajati_mln + tibbiy_xizmat_xarajati_mln + boshqa_xarajatlar_mln, 2)
    bemorlar_soni = np.random.randint(500, 5000, n)
    
    # Sifat ko'rsatkichi (1-100 ball)
    xizmat_sifati_bali = np.clip(np.round(50 + (jami_xarajat_mln / bemorlar_soni) * 50 + np.random.normal(0, 5, n), 1), 50, 100)
    shifokorlar_soni = np.random.randint(20, 250, n)

    df = pd.DataFrame({
        "Shifoxona_ID": shifoxonalar,
        "Hudud": viloyatlar,
        "Bemorlar_Soni": bemorlar_soni,
        "Shifokorlar_Soni": shifokorlar_soni,
        "Ovqatlanish_Xarajati_mln": ovqatlanish_xarajati_mln,
        "Dorilar_Xarajati_mln": dorilar_xarajati_mln,
        "Tibbiy_Xizmat_Xarajati_mln": tibbiy_xizmat_xarajati_mln,
        "Boshqa_Xarajatlar_mln": boshqa_xarajatlar_mln,
        "Jami_Xarajat_mln": jami_xarajat_mln,
        "Xizmat_Sifati_Bali": xizmat_sifati_bali
    })
    return df

# ==============================================================================
# 2-TOPSHIRIQ: Qurilish (Loyiha hajmi, pudratchi tajribasi, muddat va byudjet)
# ==============================================================================
def generate_task_2(n=200):
    loyihalar = [f"Loyiha_{i:03d}" for i in range(1, n + 1)]
    bino_turi = np.random.choice(["Turar joy majmuasi", "Maktab / Bog'cha", "Tibbiyot markazi", "Biznes markazi", "Sanoat korxonasi"], n)
    loyiha_maydoni_kv_m = np.random.randint(1000, 50000, n)
    pudratchi_tajribasi_yil = np.random.randint(1, 30, n)
    rejalashtirilgan_byudjet_mlrd = np.round(loyiha_maydoni_kv_m * np.random.uniform(4, 7, n) / 1000, 2)
    
    # Amaldagi byudjet (tajribaga teskari proportsional og'ish bilan)
    byudjet_oshish_foizi = np.round(np.maximum(0, 25 - pudratchi_tajribasi_yil * 0.7 + np.random.normal(0, 5, n)), 1)
    haqiqiy_xarajat_mlrd = np.round(rejalashtirilgan_byudjet_mlrd * (1 + byudjet_oshish_foizi / 100), 2)
    
    reja_muddat_oy = np.random.randint(6, 36, n)
    kechikish_oy = np.maximum(0, np.round((20 - pudratchi_tajribasi_yil * 0.6 + np.random.normal(0, 3, n)) / 4).astype(int))
    haqiqiy_muddat_oy = reja_muddat_oy + kechikish_oy
    ishchilar_soni = (loyiha_maydoni_kv_m / 200 + np.random.randint(10, 50, n)).astype(int)

    df = pd.DataFrame({
        "Loyiha_ID": loyihalar,
        "Bino_Turi": bino_turi,
        "Maydon_kv_m": loyiha_maydoni_kv_m,
        "Ishchilar_Soni": ishchilar_soni,
        "Pudratchi_Tajribasi_Yil": pudratchi_tajribasi_yil,
        "Rejalashtirilgan_Byudjet_mlrd": rejalashtirilgan_byudjet_mlrd,
        "Haqiqiy_Xarajat_mlrd": haqiqiy_xarajat_mlrd,
        "Byudjet_Ortishi_Foiz": byudjet_oshish_foizi,
        "Reja_Muddat_Oy": reja_muddat_oy,
        "Kechikish_Oy": kechikish_oy,
        "Haqiqiy_Muddat_Oy": haqiqiy_muddat_oy,
        "Muvaffaqiyatli_Tugallandi": np.where((kechikish_oy <= 1) & (byudjet_oshish_foizi <= 5), "Ha", "Yo'q")
    })
    return df

# ==============================================================================
# 3-TOPSHIRIQ: Ta'lim (Talabalar davomati, natijalar va kursda qolish xavfi)
# ==============================================================================
def generate_task_3(n=250):
    talabalar = [f"Talaba_{i:04d}" for i in range(1, n + 1)]
    kurs = np.random.choice([1, 2, 3, 4], n)
    fakultet = np.random.choice(["KIF", "DIF", "AKT", "Iqtisodiyot", "Telekommunikatsiya"], n)
    davomat_foizi = np.round(np.clip(np.random.normal(82, 12, n), 40, 100), 1)
    oraliq_baho_ortacha = np.round(np.clip(np.random.normal(70, 15, n), 45, 100), 1)
    mustaqil_ish_topshirildi_foiz = np.round(np.clip(davomat_foizi * 0.9 + np.random.normal(0, 8, n), 30, 100), 1)
    ijtimoiy_yordam_oladi = np.random.choice(["Ha", "Yo'q"], n, p=[0.25, 0.75])
    yotoqxonada_turadi = np.random.choice(["Ha", "Yo'q"], n, p=[0.45, 0.55])
    haftalik_ishlash_soati = np.random.choice([0, 10, 20, 30, 40], n, p=[0.4, 0.2, 0.2, 0.15, 0.05])
    
    # Xavf darajasi formulasi
    xavf_bali = 100 - (davomat_foizi * 0.4 + oraliq_baho_ortacha * 0.4 + mustaqil_ish_topshirildi_foiz * 0.2) + (haftalik_ishlash_soati * 0.3)
    xavf_guruhi = []
    for x in xavf_bali:
        if x > 35:
            xavf_guruhi.append("Yuqori xavf (qolish xavfi bor)")
        elif x > 20:
            xavf_guruhi.append("O'rta xavf")
        else:
            xavf_guruhi.append("Past xavf (xavfsiz)")

    df = pd.DataFrame({
        "Talaba_ID": talabalar,
        "Fakultet": fakultet,
        "Kurs": kurs,
        "Davomat_Foizi": davomat_foizi,
        "Oraliq_Baho_Ortacha": oraliq_baho_ortacha,
        "Mustaqil_Ish_Foizi": mustaqil_ish_topshirildi_foiz,
        "Ijtimoiy_Yordam": ijtimoiy_yordam_oladi,
        "Yotoqxonada_Turadi": yotoqxonada_turadi,
        "Haftalik_Ishlash_Soati": haftalik_ishlash_soati,
        "Xavf_Guruhi": xavf_guruhi
    })
    return df

# ==============================================================================
# 4-TOPSHIRIQ: Sog'liqni saqlash (Shifoxonadan keyin bemor sog'lig'i, parvarish)
# ==============================================================================
def generate_task_4(n=200):
    bemorlar = [f"Bemor_{i:04d}" for i in range(1, n + 1)]
    yosh = np.random.randint(18, 85, n)
    tashxis = np.random.choice(["Gipertoniya", "Qandli diabet", "Pnevmoniya", "Infarkt reabilitatsiya", "Oshqozon yarasi"], n)
    davolanish_muddati_kun = np.random.randint(3, 25, n)
    dorilarga_rioya_qilish_foiz = np.round(np.clip(np.random.normal(80, 15, n), 30, 100), 1)
    reja_boyicha_shifokor_koriguvi = np.random.choice(["Ha", "Yo'q"], n, p=[0.75, 0.25])
    parhezga_rioya_qilish = np.random.choice(["A'lo", "Yaxshi", "Qoniqarsiz"], n, p=[0.35, 0.45, 0.20])
    
    # Qayta kasalxonaga yotqizilish xavfi
    qayta_yotqizildi = []
    for d, k, p in zip(dorilarga_rioya_qilish_foiz, reja_boyicha_shifokor_koriguvi, parhezga_rioya_qilish):
        prob = 0.05
        if d < 60: prob += 0.35
        if k == "Yo'q": prob += 0.25
        if p == "Qoniqarsiz": prob += 0.20
        qayta_yotqizildi.append("Ha" if random.random() < prob else "Yo'q")

    df = pd.DataFrame({
        "Bemor_ID": bemorlar,
        "Yosh": yosh,
        "Asosiy_Tashxis": tashxis,
        "Statsionar_Kunlar": davolanish_muddati_kun,
        "Dorilarga_Rioya_Foizi": dorilarga_rioya_qilish_foiz,
        "Shifokor_Nazoratiga_Keldi": reja_boyicha_shifokor_koriguvi,
        "Parhezga_Rioya": parhezga_rioya_qilish,
        "Soglik_Tiklanish_Darajasi": np.random.choice(["To'liq tuzaldi", "Yaxshilandi", "O'zgarishsiz", "Yomonlashdi"], n, p=[0.45, 0.35, 0.15, 0.05]),
        "30_Kun_Ichida_Qayta_Yotqizildi": qayta_yotqizildi
    })
    return df

# ==============================================================================
# 5-TOPSHIRIQ: Ta'lim (O'qitish usullari samaradorligi va o'quvchilar faolligi)
# ==============================================================================
def generate_task_5(n=200):
    guruhlar = [f"Guruh_{i:03d}" for i in range(1, n + 1)]
    fan = np.random.choice(["Matematika", "Dasturlash", "Fizika", "Ingliz tili", "Sun'iy intellekt"], n)
    oqitish_usuli = np.random.choice(["An'anaviy ma'ruza", "Interaktiv / Gamifikatsiya", "Loyiha asosida ta'lim (PBL)", "Aralash ta'lim (Blended)"], n)
    oquvchilar_soni = np.random.randint(20, 35, n)
    darsdagi_faollik_foiz = []
    yakuniy_baho_ortacha = []
    
    for u in oqitish_usuli:
        if u == "Interaktiv / Gamifikatsiya":
            faollik = np.random.normal(88, 6)
            baho = np.random.normal(84, 5)
        elif u == "Loyiha asosida ta'lim (PBL)":
            faollik = np.random.normal(85, 7)
            baho = np.random.normal(86, 6)
        elif u == "Aralash ta'lim (Blended)":
            faollik = np.random.normal(78, 8)
            baho = np.random.normal(79, 7)
        else: # An'anaviy
            faollik = np.random.normal(65, 10)
            baho = np.random.normal(72, 8)
        darsdagi_faollik_foiz.append(np.clip(np.round(faollik, 1), 40, 100))
        yakuniy_baho_ortacha.append(np.clip(np.round(baho, 1), 50, 100))

    df = pd.DataFrame({
        "Guruh_ID": guruhlar,
        "Fan_Nomi": fan,
        "Oqitish_Usuli": oqitish_usuli,
        "Oquvchilar_Soni": oquvchilar_soni,
        "Darsdagi_Faollik_Foiz": darsdagi_faollik_foiz,
        "Yakuniy_Ortacha_Baho": yakuniy_baho_ortacha,
        "Qoniqish_Indeksi_10_Ballik": np.round(np.array(yakuniy_baho_ortacha) / 10 * 0.95 + np.random.normal(0, 0.3, n), 1)
    })
    return df

# ==============================================================================
# 6-TOPSHIRIQ: Sog'liqni saqlash (Epidemiologiya, aholi salomatligi va atrof-muhit)
# ==============================================================================
def generate_task_6(n=180):
    hududlar = [f"Tuman_{i:03d}" for i in range(1, n + 1)]
    aholi_soni_ming = np.random.randint(30, 400, n)
    havo_sifati_AQI = np.random.randint(30, 250, n) # Air Quality Index
    yashil_maydon_foizi = np.round(np.clip(50 - (havo_sifati_AQI * 0.15) + np.random.normal(0, 5, n), 5, 60), 1)
    sanoat_korxonalari_soni = np.random.randint(2, 60, n)
    toza_ichimlik_suvi_taminoti_foiz = np.round(np.clip(np.random.normal(85, 10, n), 50, 100), 1)
    
    # Nafas yo'llari kasalliklari darajasi (1000 kishiga)
    respirator_kasalliklar_1000 = np.round(15 + havo_sifati_AQI * 0.25 + sanoat_korxonalari_soni * 0.3 + np.random.normal(0, 3, n), 1)
    ortacha_umr_davomiyligi = np.round(np.clip(76 - (havo_sifati_AQI * 0.03) + (toza_ichimlik_suvi_taminoti_foiz * 0.05), 65, 80), 1)

    df = pd.DataFrame({
        "Hudud_ID": hududlar,
        "Aholi_Ming": aholi_soni_ming,
        "Havo_Sifati_AQI": havo_sifati_AQI,
        "Yashil_Maydon_Foizi": yashil_maydon_foizi,
        "Sanoat_Korxonalari_Soni": sanoat_korxonalari_soni,
        "Ichimlik_Suvi_Taminoti_Foiz": toza_ichimlik_suvi_taminoti_foiz,
        "Nafas_Kasalliklari_1000ga": respirator_kasalliklar_1000,
        "Ortacha_Umr_Davomiyligi": ortacha_umr_davomiyligi
    })
    return df

# ==============================================================================
# 7-TOPSHIRIQ: Sport (Futbol jamoalari statistikasi, gollar, natijalar)
# ==============================================================================
def generate_task_7(n=100):
    jamoalar = [f"Futbol_Klub_{i:03d}" for i in range(1, n + 1)]
    ligalar = np.random.choice(["Superliga", "Pro liga", "Birinchi liga"], n, p=[0.3, 0.4, 0.3])
    oyinlar_soni = np.full(n, 30) # Mavsumda 30 o'yin
    galabalar = np.random.randint(3, 24, n)
    duranglar = np.random.randint(2, 12, n)
    maglubiyatlar = []
    for g, d in zip(galabalar, duranglar):
        m = 30 - g - d
        if m < 0:
            d = 30 - g
            m = 0
        maglubiyatlar.append(m)
    maglubiyatlar = np.array(maglubiyatlar)
    
    ochkolar = galabalar * 3 + duranglar * 1
    kiritilgan_gollar = np.round(galabalar * 1.8 + duranglar * 0.8 + np.random.randint(0, 10, n)).astype(int)
    otkazilgan_gollar = np.round(maglubiyatlar * 1.7 + duranglar * 0.8 + np.random.randint(0, 8, n)).astype(int)
    gollar_farqi = kiritilgan_gollar - otkazilgan_gollar
    xavfli_vaziyatlar_soni = (kiritilgan_gollar * 3.5 + np.random.randint(20, 60, n)).astype(int)
    oxirgi_5_yillik_sovrinlar = np.random.choice([0, 1, 2, 3, 5, 8], n, p=[0.6, 0.2, 0.1, 0.05, 0.03, 0.02])

    df = pd.DataFrame({
        "Jamoa_Nomi": jamoalar,
        "Liga": ligalar,
        "Oyinlar_Soni": oyinlar_soni,
        "Galabalar": galabalar,
        "Duranglar": duranglar,
        "Maglubiyatlar": maglubiyatlar,
        "Ochkolar": ochkolar,
        "Urilgan_Gollar": kiritilgan_gollar,
        "Otkazilgan_Gollar": otkazilgan_gollar,
        "Gollar_Farqi": gollar_farqi,
        "Xavfli_Vaziyatlar": xavfli_vaziyatlar_soni,
        "Sovrinlar_Oxirgi_5_Yil": oxirgi_5_yillik_sovrinlar
    })
    return df

# ==============================================================================
# 8-TOPSHIRIQ: Jamoat transporti (Yo'nalishlar aniqligi, yo'lovchilar yuki)
# ==============================================================================
def generate_task_8(n=150):
    yonalishlar = [f"Avtobus_Yonalish_{i:03d}" for i in range(1, n + 1)]
    transport_turi = np.random.choice(["Avtobus", "Elektrobus", "Tramvay", "Metro liniyasi"], n, p=[0.5, 0.25, 0.1, 0.15])
    yonalish_uzunligi_km = np.round(np.random.uniform(8, 45, n), 1)
    oraliq_vaqt_reja_min = np.random.choice([5, 8, 10, 12, 15, 20], n)
    kechikish_ortacha_min = np.round(np.clip(np.random.exponential(3, n), 0.5, 25), 1)
    kunlik_yolovchilar_ming = np.round(yonalish_uzunligi_km * np.random.uniform(0.4, 1.2, n), 1)
    tigiz_vaqtdagi_yuklanish_foiz = np.round(np.clip(np.random.normal(82, 14, n), 40, 120), 1)
    infratuzilma_sifati_10 = np.random.choice([5, 6, 7, 8, 9, 10], n, p=[0.1, 0.15, 0.3, 0.25, 0.15, 0.05])
    reys_rejasining_bajarilishi_foiz = np.round(np.clip(100 - kechikish_ortacha_min * 1.5, 60, 99.5), 1)

    df = pd.DataFrame({
        "Yonalish_ID": yonalishlar,
        "Transport_Turi": transport_turi,
        "Uzunlik_km": yonalish_uzunligi_km,
        "Rejaviy_Interval_min": oraliq_vaqt_reja_min,
        "Ortacha_Kechikish_min": kechikish_ortacha_min,
        "Kunlik_Yolovchilar_Ming": kunlik_yolovchilar_ming,
        "Tigiz_Vaqtda_Yuklanish_Foiz": tigiz_vaqtdagi_yuklanish_foiz,
        "Infratuzilma_Sifati_10": infratuzilma_sifati_10,
        "Reja_Bajarilishi_Foiz": reys_rejasining_bajarilishi_foiz
    })
    return df

# ==============================================================================
# 9-TOPSHIRIQ: Demografiya (Ijtimoiy-iqtisodiy omillar va saylov tendensiyalari)
# ==============================================================================
def generate_task_9(n=150):
    okruglar = [f"Saylov_Okrugi_{i:03d}" for i in range(1, n + 1)]
    aholi_soni = np.random.randint(15000, 80000, n)
    saylovchilar_soni = (aholi_soni * np.random.uniform(0.68, 0.76, n)).astype(int)
    ortacha_yosh = np.round(np.random.uniform(28, 48, n), 1)
    ortacha_daromad_mln = np.round(np.random.uniform(2.5, 12.0, n), 2)
    oliy_malumotlilar_foiz = np.round(np.clip(np.random.normal(32, 10, n), 10, 75), 1)
    ishsizlik_darajasi_foiz = np.round(np.clip(np.random.normal(7, 2.5, n), 2.5, 16), 1)
    saylovda_qatnashish_davomati_foiz = np.round(np.clip(60 + oliy_malumotlilar_foiz * 0.2 + (ortacha_yosh - 30) * 0.3 + np.random.normal(0, 3, n), 45, 96), 1)
    etakchi_partiya_ovoz_foizi = np.round(np.clip(np.random.normal(54, 8, n), 35, 78), 1)

    df = pd.DataFrame({
        "Okrug_ID": okruglar,
        "Umumiy_Aholi": aholi_soni,
        "Saylovchilar_Soni": saylovchilar_soni,
        "Ortacha_Yosh": ortacha_yosh,
        "Ortacha_Daromad_mln": ortacha_daromad_mln,
        "Oliy_Malumotlilar_Foiz": oliy_malumotlilar_foiz,
        "Ishsizlik_Darajasi_Foiz": ishsizlik_darajasi_foiz,
        "Saylovchilar_Davomati_Foiz": saylovda_qatnashish_davomati_foiz,
        "Goliblangan_Partiya_Ovozi_Foiz": etakchi_partiya_ovoz_foizi
    })
    return df

# ==============================================================================
# 10-TOPSHIRIQ: Sport (Futbolchilar jarohat turi, o'yinchi pozitsiyasi va sharoit)
# ==============================================================================
def generate_task_10(n=220):
    futbolchilar = [f"Oyuvchi_{i:04d}" for i in range(1, n + 1)]
    yosh = np.random.randint(17, 36, n)
    pozitsiya = np.random.choice(["Darvozabon", "Himoyachi", "Yarim himoyachi", "Hujumchi"], n, p=[0.1, 0.35, 0.35, 0.2])
    jaroxat_turi = np.random.choice(["Mialgiya (mushak cho'zilishi)", "Tizza bo'g'imi (menisk/xoch)", "Boldir/to'piq sinishi/bog'lam", "Pay cho'zilishi", "Miya chayqalishi"], n)
    maydon_qoplamasi = np.random.choice(["Tabiiy maysa", "Sun'iy maysa", "Gibrid chim"], n, p=[0.55, 0.30, 0.15])
    ob_havo = np.random.choice(["Quruq / quyoshli", "Yomg'irli / sirpanchiq", "Sovuq / qorli", "Jazirama issiq"], n)
    tiklanish_muddati_hafta = []
    
    for j in jaroxat_turi:
        if "Tizza" in j:
            tiklanish_muddati_hafta.append(random.randint(8, 36))
        elif "sinishi" in j:
            tiklanish_muddati_hafta.append(random.randint(10, 24))
        elif "Pay" in j:
            tiklanish_muddati_hafta.append(random.randint(3, 10))
        elif "Mialgiya" in j:
            tiklanish_muddati_hafta.append(random.randint(1, 4))
        else:
            tiklanish_muddati_hafta.append(random.randint(2, 6))

    df = pd.DataFrame({
        "Futbolchi_ID": futbolchilar,
        "Yosh": yosh,
        "Pozitsiya": pozitsiya,
        "Jaroxat_Turi": jaroxat_turi,
        "Maydon_Qoplamasi": maydon_qoplamasi,
        "Ob_Havo_Sharoiti": ob_havo,
        "Tiklanish_Muddati_Hafta": tiklanish_muddati_hafta,
        "Jarrohlik_Talab_Qilindi": ["Ha" if t > 12 else "Yo'q" for t in tiklanish_muddati_hafta]
    })
    return df

# ==============================================================================
# 11-TOPSHIRIQ: Jamoat transporti (Tozalik, qulaylik, yo'lovchilar qoniqishi)
# ==============================================================================
def generate_task_11(n=250):
    sorov_id = [f"Respondent_{i:04d}" for i in range(1, n + 1)]
    transport_turi = np.random.choice(["Avtobus", "Metro", "Taksi / Marshrutka"], n, p=[0.5, 0.35, 0.15])
    tozalik_bali_5 = np.random.choice([1, 2, 3, 4, 5], n, p=[0.05, 0.15, 0.35, 0.30, 0.15])
    qulaylik_konditsioner_5 = np.random.choice([1, 2, 3, 4, 5], n, p=[0.1, 0.2, 0.3, 0.25, 0.15])
    oz_vaqtida_kelish_5 = np.random.choice([1, 2, 3, 4, 5], n, p=[0.08, 0.17, 0.30, 0.30, 0.15])
    tolov_qulayligi_5 = np.random.choice([3, 4, 5], n, p=[0.1, 0.4, 0.5]) # ATTO, kartalar
    
    umumiy_qoniqish_foiz = np.round(((tozalik_bali_5 + qulaylik_konditsioner_5 + oz_vaqtida_kelish_5 + tolov_qulayligi_5) / 20) * 100 + np.random.normal(0, 3, n), 1)
    umumiy_qoniqish_foiz = np.clip(umumiy_qoniqish_foiz, 20, 100)

    df = pd.DataFrame({
        "Sorov_ID": sorov_id,
        "Transport_Turi": transport_turi,
        "Tozalik_Bali_5": tozalik_bali_5,
        "Qulaylik_Konditsioner_5": qulaylik_konditsioner_5,
        "Oz_Vaqtida_Kelish_5": oz_vaqtida_kelish_5,
        "Tolov_Qulayligi_5": tolov_qulayligi_5,
        "Umumiy_Qoniqish_Foiz": umumiy_qoniqish_foiz,
        "Tavsiya_Qiladi": ["Ha" if q >= 70 else "Yo'q" for q in umumiy_qoniqish_foiz]
    })
    return df

# ==============================================================================
# 12-TOPSHIRIQ: Demografiya (Aholi o'sishi, migratsiya, urbanizatsiya)
# ==============================================================================
def generate_task_12(n=160):
    hududlar = [f"Hudud_{i:03d}" for i in range(1, n + 1)]
    viloyat = np.random.choice(["Toshkent vil.", "Samarqand", "Farg'ona", "Andijon", "Qashqadaryo", "Surxondaryo", "Buxoro", "Xorazm", "Qoraqalpog'iston"], n)
    aholi_soni_ming = np.random.randint(50, 600, n)
    aholi_yillik_osish_surati_foiz = np.round(np.random.normal(1.8, 0.4, n), 2)
    urbanizatsiya_darajasi_foiz = np.round(np.clip(np.random.normal(48, 14, n), 15, 88), 1)
    sof_migratsiya_salbiy_ijobiy_ming = np.round(np.random.normal(-0.5, 2.5, n), 2)
    aholi_jon_boshiga_YaHM_mln_som = np.round(np.clip(10 + (urbanizatsiya_darajasi_foiz * 0.4) + np.random.normal(0, 4, n), 6, 60), 2)
    sanoat_ulushi_foiz = np.round(np.clip(urbanizatsiya_darajasi_foiz * 0.6 + np.random.normal(0, 5, n), 10, 65), 1)

    df = pd.DataFrame({
        "Hudud_ID": hududlar,
        "Viloyat": viloyat,
        "Aholi_Ming": aholi_soni_ming,
        "Yillik_Osish_Foiz": aholi_yillik_osish_surati_foiz,
        "Urbanizatsiya_Darajasi_Foiz": urbanizatsiya_darajasi_foiz,
        "Sof_Migratsiya_Ming": sof_migratsiya_salbiy_ijobiy_ming,
        "Jon_Boshiga_Daromad_mln": aholi_jon_boshiga_YaHM_mln_som,
        "Iqtisodiyotda_Sanoat_Ulushi_Foiz": sanoat_ulushi_foiz
    })
    return df

# Barcha vazifalar xaritasi
TASKS = [
    {
        "id": 1,
        "nom": "Sog'liqni saqlash",
        "tavsif": "Ko'rsatilayotgan tibbiy xizmat sifatini aniqlash uchun shifoxonalarning sarflanadigan xarajat miqdorini ovqatlanish, dorilar olish, tibbiy xizmat, boshqa xizmatlar haqidagi ma’lumotlar to’plangan jadvali.",
        "fayl": "topshiriq_01_sogliqni_saqlash_xarajatlar.csv",
        "func": generate_task_1
    },
    {
        "id": 2,
        "nom": "Qurilish",
        "tavsif": "Loyiha hajmi va pudratchi tajribasi kabi omillarni hisobga olgan holda, qurilish loyihalarini belgilangan muddat va byudjet doirasida yakunlash ma’lumotlari.",
        "fayl": "topshiriq_02_qurilish_loyihalari.csv",
        "func": generate_task_2
    },
    {
        "id": 3,
        "nom": "Ta’lim",
        "tavsif": "O'quv natijalari, davomat va ijtimoiy-iqtisodiy omillarga asoslanib, kursdan-kursga qolish xavfi ostida bo’lgan talabalar ma’lumotlari.",
        "fayl": "topshiriq_03_talim_talabalar_xavfi.csv",
        "func": generate_task_3
    },
    {
        "id": 4,
        "nom": "Sog’liqni saqlash",
        "tavsif": "Shifoxonadan chiqqandan keyin bemorning sog'lig'i, davolanishga rioya qilish va keyingi parvarishlar tahlili haqidagi ma’lumotlar.",
        "fayl": "topshiriq_04_bemor_reabilitatsiyasi.csv",
        "func": generate_task_4
    },
    {
        "id": 5,
        "nom": "Ta’lim",
        "tavsif": "Ta’lim texnikasi, o‘quvchilarning faolligi va ta’lim natijalari to‘g‘risidagi ma’lumotlardan foydalangan holda o‘quvchilar faoliyatini yaxshilashda turli o‘qitish usullarining samaradorligi.",
        "fayl": "topshiriq_05_oqitish_usullari_samaradorligi.csv",
        "func": generate_task_5
    },
    {
        "id": 6,
        "nom": "Sog’liqni saqlash",
        "tavsif": "Epidemiologik ma'lumotlar, aholi salomatligi ko'rsatkichlari va atrof-muhit omillari asosida aholi salomatligi haqidagi ma’lumotlar.",
        "fayl": "topshiriq_06_epidemiologiya_atrof_muhit.csv",
        "func": generate_task_6
    },
    {
        "id": 7,
        "nom": "Sport",
        "tavsif": "Futbol jamoasining g’alabalar soni, duranglar soni, mag’lubiyatlar soni, oxirgi besh yillikda erishgan sovrinlar soni, kiritgan va o’tkazib yuborgan gollar statistikasi.",
        "fayl": "topshiriq_07_futbol_klublar_statistikasi.csv",
        "func": generate_task_7
    },
    {
        "id": 8,
        "nom": "Jamoat transporti",
        "tavsif": "Umumiy tizim samaradorligini oshirish uchun jamoat transporti yo'nalishlarini aniqlik, yo'lovchilar yuki va infratuzilma sifati kabi ma’lumotlar.",
        "fayl": "topshiriq_08_jamoat_transporti_samaradorligi.csv",
        "func": generate_task_8
    },
    {
        "id": 9,
        "nom": "Demografiya",
        "tavsif": "Demografik ma'lumotlar, ijtimoiy-iqtisodiy omillar va ovoz berishning oldingi tendensiyalari haqidagi ma’lumotlar.",
        "fayl": "topshiriq_09_demografiya_saylov_tendensiyasi.csv",
        "func": generate_task_9
    },
    {
        "id": 10,
        "nom": "Sport",
        "tavsif": "Futbol jamoasida o’yinchilarning jarohat turi, o'yinchi pozitsiyasi va atrof-muhit sharoitlariga qarab sport jarohatlari haqidagi ma’lumotlar.",
        "fayl": "topshiriq_10_sportchilar_jaroxatlari.csv",
        "func": generate_task_10
    },
    {
        "id": 11,
        "nom": "Jamoat transporti",
        "tavsif": "Tozalik, qulaylik va o'z vaqtida ishlash kabi omillarga asoslanib, jamoat transporti xizmatlaridan yo'lovchilarning qoniqishi ma’lumotlari.",
        "fayl": "topshiriq_11_transport_yolovchilar_qoniqishi.csv",
        "func": generate_task_11
    },
    {
        "id": 12,
        "nom": "Demografiya",
        "tavsif": "Aholi o'sishi, migratsiya shakllari va iqtisodiy ko'rsatkichlarni tahlil qilish orqali turli demografik guruhlarda urbanizatsiya tendensiyalari haqida ma’lumotlar.",
        "fayl": "topshiriq_12_urbanizatsiya_tendensiyalari.csv",
        "func": generate_task_12
    }
]

# 1. Hamma datasetlarni hosil qilish va CSV larga saqlash
generated_dfs = {}
for t in TASKS:
    df = t["func"]()
    filepath = os.path.join(DATASET_DIR, t["fayl"])
    df.to_csv(filepath, index=False, encoding='utf-8-sig')
    generated_dfs[t["id"]] = (t, df, filepath)
    print(f"[OK] {t['id']}-topshiriq saqlandi: {t['fayl']} (Qatorlar soni: {len(df)}, Ustunlar: {len(df.columns)})")

# ==============================================================================
# DOCX HISOBOTINI YARATISH
# ==============================================================================
def create_docx_report():
    doc = Document()
    
    # Sahifa chetki masofalari (Margins)
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.9)
        section.right_margin = Inches(0.8)
        
    # Sarlavha
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = title.add_run("MA'LUMOTLARNI INTELLEKTUAL TAHLILI (DATA MINING)\n1-AMALIY MASHG'ULOT HISOBOTI")
    run_title.bold = True
    run_title.font.size = Pt(16)
    run_title.font.name = "Arial"
    run_title.font.color.rgb = RGBColor(0, 51, 102)

    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_sub = subtitle.add_run("Mavzu: Python yordamida ma'lumotlar to'plamini (dataset) shakllantirish, CSV formatida saqlash va dastlabki tahlil\n(Barcha 12 ta topshiriq variantlari to'liq yechimi bilan)")
    run_sub.italic = True
    run_sub.font.size = Pt(11)
    run_sub.font.name = "Arial"
    run_sub.font.color.rgb = RGBColor(100, 100, 100)

    doc.add_paragraph()

    # 1. Ishning maqsadi
    h1 = doc.add_heading("1. Ishning maqsadi va vazifalari", level=1)
    p_maqsad = doc.add_paragraph(
        "Data Mining (ma'lumotlarni qazib olish) sohasida ma'lumotlarni yig'ish, generatsiya qilish va tayyorlash jarayonlarini o'rganish; "
        "Python dasturlash tilining pandas va numpy kutubxonalari yordamida turli sohalar (sog'liqni saqlash, qurilish, ta'lim, sport, jamoat transporti va demografiya) "
        "uchun ko'p o'lchovli ma'lumotlar to'plamlarini (DataFrames) avtomatlashtirilgan tarzda yaratish, ularni .csv fayl ko'rinishida saqlash va tahlil qilish ko'nikmalarini shakllantirish."
    )
    p_maqsad.paragraph_format.line_spacing = 1.15

    # 2. Nazariy asoslar: 11 ta bosqich
    doc.add_heading("2. Data Miningda ma'lumotlarni yig'ish va tayyorlashning 11 ta bosqichi", level=1)
    stages = [
        ("1. Muammo ta'rifi", "Loyiha maqsadlarini aniqlash va qaysi masalani yechish kerakligini belgilash."),
        ("2. Ma'lumotlarni o'rganish", "Soha xususiyatlarini tushunish, potentsial manbalarni tadqiq qilish."),
        ("3. Ma'lumotlarni tanlash", "Tahlil uchun zaruriy va sifatli ma'lumot manbalarini ajratib olish."),
        ("4. Ma'lumotlarni tozalash", "Bo'sh (NaN) qiymatlar, xatoliklar va nomuvofiqliklarni to'g'rilash yoki olib tashlash."),
        ("5. Oldindan ishlov berish", "Raqamli ma'lumotlarni masshtablash/normallashtirish va matnlarni kodlash."),
        ("6. Ma'lumotlar integratsiyasi", "Turli manbalardan olingan jadvallarni yagona datasetga birlashtirish."),
        ("7. Ma'lumotlarni qisqartirish", "O'lchamlarni qisqartirish (Feature Selection) va muhim ustunlarni tanlash."),
        ("8. Tanlanma olish (Sampling)", "Katta hajmli datasetlardan reprezentativ kichik tanlanma ajratib olish."),
        ("9. Ma'lumotlarni o'zgartirish", "Algoritmlarga moslashtirish, yangi ko'rsatkichlar (features) yaratish."),
        ("10. Ma'lumotlarni bo'lish", "O'qitish (train) va sinov (test) to'plamlariga ajratish."),
        ("11. Ma'lumotlarni saqlash", "Tayyor tozalangan ma'lumotlarni CSV/SQL/Parquet formatlarida saqlash.")
    ]
    for s_nomi, s_sharh in stages:
        p = doc.add_paragraph()
        run_b = p.add_run(f"• {s_nomi}: ")
        run_b.bold = True
        p.add_run(s_sharh)

    doc.add_page_break()

    # 3. Topshiriqlar yechimi
    doc.add_heading("3. Amaliy topshiriqlar va ularning dasturiy yechimlari", level=1)

    for item in TASKS:
        t_id = item["id"]
        t_nom = item["nom"]
        t_tavsif = item["tavsif"]
        t_fayl = item["fayl"]
        _, df, _ = generated_dfs[t_id]

        doc.add_heading(f"Topshiriq #{t_id}: {t_nom}", level=2)
        
        # Tavsif
        p_desc = doc.add_paragraph()
        r_lbl = p_desc.add_run("Topshiriq sharti: ")
        r_lbl.bold = True
        p_desc.add_run(t_tavsif)

        p_file = doc.add_paragraph()
        p_file.add_run("Hosil qilingan fayl: ").bold = True
        p_file.add_run(f"datasets/{t_fayl} (Jami {len(df)} ta satr, {len(df.columns)} ta ustun)")

        # Jadval namunasi (head 5)
        doc.add_paragraph("Shakllantirilgan ma'lumotlar jadvalining dastlabki 5 ta qatori (namuna):").italic = True
        
        sample_df = df.head(5)
        cols = list(sample_df.columns)
        table = doc.add_table(rows=len(sample_df) + 1, cols=len(cols))
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = True

        # Header
        hdr_cells = table.rows[0].cells
        for idx, col_name in enumerate(cols):
            hdr_cells[idx].text = str(col_name)
            for p in hdr_cells[idx].paragraphs:
                for run in p.runs:
                    run.bold = True
                    run.font.size = Pt(8.5)
            # Orqa fon rangi header uchun
            shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="003366"/>')
            hdr_cells[idx]._tc.get_or_add_tcPr().append(shading)
            for p in hdr_cells[idx].paragraphs:
                for run in p.runs:
                    run.font.color.rgb = RGBColor(255, 255, 255)

        # Qatorlar
        for r_idx, row in sample_df.iterrows():
            row_cells = table.rows[r_idx + 1].cells
            for c_idx, col_name in enumerate(cols):
                val = str(row[col_name])
                row_cells[c_idx].text = val
                for p in row_cells[c_idx].paragraphs:
                    for run in p.runs:
                        run.font.size = Pt(8)

        doc.add_paragraph() # Bo'sh joy

    # 4. Xulosa
    doc.add_heading("4. Xulosa", level=1)
    doc.add_paragraph(
        "Mazkur 1-amaliy mashg'ulot davomida Data Mining fanining eng muhim fundamenti hisoblangan "
        "ma'lumotlarni yig'ish, ularga dastlabki ishlov berish va kompyuter xotirasida saqlash bosqichlari to'liq o'zlashtirildi. "
        "Python dasturlash tilining zamonaviy 'pandas' va 'numpy' kutubxonalari yordamida ijtimoiy-iqtisodiy, tibbiyot, sport, ta'lim, qurilish va transport "
        "sohalariga oid 12 ta real vaziyat modeli yaratildi. Har bir model asosida ko'p parametrli ma'lumotlar to'plamlari shakllantirilib, "
        "ular keyingi mashg'ulotlarda mashinali o'rganish (Machine Learning) modellariga kiritish uchun standart .csv formatiga eksport qilindi."
    )

    doc_out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "1-amaliy_mashg'ulot_hisoboti.docx")
    doc.save(doc_out)
    print("[OK] Hisobot fayli saqlandi:", doc_out)

create_docx_report()

# ==============================================================================
# JUPYTER NOTEBOOK YARATISH
# ==============================================================================
def create_jupyter_notebook():
    notebook = {
        "cells": [
            {
                "cell_type": "markdown",
                "metadata": {},
                "source": [
                    "# 1-Amaliy Mashg'ulot: Data Mining - Ma'lumotlarni yig'ish va tahlil qilish\n",
                    "Ushbu daftarda 1-amaliy mashg'ulotdagi barcha 12 ta topshiriq uchun datasetlarni yaratish, CSV faylga saqlash va tahlil qilish kodlari jamlangan."
                ]
            },
            {
                "cell_type": "code",
                "execution_count": None,
                "metadata": {},
                "outputs": [],
                "source": [
                    "import pandas as pd\n",
                    "import numpy as np\n",
                    "import os\n",
                    "\n",
                    "os.makedirs('datasets', exist_ok=True)\n",
                    "print('Kutubxonalar muvaffaqiyatli yuklandi!')"
                ]
            }
        ],
        "metadata": {
            "kernelspec": {
                "display_name": "Python 3",
                "language": "python",
                "name": "python3"
            },
            "language_info": {
                "name": "python",
                "version": "3.12.0"
            }
        },
        "nbformat": 4,
        "nbformat_minor": 2
    }

    # Har bir vazifani notebook kataklariga qo'shish
    task_code_snippets = [
        ("""# 1-Topshiriq: Sog'liqni saqlash (Shifoxona xarajatlari va sifat)
np.random.seed(42)
n = 200
df1 = pd.DataFrame({
    'Shifoxona_ID': [f'Shifoxona_{i:03d}' for i in range(1, n+1)],
    'Hudud': np.random.choice(['Toshkent', 'Samarqand', 'Fargona', 'Buxoro'], n),
    'Bemorlar_Soni': np.random.randint(500, 5000, n),
    'Ovqatlanish_Xarajati_mln': np.round(np.random.uniform(50, 300, n), 2),
    'Dorilar_Xarajati_mln': np.round(np.random.uniform(120, 800, n), 2),
    'Tibbiy_Xizmat_Xarajati_mln': np.round(np.random.uniform(200, 1200, n), 2),
    'Boshqa_Xarajatlar_mln': np.round(np.random.uniform(30, 250, n), 2)
})
df1['Jami_Xarajat_mln'] = df1.iloc[:, 3:7].sum(axis=1)
df1['Xizmat_Sifati_Bali'] = np.clip(np.round(60 + (df1['Jami_Xarajat_mln']/df1['Bemorlar_Soni'])*40, 1), 50, 100)
df1.to_csv('datasets/topshiriq_01_sogliqni_saqlash_xarajatlar.csv', index=False)
df1.head()"""),

        ("""# 2-Topshiriq: Qurilish loyihalari
n = 200
df2 = pd.DataFrame({
    'Loyiha_ID': [f'Loyiha_{i:03d}' for i in range(1, n+1)],
    'Bino_Turi': np.random.choice(['Turar joy', 'Maktab', 'Biznes markaz'], n),
    'Maydon_kv_m': np.random.randint(1000, 50000, n),
    'Pudratchi_Tajribasi_Yil': np.random.randint(1, 30, n),
    'Rejalashtirilgan_Byudjet_mlrd': np.round(np.random.uniform(5, 100, n), 2),
    'Haqiqiy_Xarajat_mlrd': np.round(np.random.uniform(5, 110, n), 2),
    'Reja_Muddat_Oy': np.random.randint(6, 36, n),
    'Kechikish_Oy': np.random.randint(0, 12, n)
})
df2.to_csv('datasets/topshiriq_02_qurilish_loyihalari.csv', index=False)
df2.head()"""),

        ("""# 3-Topshiriq: Ta'lim (Talabalar davomati va o'zlashtirishi)
n = 250
df3 = pd.DataFrame({
    'Talaba_ID': [f'Talaba_{i:04d}' for i in range(1, n+1)],
    'Kurs': np.random.choice([1, 2, 3, 4], n),
    'Davomat_Foizi': np.round(np.random.uniform(50, 100, n), 1),
    'Oraliq_Baho': np.round(np.random.uniform(40, 100, n), 1),
    'Mustaqil_Ish_Foizi': np.round(np.random.uniform(40, 100, n), 1),
    'Ijtimoiy_Yordam': np.random.choice(['Ha', 'Yoq'], n)
})
df3['Xavf_Guruhi'] = np.where((df3['Davomat_Foizi'] < 70) | (df3['Oraliq_Baho'] < 60), 'Yuqori xavf', 'Xavfsiz')
df3.to_csv('datasets/topshiriq_03_talim_talabalar_xavfi.csv', index=False)
df3.head()"""),

        ("""# 4-Topshiriq: Sog'liqni saqlash (Bemor parvarishi)
n = 200
df4 = pd.DataFrame({
    'Bemor_ID': [f'Bemor_{i:04d}' for i in range(1, n+1)],
    'Yosh': np.random.randint(18, 85, n),
    'Tashxis': np.random.choice(['Gipertoniya', 'Diabet', 'Pnevmoniya'], n),
    'Dorilarga_Rioya_Foizi': np.round(np.random.uniform(40, 100, n), 1),
    'Shifokor_Nazoratiga_Keldi': np.random.choice(['Ha', 'Yoq'], n),
    'Qayta_Yotqizildi': np.random.choice(['Ha', 'Yoq'], n, p=[0.2, 0.8])
})
df4.to_csv('datasets/topshiriq_04_bemor_reabilitatsiyasi.csv', index=False)
df4.head()"""),

        ("""# 5-Topshiriq: Ta'lim (O'qitish usullari samaradorligi)
n = 200
df5 = pd.DataFrame({
    'Guruh_ID': [f'Guruh_{i:03d}' for i in range(1, n+1)],
    'Fan': np.random.choice(['Matematika', 'Dasturlash', 'Fizika'], n),
    'Oqitish_Usuli': np.random.choice(['Ananaviy', 'Interaktiv', 'Loyiha asosida'], n),
    'Darsdagi_Faollik_Foiz': np.round(np.random.uniform(50, 100, n), 1),
    'Yakuniy_Ortacha_Baho': np.round(np.random.uniform(55, 98, n), 1)
})
df5.to_csv('datasets/topshiriq_05_oqitish_usullari_samaradorligi.csv', index=False)
df5.head()"""),

        ("""# 6-Topshiriq: Epidemiologiya va atrof-muhit
n = 180
df6 = pd.DataFrame({
    'Tuman_ID': [f'Tuman_{i:03d}' for i in range(1, n+1)],
    'Havo_AQI': np.random.randint(40, 250, n),
    'Yashil_Maydon_Foiz': np.round(np.random.uniform(5, 50, n), 1),
    'Sanoat_Korxonalari_Soni': np.random.randint(2, 50, n),
    'Respirator_Kasalliklar_1000ga': np.round(np.random.uniform(15, 80, n), 1)
})
df6.to_csv('datasets/topshiriq_06_epidemiologiya_atrof_muhit.csv', index=False)
df6.head()"""),

        ("""# 7-Topshiriq: Sport (Futbol klublari)
n = 100
df7 = pd.DataFrame({
    'Klub_Nomi': [f'Klub_{i:03d}' for i in range(1, n+1)],
    'Galabalar': np.random.randint(5, 24, n),
    'Duranglar': np.random.randint(3, 12, n),
    'Maglubiyatlar': np.random.randint(2, 18, n),
    'Urilgan_Gollar': np.random.randint(20, 85, n),
    'Otkazilgan_Gollar': np.random.randint(15, 60, n),
    'Oxirgi_5_Yil_Sovrinlar': np.random.choice([0, 1, 2, 4], n, p=[0.7, 0.15, 0.1, 0.05])
})
df7['Ochkolar'] = df7['Galabalar']*3 + df7['Duranglar']
df7.to_csv('datasets/topshiriq_07_futbol_klublar_statistikasi.csv', index=False)
df7.head()"""),

        ("""# 8-Topshiriq: Jamoat transporti samaradorligi
n = 150
df8 = pd.DataFrame({
    'Yonalish_ID': [f'Yonalish_{i:03d}' for i in range(1, n+1)],
    'Transport': np.random.choice(['Avtobus', 'Elektrobus', 'Metro'], n),
    'Uzunlik_km': np.round(np.random.uniform(10, 40, n), 1),
    'Ortacha_Kechikish_min': np.round(np.random.uniform(1, 20, n), 1),
    'Kunlik_Yolovchilar_Ming': np.round(np.random.uniform(5, 30, n), 1),
    'Infratuzilma_Sifati_10': np.random.randint(5, 11, n)
})
df8.to_csv('datasets/topshiriq_08_jamoat_transporti_samaradorligi.csv', index=False)
df8.head()"""),

        ("""# 9-Topshiriq: Demografiya va saylov tendensiyasi
n = 150
df9 = pd.DataFrame({
    'Okrug_ID': [f'Okrug_{i:03d}' for i in range(1, n+1)],
    'Ortacha_Yosh': np.round(np.random.uniform(25, 50, n), 1),
    'Ortacha_Daromad_mln': np.round(np.random.uniform(2, 12, n), 2),
    'Oliy_Malumot_Foiz': np.round(np.random.uniform(10, 70, n), 1),
    'Ishsizlik_Foiz': np.round(np.random.uniform(3, 15, n), 1),
    'Davomat_Foiz': np.round(np.random.uniform(50, 95, n), 1)
})
df9.to_csv('datasets/topshiriq_09_demografiya_saylov_tendensiyasi.csv', index=False)
df9.head()"""),

        ("""# 10-Topshiriq: Sport jarohatlari
n = 200
df10 = pd.DataFrame({
    'Oyuvchi_ID': [f'Oyuvchi_{i:04d}' for i in range(1, n+1)],
    'Pozitsiya': np.random.choice(['Himoyachi', 'Hujumchi', 'Yarim himoyachi', 'Darvozabon'], n),
    'Jaroxat_Turi': np.random.choice(['Mushak', 'Tizza', 'Boldir', 'Pay'], n),
    'Maydon_Qoplamasi': np.random.choice(['Tabiiy chim', 'Suniy chim'], n),
    'Tiklanish_Hafta': np.random.randint(2, 28, n)
})
df10.to_csv('datasets/topshiriq_10_sportchilar_jaroxatlari.csv', index=False)
df10.head()"""),

        ("""# 11-Topshiriq: Jamoat transportidan qoniqish
n = 250
df11 = pd.DataFrame({
    'Respondent_ID': [f'Respondent_{i:04d}' for i in range(1, n+1)],
    'Transport': np.random.choice(['Avtobus', 'Metro'], n),
    'Tozalik_5': np.random.randint(1, 6, n),
    'Qulaylik_5': np.random.randint(1, 6, n),
    'Oz_Vaqtida_5': np.random.randint(1, 6, n)
})
df11['Umumiy_Qoniqish_Foiz'] = np.round((df11.iloc[:, 2:5].sum(axis=1) / 15) * 100, 1)
df11.to_csv('datasets/topshiriq_11_transport_yolovchilar_qoniqishi.csv', index=False)
df11.head()"""),

        ("""# 12-Topshiriq: Urbanizatsiya tendensiyalari
n = 160
df12 = pd.DataFrame({
    'Hudud_ID': [f'Hudud_{i:03d}' for i in range(1, n+1)],
    'Viloyat': np.random.choice(['Toshkent', 'Samarqand', 'Fargona', 'Buxoro'], n),
    'Yillik_Osish_Foiz': np.round(np.random.uniform(1.0, 3.0, n), 2),
    'Urbanizatsiya_Darajasi_Foiz': np.round(np.random.uniform(20, 85, n), 1),
    'Sof_Migratsiya_Ming': np.round(np.random.uniform(-3, 3, n), 2),
    'Jon_Boshiga_Daromad_mln': np.round(np.random.uniform(8, 45, n), 2)
})
df12.to_csv('datasets/topshiriq_12_urbanizatsiya_tendensiyalari.csv', index=False)
df12.head()""")
    ]

    for snippet in task_code_snippets:
        notebook["cells"].append({
            "cell_type": "code",
            "execution_count": None,
            "metadata": {},
            "outputs": [],
            "source": [line + "\n" for line in snippet.split("\n")]
        })

    nb_out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "1-amaliy_mashgulot.ipynb")
    with open(nb_out, "w", encoding="utf-8") as f:
        json.dump(notebook, f, indent=2, ensure_ascii=False)
    print("[OK] Jupyter Notebook saqlandi:", nb_out)

create_jupyter_notebook()
print("Hammasi muvaffaqiyatli yakunlandi!")
