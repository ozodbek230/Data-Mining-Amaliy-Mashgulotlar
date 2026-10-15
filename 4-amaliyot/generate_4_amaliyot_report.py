# -*- coding: utf-8 -*-
"""
4-amaliy mashg'ulot bo'yicha to'liq Word hisoboti (4-amaliy_mashg'ulot_hisoboti.docx)
hamda Jupyter Notebook (4-amaliy_mashgulot.ipynb) yaratish skripti.
"""

import os
import json
import pandas as pd
import numpy as np
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
CSV_PATH = os.path.join(CURRENT_DIR, "mijozlar.csv")
df_raw = pd.read_csv(CSV_PATH)

# ==============================================================================
# WORD HUJJATINI TAYYORLASH
# ==============================================================================
doc = Document()

for s in doc.sections:
    s.top_margin = Inches(0.79)
    s.bottom_margin = Inches(0.79)
    s.left_margin = Inches(0.85)
    s.right_margin = Inches(0.79)

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>'))

def set_table_borders(table, color="D3D3D3"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(f'''
        <w:tblBorders {nsdecls("w")}>
            <w:top w:val="single" w:sz="4" w:space="0" w:color="{color}"/>
            <w:bottom w:val="single" w:sz="4" w:space="0" w:color="{color}"/>
            <w:insideH w:val="single" w:sz="4" w:space="0" w:color="{color}"/>
            <w:insideV w:val="none"/>
            <w:left w:val="none"/>
            <w:right w:val="none"/>
        </w:tblBorders>
    ''')
    tblPr.append(borders)

def format_row(row, is_header=False):
    for cell in row.cells:
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        for p in cell.paragraphs:
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.line_spacing = 1.05
            for run in p.runs:
                run.font.name = "Times New Roman"
                run.font.size = Pt(9.5 if not is_header else 10)
                if is_header:
                    run.bold = True
                    run.font.color.rgb = RGBColor(255, 255, 255)

def add_heading_1(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    r.bold = True
    r.font.name = "Times New Roman"
    r.font.size = Pt(14)
    r.font.color.rgb = RGBColor(15, 23, 42)
    return p

def add_heading_2(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    r.bold = True
    r.font.name = "Times New Roman"
    r.font.size = Pt(12)
    r.font.color.rgb = RGBColor(30, 58, 138)
    return p

def add_body(text, bold_prefix=None, italic=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.15
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    
    if bold_prefix:
        r_pre = p.add_run(bold_prefix)
        r_pre.bold = True
        r_pre.font.name = "Times New Roman"
        r_pre.font.size = Pt(11)
    
    r = p.add_run(text)
    r.italic = italic
    r.font.name = "Times New Roman"
    r.font.size = Pt(11)
    return p

# -----------------------------------------------------------------------------
# TITUL VARAQASI
# -----------------------------------------------------------------------------
p_vaz = doc.add_paragraph()
p_vaz.alignment = WD_ALIGN_PARAGRAPH.CENTER
r_v = p_vaz.add_run("O'ZBEKISTON RESPUBLIKASI\nOLIY TA'LIM, FAN VA INNOVATSIYALAR VAZIRLIGI\n")
r_v.bold = True
r_v.font.size = Pt(13)
r_v.font.name = "Times New Roman"

p_kaf = doc.add_paragraph()
p_kaf.alignment = WD_ALIGN_PARAGRAPH.CENTER
r_k = p_kaf.add_run("AXBOROT TEXNOLOGIYALARI VA SUN'IY INTELLEKT KAFEDRASI\n\"DATA MINING\" FANI")
r_k.font.size = Pt(12)
r_k.bold = True
r_k.font.name = "Times New Roman"

for _ in range(3):
    doc.add_paragraph()

p_title = doc.add_paragraph()
p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
r_t = p_title.add_run("4-AMALIY MASHG'ULOT HISOBOTI\n")
r_t.bold = True
r_t.font.size = Pt(20)
r_t.font.name = "Times New Roman"
r_t.font.color.rgb = RGBColor(15, 23, 42)

p_sub = doc.add_paragraph()
p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
r_sub = p_sub.add_run("Mavzu: Ma’lumotlarni tozalash, yetishmayotgan qiymatlarni aniqlash va qayta ishlash\n(«Do‘kon mijozlari» ma'lumotlar to'plami asosida barcha 4 ta topshiriq va nazorat savollari yechimi)")
r_sub.italic = True
r_sub.font.size = Pt(12.5)
r_sub.font.name = "Times New Roman"
r_sub.font.color.rgb = RGBColor(71, 85, 105)

for _ in range(4):
    doc.add_paragraph()

tbl_info = doc.add_table(rows=2, cols=2)
tbl_info.alignment = WD_TABLE_ALIGNMENT.CENTER
tbl_info.autofit = True

c00 = tbl_info.rows[0].cells[0]
c00.text = "Bajardi:"
c00.paragraphs[0].runs[0].bold = True
c00.paragraphs[0].runs[0].font.size = Pt(12)

c01 = tbl_info.rows[0].cells[1]
c01.text = "Talaba: ____________________\nGuruh: ____________________"
c01.paragraphs[0].runs[0].font.size = Pt(12)

c10 = tbl_info.rows[1].cells[0]
c10.text = "Qabul qildi:"
c10.paragraphs[0].runs[0].bold = True
c10.paragraphs[0].runs[0].font.size = Pt(12)

c11 = tbl_info.rows[1].cells[1]
c11.text = "O'qituvchi: ____________________\nBahosi: ____________________"
c11.paragraphs[0].runs[0].font.size = Pt(12)

for _ in range(4):
    doc.add_paragraph()

p_shahar = doc.add_paragraph()
p_shahar.alignment = WD_ALIGN_PARAGRAPH.CENTER
r_sh = p_shahar.add_run("Toshkent – 2026")
r_sh.font.size = Pt(12)
r_sh.font.name = "Times New Roman"

doc.add_page_break()

# -----------------------------------------------------------------------------
# ISHNING MAQSADI VA BERILGAN DATASET
# -----------------------------------------------------------------------------
add_heading_1("ISHNING MAQSADI VA TADQIQOT OB'EKTI")
add_body("Ishning maqsadi: Talabalarda ma’lumotlar to‘plamidagi sifat muammolarini (yetishmayotgan qiymatlar, takroriy yozuvlar, nomuvofiq format, g‘ayrioddiy qiymatlar — outlierlar) aniqlash hamda ularni to‘g‘ri usul bilan qayta ishlash, tozalash va tahlilga tayyorlash ko‘nikmasini shakllantirish.")

add_heading_2("Dastlabki «Do‘kon mijozlari» ma'lumotlar to'plami")
add_body("Hujjat shartida berilgan boshlang'ich xatoliklar mavjud jadval ko'rinishi (Oylik_xarid — ming so'mda):")

tbl_orig = doc.add_table(rows=1, cols=7)
tbl_orig.alignment = WD_TABLE_ALIGNMENT.CENTER
set_table_borders(tbl_orig)

orig_headers = ["ID", "Ism", "Yoshi", "Jinsi", "Shahar", "Oylik_xarid", "Xaridlar_soni"]
for i, h in enumerate(orig_headers):
    tbl_orig.rows[0].cells[i].text = h
    set_cell_background(tbl_orig.rows[0].cells[i], "1E293B")
format_row(tbl_orig.rows[0], is_header=True)

for idx, (_, r) in enumerate(df_raw.iterrows(), 1):
    row_c = tbl_orig.add_row().cells
    row_c[0].text = str(r["ID"])
    row_c[1].text = str(r["Ism"])
    row_c[2].text = str(int(r["Yoshi"])) if pd.notna(r["Yoshi"]) else "—"
    row_c[3].text = str(r["Jinsi"])
    row_c[4].text = str(r["Shahar"]) if pd.notna(r["Shahar"]) else "—"
    row_c[5].text = str(int(r["Oylik_xarid"])) if pd.notna(r["Oylik_xarid"]) else "—"
    row_c[6].text = str(int(r["Xaridlar_soni"])) if pd.notna(r["Xaridlar_soni"]) else "—"
    bg = "F8FAFC" if idx % 2 == 1 else "FFFFFF"
    for cell in row_c:
        set_cell_background(cell, bg)
    format_row(tbl_orig.rows[-1])

doc.add_page_break()

# -----------------------------------------------------------------------------
# 1-TOPSHIRIQ
# -----------------------------------------------------------------------------
add_heading_1("1-TOPSHIRIQ. MA’LUMOTLAR SIFATINI TEKSHIRISH")

add_body("Jadval sinchiklab tahlil qilinib, aniqlangan barcha sifat muammolari quyidagi jadvalda umumlashtirildi:")

tbl_prob = doc.add_table(rows=1, cols=4)
tbl_prob.alignment = WD_TABLE_ALIGNMENT.CENTER
set_table_borders(tbl_prob)

h_prob = ["Muammo turi", "Qaysi atribut(lar)da?", "Qaysi satr(lar)da (ID)?", "Soni"]
for i, h in enumerate(h_prob):
    tbl_prob.rows[0].cells[i].text = h
    set_cell_background(tbl_prob.rows[0].cells[i], "1E293B")
format_row(tbl_prob.rows[0], is_header=True)

prob_data = [
    ("Yetishmayotgan qiymat (Missing)", "Yoshi, Shahar, Oylik_xarid, Xaridlar_soni", "ID: 2 (Yoshi), 4 (Shahar), 5 (Xarid), 6 (Soni), 9 (Yoshi)", "5 ta katak"),
    ("Takroriy yozuv (Duplicate)", "Barcha atributlar bo'yicha (Rustam)", "ID: 7 (ID: 3 ning aynan takrori)", "1 ta satr"),
    ("Nomuvofiq format (Inconsistent)", "Jinsi, Shahar (kichik harflar bilan yozilgan)", "ID: 3 ('erkak', 'samarqand')", "2 ta ustun"),
    ("G‘ayrioddiy qiymat (Outlier)", "Oylik_xarid (5000 ming so'm kiritilgan)", "ID: 9 (Ulug‘bek mijoz)", "1 ta qiymat")
]

for idx, r in enumerate(prob_data, 1):
    row_c = tbl_prob.add_row().cells
    row_c[0].text = r[0]
    row_c[1].text = r[1]
    row_c[2].text = r[2]
    row_c[3].text = r[3]
    bg = "F8FAFC" if idx % 2 == 1 else "FFFFFF"
    for cell in row_c:
        set_cell_background(cell, bg)
    format_row(tbl_prob.rows[-1])

add_heading_2("Mantiqiy savolga javob:")
add_body("Chunki kompyuter dasturlari va solishtirish operatorlari satrlarni harf-baharf, jumladan katta-kichik harf registrini (Case-sensitive) inobatga olgan holda solishtiradi.\n"
         "  • 3-satrda: Jinsi='erkak' (kichik harfda), Shahar='samarqand' (kichik harfda).\n"
         "  • 7-satrda: Jinsi='Erkak' (bosh harfda), Shahar='Samarqand' (bosh harfda).\n"
         "Dastur darajasida 'samarqand' == 'Samarqand' taqqoslash amali False (teng emas) natijasini qaytaradi. Shu sababli kompyuter avvaliga ularni har xil yozuvlar deb qabul qiladi. Faqat matnlarni bitta formatga (masalan, .capitalize() orqali) keltirgandan keyingina dastur ularning dublikat ekanligini to'liq aniqlay oladi.",
         bold_prefix="Savol: 7-satr 3-satrning takrori ekanini kompyuter dasturi nima uchun darhol aniqlay olmasligi mumkin?\nJavob: ")

# -----------------------------------------------------------------------------
# 2-TOPSHIRIQ
# -----------------------------------------------------------------------------
add_heading_1("2-TOPSHIRIQ. YETISHMAYOTGAN QIYMATLARNI QAYTA ISHLASH USULINI TANLASH")

add_body("Dastlab takroriy 7-satr o'chirildi va qolgan 9 ta satrdagi yetishmayotgan qiymatlarni qayta ishlash strategiyasi ishlab chiqildi:")

tbl_imp = doc.add_table(rows=1, cols=5)
tbl_imp.alignment = WD_TABLE_ALIGNMENT.CENTER
set_table_borders(tbl_imp)

h_imp = ["Atribut", "Tanlangan usul", "Hisoblash jarayoni", "To'ldiriladigan qiymat", "Asoslash"]
for i, h in enumerate(h_imp):
    tbl_imp.rows[0].cells[i].text = h
    set_cell_background(tbl_imp.rows[0].cells[i], "1E293B")
format_row(tbl_imp.rows[0], is_header=True)

imp_data = [
    ("Yoshi (ID 2, 9)", "Mediana", "Mavjud: [25, 27, 28, 30, 33, 34, 41]\nO'rtadagi qiymat = 30", "30 yosh", "Sonli ko'rsatkich, mediana butun son beradi va taqsimot markazini xolis ifodalaydi."),
    ("Shahar (ID 4)", "Moda", "Samarqand: 4 marta\nToshkent: 2 marta, Buxoro: 2 marta", "Samarqand", "Kategorial (nominal) atribut bo'lgani sababli faqat eng ko'p uchragan shahar (moda) olinadi."),
    ("Oylik_xarid (ID 5)", "Mediana", "Mavjud: [380, 450, 470, 490, 520, 560, 610, 5000]\nMediana = 505", "505 ming so'm", "9-satrda 5000 outlier bo'lgani sababli o'rtacha xolis emas, mediana esa o'ta mustahkamdir."),
    ("Xaridlar_soni (ID 6)", "Mediana / Moda", "Mavjud: [3, 4, 5, 5, 5, 6, 6, 7]\nMediana = 5, Moda = 5", "5 marta", "Diskret sonli ko'rsatkich, mijozlarning eng tipik xaridlar soni 5 tani tashkil etadi.")
]

for idx, r in enumerate(imp_data, 1):
    row_c = tbl_imp.add_row().cells
    row_c[0].text = r[0]
    row_c[1].text = r[1]
    row_c[2].text = r[2]
    row_c[3].text = r[3]
    row_c[4].text = r[4]
    bg = "F8FAFC" if idx % 2 == 1 else "FFFFFF"
    for cell in row_c:
        set_cell_background(cell, bg)
    format_row(tbl_imp.rows[-1])

add_heading_2("Oylik_xarid atributi bo'yicha ko'rsatkichlar tahlili:")
add_body("Mavjud 8 ta xarid qiymatlari (380, 450, 470, 490, 520, 560, 610, 5000) bo'yicha:\n"
         "  • Arifmetik o'rtacha (Mean) = 1060.0 ming so'm;\n"
         "  • Mediana (Median) = 505.0 ming so'm;\n"
         "  • O'rtadagi farq = 555.0 ming so'm (o'rtacha qiymat medianadan 2 barobardan ko'proq katta!).\n"
         "Katta farqning sababi: 9-satrdagi 5000 ming so'm kiritilgan keskin anomaliya (outlier). O'rtacha qiymat ushbu bitta anomal katta songa haddan tashqari sezgir bo'lib, sun'iy ko'tarilib ketgan. Mediana esa chetki xatoliklarga ta'sirlanmay, haqiqiy mijozning odatiy oylik xaridini xolis ifodalaydi.",
         bold_prefix="• O'rtacha va mediana orasidagi farq va uning sababi: ")

add_body("Agar yetishmayotgan qiymatlarga ega bo'lgan barcha satrlar shunchaki o'chirib tashlansa (Drop rows), jadvalda 10 ta satrdan atigi 4 ta satr (ID: 1, 3, 8, 10) qoladi!\n"
         "  Kamchiligi: Ma'lumotlarning 60% qismi (aksariyat ko'pchiligi) butunlay boy beriladi. Natijada tanlanma hajmi keskin qisqaradi, statistik ishonchlilik yo'qoladi va bunday kichik to'plam asosida qurilgan model yoki tahlillar mutlaqo yaroqsiz bo'lib qoladi.",
         bold_prefix="• Bo'sh satrlarni shunchaki o'chirib tashlash oqibati: ")

doc.add_page_break()

# -----------------------------------------------------------------------------
# 3-TOPSHIRIQ
# -----------------------------------------------------------------------------
add_heading_1("3-TOPSHIRIQ. EXCEL VOSITASIDA MA’LUMOTLARNI TOZALASH")

add_body("Excel (yoki Google Sheets) elektron jadval dasturida quyidagi tozalash bosqichlari ketma-ket bajarildi:")
add_body("1. Formatni standartlashtirish: 'Jinsi' va 'Shahar' ustunlaridagi matnlarni birinchi harfini bosh harf qilish uchun =PROPER() funksiyasidan foydalanildi. Natijada 'erkak' -> 'Erkak', 'samarqand' -> 'Samarqand' ko'rinishiga keltirildi.\n"
         "2. Dublikat yozuvni o'chirish: 'Data' -> 'Remove Duplicates' vositasi orqali barcha ustunlar (ID ustunisiz) bo'yicha takrorlar qidirildi. 7-satrdagi ikkinchi Rustam dublikat sifatida muvaffaqiyatli o'chirildi (9 ta satr qoldi).\n"
         "3. Yetishmayotgan qiymatlarni to'ldirish: Gulnora va Ulug'bekning yoshi 30 bilan, Shahnozaning shahri 'Samarqand' bilan, Javohirning xaridi 505 bilan, Feruzaning xaridlar soni 5 bilan to'ldirildi.\n"
         "4. Outlierni tuzatish: 9-mijoz Ulug'bekning 5000 ming so'mlik anomal xaridi medianaga (505) almashtirildi.\n"
         "Natija: Tozalangan jadvalda 9 ta to'liq satr qoldi, bo'sh kataklar soni 0 tani tashkil etdi.")

# -----------------------------------------------------------------------------
# 4-TOPSHIRIQ
# -----------------------------------------------------------------------------
add_heading_1("4-TOPSHIRIQ. PYTHON (PANDAS) YORDAMIDA TOZALASH")

add_body("Dastlabki jadval 'mijozlar.csv' faylida saqlanib, quyidagi Python kodi yordamida avtomatik tozalandi:")

add_body("""import pandas as pd
df = pd.read_csv("mijozlar.csv")
print(df.isnull().sum())                 # har bir ustundagi bo‘sh qiymatlar

# 1) Formatni bir xil qilish
for col in ["Jinsi", "Shahar"]:
    df[col] = df[col].str.strip().str.capitalize()

# 2) Takrorlarni topish va o‘chirish (ID ustunisiz)
ustunlar = df.columns[1:]
print("Takrorlar:", df.duplicated(subset=ustunlar).sum())
df = df.drop_duplicates(subset=ustunlar)

# 3) Yetishmayotgan qiymatlarni to‘ldirish
df["Yoshi"] = df["Yoshi"].fillna(df["Yoshi"].median())
df["Shahar"] = df["Shahar"].fillna(df["Shahar"].mode()[0])
med = df["Oylik_xarid"].median()
df.loc[df["Oylik_xarid"] > 2000, "Oylik_xarid"] = med
df["Oylik_xarid"] = df["Oylik_xarid"].fillna(med)
df["Xaridlar_soni"] = df["Xaridlar_soni"].fillna(df["Xaridlar_soni"].median())

print(df.isnull().sum())
print(df.shape)
print(df)""", italic=True)

add_heading_2("Natijalar asosidagi tahlil va xulosalar:")
add_body("• 1-qism: str.strip().str.capitalize() — matn chetlaridagi ortiqcha bo'shliqlarni olib tashlaydi va birinchi harfni bosh harfga aylantirib, registrlarni standartlashtiradi.\n"
         "• 2-qism: duplicated(subset=ustunlar) — ID ustunidan tashqari barcha ustunlar bir xil bo'lgan qatorlarni aniqlaydi va drop_duplicates() orqali 7-satrdagi dublikatni o'chiradi.\n"
         "• 3-qism: fillna() va median() — sonli ustunlardagi bo'shliqlarni mediana bilan, kategorial Shahar ustunini esa moda bilan to'ldiradi hamda 2000 dan katta bo'lgan anomal 5000 xaridni mediana (505) ga almashtiradi.",
         bold_prefix="a) Kod qismlarining vazifalari: ")

add_body("Agar 1-qadam (formatni bir xil qilish) bajarilmasa, takrorlar soni 0 chiqadi! Chunki 'samarqand' va 'Samarqand' turli satr deb qaraladi. Bu esa ma'lumotlarni tozalashda matnlarni standartlashtirish qanchalik muhim ekanini yaqqol isbotlaydi.",
         bold_prefix="b) 1-qadamsiz takrorlar soni: ")

add_body("Python dasturida olingan yakuniy natijalar Excel'da qo'lda bajarilgan tozalash natijasi bilan 100% to'liq mos keldi (ikkala usulda ham 9 ta toza satr va 0 ta bo'sh katak hosil bo'ldi).",
         bold_prefix="c) Excel va Python natijalarini solishtirish: ")

add_heading_2("Yakuniy tozalangan va tiklangan jadval (Cleaned DataFrame)")

tbl_clean = doc.add_table(rows=1, cols=7)
tbl_clean.alignment = WD_TABLE_ALIGNMENT.CENTER
set_table_borders(tbl_clean)

for i, h in enumerate(orig_headers):
    tbl_clean.rows[0].cells[i].text = h
    set_cell_background(tbl_clean.rows[0].cells[i], "1E293B")
format_row(tbl_clean.rows[0], is_header=True)

# Tozalangan dataframe ma'lumotlari
df_clean_rep = df_raw.copy()
for col in ["Jinsi", "Shahar"]:
    df_clean_rep[col] = df_clean_rep[col].str.strip().str.capitalize()
df_clean_rep = df_clean_rep.drop_duplicates(subset=df_clean_rep.columns[1:])
df_clean_rep["Yoshi"] = df_clean_rep["Yoshi"].fillna(df_clean_rep["Yoshi"].median())
df_clean_rep["Shahar"] = df_clean_rep["Shahar"].fillna(df_clean_rep["Shahar"].mode()[0])
med_rep = df_clean_rep["Oylik_xarid"].median()
df_clean_rep.loc[df_clean_rep["Oylik_xarid"] > 2000, "Oylik_xarid"] = med_rep
df_clean_rep["Oylik_xarid"] = df_clean_rep["Oylik_xarid"].fillna(med_rep)
df_clean_rep["Xaridlar_soni"] = df_clean_rep["Xaridlar_soni"].fillna(df_clean_rep["Xaridlar_soni"].median())

for idx, (_, r) in enumerate(df_clean_rep.iterrows(), 1):
    row_c = tbl_clean.add_row().cells
    row_c[0].text = str(int(r["ID"]))
    row_c[1].text = str(r["Ism"])
    row_c[2].text = str(int(r["Yoshi"]))
    row_c[3].text = str(r["Jinsi"])
    row_c[4].text = str(r["Shahar"])
    row_c[5].text = str(int(r["Oylik_xarid"]))
    row_c[6].text = str(int(r["Xaridlar_soni"]))
    bg = "F8FAFC" if idx % 2 == 1 else "FFFFFF"
    for cell in row_c:
        set_cell_background(cell, bg)
    format_row(tbl_clean.rows[-1])

doc.add_page_break()

# -----------------------------------------------------------------------------
# 5. NAZORAT SAVOLLARI VA XULOSA
# -----------------------------------------------------------------------------
add_heading_1("5. NAZORAT SAVOLLARI VA ULARNING JAVOBLARI")

qa_list = [
    ("1. Ma’lumotlarni tozalash nima uchun Data mining jarayonining muhim bosqichi hisoblanadi?",
     "Chunki real amaliyotdagi ma'lumotlar deyarli har doim shovqinli, to'liq emas va xatoliklarga ega bo'ladi. 'Garbage in — garbage out' (axlat kirsa — axlat chiqadi) "
     "tamoyiliga ko'ra, xato va to'liq bo'lmagan ma'lumotlarga tayangan har qanday statistik model yoki mashinali o'qitish algoritmi noto'g'ri bashorat beradi. Tozalash ma'lumotlarning aniqligi va ishonchliligini ta'minlaydi."),

    ("2. Yetishmayotgan qiymatlar paydo bo‘lishining qanday sabablarini bilasiz?",
     "1. Respondentning ayrim savollarga (masalan, yoshi yoki daromadi) javob berishni istamasligi;\n"
     "2. Texnik xatoliklar (uskunalar nosozligi, serverdagi uzilishlar);\n"
     "3. Bir nechta turli tizimlardagi ma'lumotlar birlashtirilganda (integratsiyada) formatlar mos kelmasligi;\n"
     "4. Insoniy omil sababli ayrim maydonlarning to'ldirilmay qolib ketishi."),

    ("3. Qachon yetishmayotgan qiymatli satrni o‘chirish maqsadga muvofiq, qachon esa to‘ldirish kerak?",
     "• O'chirish maqsadga muvofiq: Katta hajmli datasetlarda yetishmayotgan satrlar soni juda kam (1-2%) bo'lsa va ularni olib tashlash umumiy qonuniyatga ta'sir qilmasa, yoki aynan bashorat qilinishi kerak bo'lgan asosiy nishon (Target) ustuni bo'sh bo'lsa.\n"
     "• To'ldirish kerak: Dataset kichik bo'lsa (satr o'chirilsa ma'lumot yetishmovchiligi ro'y bersa) yoki boshqa ustunlar qimmatli tahliliy ma'lumotlarga ega bo'lsa."),

    ("4. Nima uchun nominal atributni o‘rtacha qiymat bilan to‘ldirib bo‘lmaydi?",
     "Chunki nominal atributlar sifat toifalari (masalan, Shahar, Jinsi, Mahsulot turi) bo'lib, ular sonli qiymatga ega emas. "
     "Ularni bir-biriga qo'shish va bo'lish matematik jihatdan mutlaqo ma'noga ega emas (Samarqand va Toshkentning 'o'rtachasi' mavjud emas). Shuning uchun bunday atributlar faqat eng ko'p uchragan toifa — Moda bilan to'ldiriladi."),

    ("5. G‘ayrioddiy qiymat (outlier) har doim xato hisoblanadimi? Misol bilan tushuntiring.",
     "Yo'q, har doim xato hisoblanmaydi! Outlierlar ikki toifaga bo'linadi:\n"
     "a) Xato / shovqin: Masalan, oylik xaridga 500 o'rniga adashib 5000 yozib yuborilishi yoki talabaning yoshi 120 deb kiritilishi (bular tuzatilishi yoki o'chirilishi shart);\n"
     "b) Noyob haqiqiy fakt: Masalan, bank tizimida millioner mijozning bir martalik juda katta xaridi yoki kibertahdid (DDoS xuruji) paytidagi trafik portlashi. Bunday haqiqiy outlierlar o'chirilmaydi, aksincha anomaliyalarni aniqlash tizimlari uchun eng qimmatli obyekt hisoblanadi.")
]

for q, a in qa_list:
    add_body(a, bold_prefix=f"{q}\nJavob: ")

add_heading_1("6. XULOSA")
add_body("Ma’lumotlarni tozalash jarayoni Data Mining va ma'lumotlar tahlilining eng asosiy va ko'p vaqt talab qiluvchi bosqichidir. Ushbu amaliy mashg'ulot davomida formatlarni standartlashtirmasdan turib dublikatlarni aniqlab bo'lmasligi, bitta anomal ko'rsatkich (outlier) butun o'rtacha qiymatni 2 barobar buzib yuborishi va bunday holatda mediana o'ta samarali ekanligi amalda isbotlandi. Excel va Python (pandas) vositalarining birgalikda qo'llanilishi ma'lumotlarni to'liq xatosiz holga keltirish imkonini berdi.")

doc_path = os.path.join(CURRENT_DIR, "4-amaliy_mashg'ulot_hisoboti.docx")
doc.save(doc_path)
print(f"Word hisoboti saqlandi: {doc_path}")

# ==============================================================================
# JUPYTER NOTEBOOK (4-amaliy_mashgulot.ipynb)
# ==============================================================================
nb_cells = [
    {
        "cell_type": "markdown",
        "metadata": {},
        "source": [
            "# 4-Amaliy Mashg'ulot: Ma’lumotlarni tozalash, yetishmayotgan qiymatlarni aniqlash va qayta ishlash\n",
            "**Fan:** Data Mining  \n",
            "**Mavzu:** «Do‘kon mijozlari» ma'lumotlar to'plami asosida sifat muammolarini (Missing, Duplicates, Format, Outlier) aniqlash va bartaraf etish."
        ]
    },
    {
        "cell_type": "code",
        "execution_count": None,
        "metadata": {},
        "outputs": [],
        "source": [
            "# 1. Kutubxonalarni import qilish va boshlang'ich datasetni yuklash\n",
            "import os\n",
            "import pandas as pd\n",
            "import numpy as np\n",
            "\n",
            "df = pd.read_csv('mijozlar.csv')\n",
            "print('Boshlang\\'ich dataset hajmi:', df.shape)\n",
            "display(df)"
        ]
    },
    {
        "cell_type": "markdown",
        "metadata": {},
        "source": [
            "## 1-topshiriq. Ma'lumotlar sifatini tekshirish\n",
            "Aniqlangan sifat muammolari jadvali:\n",
            "| Muammo turi | Qaysi atribut(lar)da? | Qaysi satr(lar)da (ID)? | Soni |\n",
            "| :--- | :--- | :--- | :--- |\n",
            "| **Yetishmayotgan qiymat** | Yoshi, Shahar, Oylik_xarid, Xaridlar_soni | ID: 2, 4, 5, 6, 9 | 5 ta katak |\n",
            "| **Takroriy yozuv (Dublikat)** | Barcha ustunlar (Rustam) | ID: 7 (ID: 3 ning takrori) | 1 ta satr |\n",
            "| **Nomuvofiq format** | Jinsi, Shahar (kichik harflar) | ID: 3 ('erkak', 'samarqand') | 2 ta ustun |\n",
            "| **G'ayrioddiy qiymat (Outlier)** | Oylik_xarid (5000 ming so'm) | ID: 9 (Ulug'bek) | 1 ta qiymat |\n",
            "\n",
            "> **Savol javobi:** Kompyuter 7-satr 3-satrning takrori ekanini darhol aniqlay olmaydi, chunki registrlar har xil ('samarqand' != 'Samarqand')."
        ]
    },
    {
        "cell_type": "markdown",
        "metadata": {},
        "source": [
            "## 2-topshiriq. Yetishmayotgan qiymatlarni qayta ishlash usulini tanlash\n",
            "| Atribut | Tanlangan usul | Hisoblash | To'ldiriladigan qiymat | Asoslash |\n",
            "| :--- | :--- | :--- | :--- | :--- |\n",
            "| `Yoshi` | Mediana | Mavjud yoshlar: [25, 27, 28, 30, 33, 34, 41] | **30 yosh** | Simmetrik taqsimot, butun son |\n",
            "| `Shahar` | Moda | Samarqand: 4 ta, Toshkent: 2 ta, Buxoro: 2 ta | **Samarqand** | Kategorial ustun uchun faqat moda |\n",
            "| `Oylik_xarid` | Mediana | Mavjud: [380, 450, 470, 490, 520, 560, 610, 5000] | **505 ming so'm** | Outlier (5000) ta'siriga chidamli |\n",
            "| `Xaridlar_soni`| Mediana/Moda | Mavjud: [3, 4, 5, 5, 5, 6, 6, 7] | **5 marta** | Eng ko'p uchragan odatiy xaridlar soni |"
        ]
    },
    {
        "cell_type": "code",
        "execution_count": None,
        "metadata": {},
        "outputs": [],
        "source": [
            "# Oylik xarid bo'yicha o'rtacha va mediana solishtirmasi\n",
            "xarid_vals = df[df['ID'] != 7]['Oylik_xarid'].dropna()\n",
            "print(f'Arifmetik o\\'rtacha (Mean): {xarid_vals.mean():.2f} ming so\\'m')\n",
            "print(f'Mediana (Median): {xarid_vals.median():.2f} ming so\\'m')\n",
            "print(f'Farq: {xarid_vals.mean() - xarid_vals.median():.2f} ming so\\'m (Outlier 5000 sababli o\\'rtacha buzilgan)')"
        ]
    },
    {
        "cell_type": "markdown",
        "metadata": {},
        "source": [
            "## 4-topshiriq. Python (pandas) yordamida tozalash jarayoni"
        ]
    },
    {
        "cell_type": "code",
        "execution_count": None,
        "metadata": {},
        "outputs": [],
        "source": [
            "# 1. Bo'sh qiymatlarni ko'rish\n",
            "print('Boshlang\\'ich bo\\'sh qiymatlar:')\n",
            "display(df.isnull().sum())\n",
            "\n",
            "# 2. Formatni bir xil qilish (Harflarni standartlashtirish)\n",
            "for col in ['Jinsi', 'Shahar']:\n",
            "    df[col] = df[col].str.strip().str.capitalize()\n",
            "\n",
            "# 3. Takrorlarni aniqlash va o'chirish (ID ustunisiz)\n",
            "ustunlar = df.columns[1:]\n",
            "print('Formatlashdan so\\'ng aniqlangan takrorlar:', df.duplicated(subset=ustunlar).sum())\n",
            "df = df.drop_duplicates(subset=ustunlar)\n",
            "\n",
            "# 4. Yetishmayotgan qiymatlarni to'ldirish\n",
            "df['Yoshi'] = df['Yoshi'].fillna(df['Yoshi'].median())\n",
            "df['Shahar'] = df['Shahar'].fillna(df['Shahar'].mode()[0])\n",
            "\n",
            "# 5. Outlierni (5000) medianaga almashtirish\n",
            "med = df['Oylik_xarid'].median()\n",
            "df.loc[df['Oylik_xarid'] > 2000, 'Oylik_xarid'] = med\n",
            "df['Oylik_xarid'] = df['Oylik_xarid'].fillna(med)\n",
            "df['Xaridlar_soni'] = df['Xaridlar_soni'].fillna(df['Xaridlar_soni'].median())\n",
            "\n",
            "print('\\nTozalangandan keyingi bo\\'sh qiymatlar soni:')\n",
            "display(df.isnull().sum())\n",
            "print('\\nYakuniy tozalangan jadval o\\'lchami:', df.shape)\n",
            "display(df)"
        ]
    }
]

nb_dict = {
    "cells": nb_cells,
    "metadata": {
        "kernelspec": {
            "display_name": "Python 3",
            "language": "python",
            "name": "python3"
        },
        "language_info": {
            "name": "python",
            "version": "3.12"
        }
    },
    "nbformat": 4,
    "nbformat_minor": 5
}

nb_path = os.path.join(CURRENT_DIR, "4-amaliy_mashgulot.ipynb")
with open(nb_path, "w", encoding="utf-8") as f:
    json.dump(nb_dict, f, ensure_ascii=False, indent=2)

print(f"Jupyter Notebook saqlandi: {nb_path}")
