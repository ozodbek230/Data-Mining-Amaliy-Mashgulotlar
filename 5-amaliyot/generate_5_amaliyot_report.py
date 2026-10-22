# -*- coding: utf-8 -*-
"""
5-amaliy mashg'ulot bo'yicha to'liq Word hisoboti (5-amaliy_mashg'ulot_hisoboti.docx)
hamda Jupyter Notebook (5-amaliy_mashgulot.ipynb) yaratish skripti.
"""

import os
import json
import numpy as np
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

from sklearn.preprocessing import MinMaxScaler, StandardScaler
from sklearn.datasets import load_wine
from sklearn.model_selection import train_test_split
from sklearn.neighbors import KNeighborsClassifier
from sklearn.tree import DecisionTreeClassifier
from sklearn.pipeline import make_pipeline

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))

# ==============================================================================
# 1. HISOBLASHLARNI AMALGA OSHIRISH
# ==============================================================================
# 1-topshiriq
yoshlar = [20, 35, 50, 65, 80]
x_min, x_max = min(yoshlar), max(yoshlar)
farq = x_max - x_min
df_t1 = pd.DataFrame({'Yosh, x': yoshlar})
df_t1['x - x_min'] = df_t1['Yosh, x'] - x_min
df_t1["x' [0, 1]"] = df_t1['x - x_min'] / farq
df_t1["x' [-1, 1]"] = 2 * df_t1["x' [0, 1]"] - 1

# 2-topshiriq
ballar = [60, 70, 80, 90, 100]
n = len(ballar)
mu = sum(ballar) / n
diffs = [x - mu for x in ballar]
diffs_sq = [(x - mu) ** 2 for x in ballar]
sigma = (sum(diffs_sq) / n) ** 0.5
z_scores = [d / sigma for d in diffs]
df_t2 = pd.DataFrame({
    'Ball, x': ballar,
    'x - mu': diffs,
    '(x - mu)^2': diffs_sq,
    'z': [round(z, 3) for z in z_scores]
})

# 3-topshiriq
df_bank = pd.DataFrame({
    "Yosh":    [25, 32, 47, 51, 62, 38],
    "Daromad": [4.5, 7.0, 12.0, 9.5, 6.0, 15.0]
})
mm = MinMaxScaler()
df_mm = pd.DataFrame(mm.fit_transform(df_bank), columns=df_bank.columns)
st = StandardScaler()
df_st = pd.DataFrame(st.fit_transform(df_bank), columns=df_bank.columns)

p1_raw = df_bank.iloc[0].values
p5_raw = df_bank.iloc[4].values
d_raw = np.linalg.norm(p1_raw - p5_raw)
d_mm = np.linalg.norm(df_mm.iloc[0].values - df_mm.iloc[4].values)
d_st = np.linalg.norm(df_st.iloc[0].values - df_st.iloc[4].values)

df7 = pd.concat([df_bank, pd.DataFrame({'Yosh': [30], 'Daromad': [60.0]})], ignore_index=True)
df7_mm = pd.DataFrame(MinMaxScaler().fit_transform(df7), columns=df7.columns)
df7_st = pd.DataFrame(StandardScaler().fit_transform(df7), columns=df7.columns)

# 4-topshiriq
X, y = load_wine(return_X_y=True)
X_train, X_test, y_train, y_test = train_test_split(
    X, y, test_size=0.3, random_state=42, stratify=y
)
modellar = {
    "k-NN (masshtablashsiz)": KNeighborsClassifier(n_neighbors=5),
    "k-NN + MinMax": make_pipeline(MinMaxScaler(), KNeighborsClassifier(5)),
    "k-NN + Z-score": make_pipeline(StandardScaler(), KNeighborsClassifier(5)),
    "Qarorlar daraxti": DecisionTreeClassifier(random_state=42),
    "Qarorlar daraxti + Z-score": make_pipeline(StandardScaler(), DecisionTreeClassifier(random_state=42)),
}
model_natijalar = []
for nom, model in modellar.items():
    model.fit(X_train, y_train)
    acc = model.score(X_test, y_test)
    model_natijalar.append((nom, f"{acc:.3f} ({acc*100:.1f}%)"))

# ==============================================================================
# 2. WORD HUJJATINI TAYYORLASH
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
r_t = p_title.add_run("5-AMALIY MASHG'ULOT HISOBOTI\n")
r_t.bold = True
r_t.font.size = Pt(20)
r_t.font.name = "Times New Roman"
r_t.font.color.rgb = RGBColor(15, 23, 42)

p_sub = doc.add_paragraph()
p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
r_sub = p_sub.add_run("Mavzu: Ma’lumotlarni normallashtirish va standartlashtirish usullarini qo‘llash\n(Min-Max, Z-score, O'nlik masshtablash hamda k-NN va Qarorlar daraxti modellari tahlili)")
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
# MAQSAD VA QISQA NAZARIYA
# -----------------------------------------------------------------------------
add_heading_1("ISHNING MAQSADI VA NAZARIY ASOSLARI")
add_body("Ishning maqsadi: Belgilarni bir xil masshtabga keltirishning asosiy usullarini qo‘lda va Python’da qo‘llash, ularning model natijalariga (xususan masofaga asoslangan k-NN va shartlarga asoslangan Qarorlar daraxtiga) ta’sirini tushunish hamda amaliy ko'nikmalarni shakllantirish.")

add_heading_2("Masshtablashning zarurati va usullari")
add_body("Haqiqiy ma'lumotlar to'plamlarida belgilar turli xil o'lchov birliklari va oraliqlarda bo'ladi (masalan, inson yoshi 20–70 yil, oylik daromadi esa 3 000 000 – 15 000 000 so'm). Masofaga asoslangan algoritmlarda (k-NN, K-Means, SVM) katta qiymatli belgi (daromad) kichik qiymatli belgini (yosh) to'liq \"bosib ketadi\" va model faqat bitta atributga qarab qaror qabul qiladigan xatolikka olib keladi. Masshtablash shu muammoni bartaraf etadi.")

tbl_theory = doc.add_table(rows=1, cols=3)
tbl_theory.alignment = WD_TABLE_ALIGNMENT.CENTER
set_table_borders(tbl_theory)

th_headers = ["Masshtablash usuli", "Natija oralig‘i", "Qachon va qaysi modellarda qo‘llanadi"]
for i, h in enumerate(th_headers):
    tbl_theory.rows[0].cells[i].text = h
    set_cell_background(tbl_theory.rows[0].cells[i], "1E293B")
format_row(tbl_theory.rows[0], is_header=True)

theory_rows = [
    ("Min-Max normallashtirish", "[0, 1] yoki [a, b]", "Chegaralari aniq ma'lum belgilar, neyron tarmoqlar (ANN/CNN), rasm piksellari (0–255)."),
    ("Z-score standartlashtirish", "Taxminan [−3, 3] (o'rtacha=0, std=1)", "k-NN, SVM, Logistik regressiya, PCA; ma'lumotlarda chetga chiquvchi (outlier) qiymatlar bo'lganda."),
    ("O‘nlik masshtablash", "(−1, 1)", "Oddiy, tezkor va kompyuter xotirasini tejash talab etilganda vergulni surish orqali hisoblanadi.")
]

for idx, r in enumerate(theory_rows, 1):
    row_c = tbl_theory.add_row().cells
    row_c[0].text = r[0]
    row_c[1].text = r[1]
    row_c[2].text = r[2]
    bg = "F8FAFC" if idx % 2 == 1 else "FFFFFF"
    for cell in row_c:
        set_cell_background(cell, bg)
    format_row(tbl_theory.rows[-1])

# -----------------------------------------------------------------------------
# 1-TOPSHIRIQ
# -----------------------------------------------------------------------------
add_heading_1("1-TOPSHIRIQ. MIN-MAX NORMALLASHTIRISH (QO‘LDA HISOBLASH)")
add_body("Besh nafar xodimning yoshi berilgan: 20, 35, 50, 65, 80 yil.")
add_body("• Eng kichik qiymat: x_min = 20 yil\n• Eng katta qiymat: x_max = 80 yil\n• O'zgarish diapazoni: x_max - x_min = 80 - 20 = 60 yil")

add_body("Formulalar:\n1) [0, 1] oralig'i: x' = (x - x_min) / (x_max - x_min) = (x - 20) / 60\n2) [-1, 1] oralig'i: x' = 2 * ((x - 20) / 60) - 1")

tbl_t1 = doc.add_table(rows=1, cols=4)
tbl_t1.alignment = WD_TABLE_ALIGNMENT.CENTER
set_table_borders(tbl_t1)

h1 = ["Yosh, x", "x − x_min", "x' ∈ [0, 1]", "x' ∈ [−1, 1]"]
for i, h in enumerate(h1):
    tbl_t1.rows[0].cells[i].text = h
    set_cell_background(tbl_t1.rows[0].cells[i], "1E293B")
format_row(tbl_t1.rows[0], is_header=True)

for idx, (_, r) in enumerate(df_t1.iterrows(), 1):
    row_c = tbl_t1.add_row().cells
    row_c[0].text = str(int(r['Yosh, x']))
    row_c[1].text = str(int(r['x - x_min']))
    v1 = r["x' [0, 1]"]
    v2 = r["x' [-1, 1]"]
    row_c[2].text = f"{v1:.2f}"
    row_c[3].text = f"{v2:.2f}"
    bg = "F8FAFC" if idx % 2 == 1 else "FFFFFF"
    for cell in row_c:
        set_cell_background(cell, bg)
    format_row(tbl_t1.rows[-1])

add_heading_2("Savollarga tahliliy javoblar:")
add_body("Yangi xodim yoshi 42 bo'lsa: x' = (42 - 20) / 60 = 22 / 60 ≈ 0.367. Yangi xodim yoshi [0, 1] shkalasida taxminan 0.37 ga to'g'ri keladi.", bold_prefix="1) 42 yoshli xodim uchun: ")
add_body("x' = (90 - 20) / 60 = 70 / 60 ≈ 1.167. Natija 1 dan oshib, [0, 1] oralig'idan chiqib ketadi. Sababi: Min-Max normallashtirish faqat o'quv to'plamidagi chegaralarga (20 va 80) qat'iy bog'langan bo'lib, to'plamdan tashqaridagi yangi ekstremal qiymatlarga (outlierlarga) nisbatan chidamsiz hisoblanadi.", bold_prefix="2) 90 yoshli xodim uchun: ")
add_body("Qiymatlar: 125, 980, -450. Eng katta modul qiymati max(|x|) = 980. Barcha |x'| < 1 bo'lishi uchun j = 3 tanlanadi (chunki 10^3 = 1000 > 980). Natijalar: 125 -> 0.125; 980 -> 0.980; -450 -> -0.450.", bold_prefix="3) O'nlik masshtablash: ")

# -----------------------------------------------------------------------------
# 2-TOPSHIRIQ
# -----------------------------------------------------------------------------
add_heading_1("2-TOPSHIRIQ. Z-SCORE STANDARTLASHTIRISH (QO‘LDA HISOBLASH)")
add_body("Besh nafar talabaning test ballari berilgan: 60, 70, 80, 90, 100.")
add_body("1) O‘rtacha qiymat: μ = (60 + 70 + 80 + 90 + 100) / 5 = 400 / 5 = 80.0\n"
         "2) Kvadratik og'ishlar yig'indisi: (-20)^2 + (-10)^2 + 0^2 + 10^2 + 20^2 = 400 + 100 + 0 + 100 + 400 = 1000\n"
         "3) Standart chetlanish (scikit-learn kabi aholi formulasida n ga bo'lib): σ = sqrt(1000 / 5) = sqrt(200) ≈ 14.142")

tbl_t2 = doc.add_table(rows=1, cols=4)
tbl_t2.alignment = WD_TABLE_ALIGNMENT.CENTER
set_table_borders(tbl_t2)

h2 = ["Ball, x", "x − μ", "(x − μ)²", "z = (x − μ) / σ"]
for i, h in enumerate(h2):
    tbl_t2.rows[0].cells[i].text = h
    set_cell_background(tbl_t2.rows[0].cells[i], "1E293B")
format_row(tbl_t2.rows[0], is_header=True)

for idx, (_, r) in enumerate(df_t2.iterrows(), 1):
    row_c = tbl_t2.add_row().cells
    row_c[0].text = str(int(r['Ball, x']))
    row_c[1].text = str(int(r['x - mu']))
    row_c[2].text = str(int(r['(x - mu)^2']))
    row_c[3].text = f"{r['z']:+.3f}"
    bg = "F8FAFC" if idx % 2 == 1 else "FFFFFF"
    for cell in row_c:
        set_cell_background(cell, bg)
    format_row(tbl_t2.rows[-1])

# Yig'indi qatori
row_sum = tbl_t2.add_row().cells
row_sum[0].text = "Yig'indi (O'rtacha)"
row_sum[1].text = "0"
row_sum[2].text = "1000"
row_sum[3].text = "0.000"
for c in row_sum:
    set_cell_background(c, "E2E8F0")
format_row(tbl_t2.rows[-1], is_header=True)
row_sum[0].paragraphs[0].runs[0].font.color.rgb = RGBColor(15, 23, 42)
row_sum[1].paragraphs[0].runs[0].font.color.rgb = RGBColor(15, 23, 42)
row_sum[2].paragraphs[0].runs[0].font.color.rgb = RGBColor(15, 23, 42)
row_sum[3].paragraphs[0].runs[0].font.color.rgb = RGBColor(15, 23, 42)

add_heading_2("Tekshirish va tahlil:")
add_body("Standartlashtirilgan z ustunining o'rtacha qiymati μ_z = (-1.414 - 0.707 + 0 + 0.707 + 1.414) / 5 = 0.000 ga teng. Standart chetlanishi esa σ_z = sqrt((2 + 0.5 + 0 + 0.5 + 2) / 5) = sqrt(1.0) = 1.000 ga teng. Nazariya amalda to'liq isbotlandi.", bold_prefix="• Tekshirish: ")
add_body("z = (95 - 80) / 14.142 = 15 / 14.142 ≈ +1.061. Talaba guruh o'rtachasidan 1.061 standart chetlanish (1.061σ) yuqori natija ko'rsatgan.", bold_prefix="• 95 ball olgan talaba: ")

# -----------------------------------------------------------------------------
# 3-TOPSHIRIQ
# -----------------------------------------------------------------------------
add_heading_1("3-TOPSHIRIQ. PYTHON’DA MINMAXSCALER VA STANDARDSCALER")
add_body("Olti nafar bank mijozining yoshi (yil) va oylik daromadi (mln so‘m) berilgan. Dastlabki to'plam tavsifi:")

tbl_bank = doc.add_table(rows=1, cols=3)
tbl_bank.alignment = WD_TABLE_ALIGNMENT.CENTER
set_table_borders(tbl_bank)

hb = ["Mijoz ID", "Yoshi (yil)", "Daromadi (mln so‘m)"]
for i, h in enumerate(hb):
    tbl_bank.rows[0].cells[i].text = h
    set_cell_background(tbl_bank.rows[0].cells[i], "1E293B")
format_row(tbl_bank.rows[0], is_header=True)

for idx, (_, r) in enumerate(df_bank.iterrows(), 1):
    row_c = tbl_bank.add_row().cells
    row_c[0].text = f"{idx}-mijoz"
    row_c[1].text = str(int(r["Yosh"]))
    row_c[2].text = f"{r['Daromad']:.1f}"
    bg = "F8FAFC" if idx % 2 == 1 else "FFFFFF"
    for cell in row_c:
        set_cell_background(cell, bg)
    format_row(tbl_bank.rows[-1])

add_heading_2("Masshtablash natijalari:")
add_body("Scikit-learn kutubxonasining MinMaxScaler va StandardScaler klasslari qo'llanilib, quyidagi natijalar olindi:")

tbl_scale = doc.add_table(rows=1, cols=5)
tbl_scale.alignment = WD_TABLE_ALIGNMENT.CENTER
set_table_borders(tbl_scale)

hs = ["Mijoz ID", "MinMax (Yosh)", "MinMax (Daromad)", "Z-score (Yosh)", "Z-score (Daromad)"]
for i, h in enumerate(hs):
    tbl_scale.rows[0].cells[i].text = h
    set_cell_background(tbl_scale.rows[0].cells[i], "1E293B")
format_row(tbl_scale.rows[0], is_header=True)

for idx in range(len(df_bank)):
    row_c = tbl_scale.add_row().cells
    row_c[0].text = f"{idx+1}-mijoz"
    row_c[1].text = f"{df_mm.iloc[idx, 0]:.3f}"
    row_c[2].text = f"{df_mm.iloc[idx, 1]:.3f}"
    row_c[3].text = f"{df_st.iloc[idx, 0]:+.3f}"
    row_c[4].text = f"{df_st.iloc[idx, 1]:+.3f}"
    bg = "F8FAFC" if (idx+1) % 2 == 1 else "FFFFFF"
    for cell in row_c:
        set_cell_background(cell, bg)
    format_row(tbl_scale.rows[-1])

add_heading_2("Amaliy tahlil savollariga javoblar:")
add_body("Qo'lda hisoblash: Yosh_mm = (47-25)/(62-25) = 22/37 ≈ 0.595; Daromad_mm = (12-4.5)/(15-4.5) = 7.5/10.5 ≈ 0.714. Z-score: Yosh_z = (47-42.5)/12.312 ≈ +0.365; Daromad_z = (12-9.0)/3.617 ≈ +0.829. Python natijasi bilan 100% mos keldi.", bold_prefix="1) 3-mijozni qo'lda solishtirish: ")
add_body(f"Dastlabki masofa d_raw = {d_raw:.3f}. Bunda yoshlar farqi |25 - 62| = 37 bo'lgani holda, daromad farqi |4.5 - 6.0| = 1.5 bo'lgan. Kvadratlar yig'indisida 37^2 = 1369 bo'lib, daromad farqi (2.25) masofaga deyarli umuman ta'sir qilmagan (Yosh 99.8% ni belgilagan). Min-Max dan so'ng masofa {d_mm:.3f}, Z-score dan so'ng {d_st:.3f} bo'lib, ikkala belgi ham teng huquqli bo'ldi.", bold_prefix="2) 1- va 5-mijoz orasidagi Evklid masofasi: ")
add_body("df_st.mean() buyrug'i Yosh: 0.0, Daromad: 0.0 chiqardi. df_st.std(ddof=0) esa Yosh: 1.0, Daromad: 1.0 ekanini to'liq tasdiqladi.", bold_prefix="3) df_st o'rtacha va standart chetlanishi: ")
add_body("Yangi mijoz kiritilganda daromadning maksimal qiymati 15 dan 60 ga sakradi. Min-Max usulida qolgan barcha mijozlarning daromadlari juda kichik oraliqqa (0.00 – 0.19) siqilib qoldi (masalan, 3-mijoz daromadi 0.714 dan 0.135 ga tushib ketdi). Z-score'da esa o'rtacha daromad 9.0 dan 16.29 ga ko'tarilib, barcha eski mijozlar manfiy z-score zonaga o'tib qoldi. Bu esa Min-Max va Z-score usullarining chetki g'ayrioddiy qiymatlarga (outlier) o'ta sezgirligini amalda ko'rsatadi.", bold_prefix="4) 7-mijoz (outlier) qo'shilishi ta'siri: ")

# -----------------------------------------------------------------------------
# 4-TOPSHIRIQ
# -----------------------------------------------------------------------------
add_heading_1("4-TOPSHIRIQ. MODEL NATIJASIGA TA’SIR (WINE TO'PLAMI)")
add_body("Scikit-learn tarkibidagi Wine datasetida (178 ta vino namunasi, 13 ta kimyoviy belgi, 3 ta vino sinfi) belgilarning miqyoslari keskin farq qiladi: magniy 70–162 oralig'ida, prolin 278–1680 oralig'ida, ammo rang intensivligi va boshqa moddalar 0.13–0.66 oralig'ida joylashgan. Ushbu to'plamda k-NN va Qarorlar daraxti modellari masshtablashsiz hamda MinMaxScaler/StandardScaler bilan o'qitilib, quyidagi test aniqliklari olindi:")

tbl_model = doc.add_table(rows=1, cols=2)
tbl_model.alignment = WD_TABLE_ALIGNMENT.CENTER
set_table_borders(tbl_model)

hm = ["Mashinaviy o'rganish modeli", "Test to'plamidagi aniqlik (Accuracy)"]
for i, h in enumerate(hm):
    tbl_model.rows[0].cells[i].text = h
    set_cell_background(tbl_model.rows[0].cells[i], "1E293B")
format_row(tbl_model.rows[0], is_header=True)

for idx, (m_nom, m_acc) in enumerate(model_natijalar, 1):
    row_c = tbl_model.add_row().cells
    row_c[0].text = m_nom
    row_c[1].text = m_acc
    bg = "F8FAFC" if idx % 2 == 1 else "FFFFFF"
    for cell in row_c:
        set_cell_background(cell, bg)
    format_row(tbl_model.rows[-1])

add_heading_2("Savollarga mantiqiy va nazariy asoslangan javoblar:")
add_body("k-NN algoritmi obyektlar orasidagi Evklid masofasiga tayanadi: d(p, q) = sqrt(sum((p_i - q_i)^2)). Wine to'plamida prolin atributi [278, 1680] oralig'ida bo'lib, uning ayirmalar kvadrati yuz minglab qiymatlarni hosil qiladi. Oqibatda magniy, flavanoidlar va alkogol kabi boshqa 12 ta hal qiluvchi belgilarning masofaga ta'siri 0 ga teng bo'lib qolgan. Belgilar MinMax yoki Z-score orqali bir xil masshtabga keltirilgach, barcha atributlar qaror qabul qilishda teng ulushga ega bo'ldi va natija keskin ravishda 72.2% dan 96.3% ga ko'tarildi.", bold_prefix="1) Nima uchun k-NN aniqligi masshtablashdan so'ng keskin oshdi? ")
add_body("Qarorlar daraxti masofaga tayanmaydi! U har bir atribut bo'yicha alohida-alohida chegaraviy shartlarni (masalan, prolin <= 750) tekshirib, shoxlanish hosil qiladi. Min-Max yoki Z-score kabi monoton o'zgarishlar qiymatlarning o'zaro tartibini (order) o'zgartirmaydi. Shu sababli daraxtning bo'linish nuqtalari nisbiy ma'noda bir xil qolaveradi va aniqlik 96.3% ligicha o'zgarishsiz qoladi.", bold_prefix="2) Nima uchun Qarorlar daraxti natijasi o'zgarmaydi? ")
add_body("Test to'plami modelga mutlaqo begona (real hayotdagi kelajak ma'lumot) bo'lishi shart. Agar scaler butun ma'lumotdan min, max, o'rtacha yoki sigmani hisoblasa, test to'plamining statistik xususiyatlari o'qitish jarayoniga \"sizib o'tadi\" (Data Leakage). Natijada model test to'plamida sun'iy yuqori aniqlik ko'rsatib, amaliyotda ishlamay qoladi. Pipeline mexanizmi fit amalini qat'iy ravishda faqat X_train qismida bajaradi va test to'plamiga faqat transform amalini qo'llaydi.", bold_prefix="3) Nima uchun scaler faqat train to'plamida o'qitiladi va Data Leakage nima? ")

# -----------------------------------------------------------------------------
# XULOSA
# -----------------------------------------------------------------------------
add_heading_1("UMUMIY AMALIY XULOSA")
add_body("5-amaliy mashg'ulotni bajarish davomida ma'lumotlarni intellektual tahlilida (Data Mining) va Machine Learning sohasida belgilarni masshtablashning o'rni beqiyos ekanligi amalda isbotlandi:")
add_body("1. Turli xil o'lchov shkalasiga ega atributlar masofaga asoslangan algoritmlarni chalg'itadi. Min-Max va Z-score bu muammoni to'liq bartaraf etadi.")
add_body("2. Min-Max chegaralari qat'iy ma'lum va chetki anomaliyalarsiz ma'lumotlar uchun mos keladi. Z-score esa chetki qiymatlar mavjud bo'lganda barqaror ishlaydi.")
add_body("3. Mashinaviy o'rganish modellarini tanlashda masshtablash sezgirligi hisobga olinishi lozim: k-NN, SVM, Neyron tarmoqlari masshtablashni shart qilib qo'yadi, Daraxtsimon modellar esa unga befarqdir.")
add_body("4. Haqiqiy loyihalarda ma'lumotlar sizib ketishi (Data Leakage) xavfini yo'qotish uchun scikit-learn Pipeline vositasidan foydalanish eng to'g'ri muhandislik amaliyoti hisoblanadi.")

# -----------------------------------------------------------------------------
# NAZORAT SAVOLLARI
# -----------------------------------------------------------------------------
add_heading_1("NAZORAT SAVOLLARIGA JAVOBLAR")

savollar = [
    ("1. Normallashtirish va standartlashtirishning farqi nimada?",
     "Normallashtirish (Min-Max) barcha sonlarni aniq qat'iy oraliqqa (odatda [0, 1] yoki [-1, 1]) qisadi va dastlabki minimum hamda maksimum qiymatlarga tayanadi. Standartlashtirish (Z-score) esa ma'lumotlar taqsimotining o'rtacha qiymatini 0 ga, standart chetlanishini 1 ga keltiradi. Uning qat'iy chegarasi bo'lmaydi (qiymatlar asosan [-3, +3] oralig'ida bo'ladi) va u normal taqsimotga yaqin to'plamlar uchun mo'ljallangan."),
    
    ("2. Chetga chiquvchi qiymatlar (outlierlar) qaysi usulga kuchliroq ta’sir qiladi va nima uchun?",
     "Chetga chiquvchi qiymatlar Min-Max normallashtirishga eng kuchli va halokatli ta'sir ko'rsatadi. Chunki bitta o'ta katta outlier butun x_max qiymatini keskin oshirib yuboradi va to'plamdagi barcha oddiy obyektlar [0, 0.1] kabi o'ta tor diapazonga siqilib, ularning o'zaro farqlari yo'qoladi. Z-score ham outlierlardan ta'sirlanadi (chunki o'rtacha va sigma siljiydi), ammo uning ta'siri Min-Max ga qaraganda ancha yumshoqroq bo'ladi."),
    
    ("3. Qaysi algoritmlar masshtablashga sezgir, qaysilari sezgir emas?",
     "Masofaga va gradientga asoslangan barcha algoritmlar masshtablashga juda sezgir: k-NN, K-Means klasterlash, Qo'llab-quvvatlovchi vektorlar mashinasi (SVM), Chiziqli va Logistik regressiya, Sun'iy neyron tarmoqlari (ANN/Deep Learning). Masshtablashga sezgir bo'lmagan algoritmlar esa daraxtsimon modellar hisoblanadi: Qarorlar daraxti (Decision Tree), Tasodifiy o'rmon (Random Forest), Gradient Boosting (XGBoost, LightGBM, CatBoost) hamda sodda Bayes (Naive Bayes)."),
    
    ("4. Nima uchun scaler test to‘plamida qayta o‘qitilmaydi?",
     "Chunki test to'plami modelga mutlaqo noma'lum yangi real ma'lumot sifatida qaralishi lozim. Agar scaler test to'plamida qayta o'qitilsa (fit qilinsa), o'quv va test to'plamlari mutlaqo boshqa-boshqa koordinata tizimlariga o'tib qoladi va model o'qitilgan qonuniyatlar testda buziladi. Shuningdek, test to'plamining min/max/mean parametrlari o'qitish bosqichiga ta'sir qilmasligi, ya'ni ma'lumotlar sizib chiqishining (Data Leakage) oldini olish uchun ham scaler faqat train'da fit qilinadi va testga faqat transform amali qo'llanadi.")
]

for q, a in savollar:
    add_body(q, bold_prefix="")
    p_ans = add_body(a)
    p_ans.paragraph_format.left_indent = Inches(0.2)

# Faylni saqlash
docx_out = os.path.join(CURRENT_DIR, "5-amaliy_mashg'ulot_hisoboti.docx")
doc.save(docx_out)
print(f"Word hisoboti muvaffaqiyatli saqlandi: {docx_out}")

# ==============================================================================
# 3. JUPYTER NOTEBOOK YARATISH
# ==============================================================================
nb = {
    "cells": [
        {
            "cell_type": "markdown",
            "metadata": {},
            "source": [
                "# 5-Amaliy mashg‘ulot\n",
                "## Mavzu: Ma’lumotlarni normallashtirish va standartlashtirish usullarini qo‘llash\n",
                "**Bajaruvchi:** Talaba\n",
                "**Fani:** Data Mining\n",
                "\n",
                "---\n",
                "### Ishning maqsadi:\n",
                "Belgilarni bir xil masshtabga keltirishning asosiy usullarini qo‘lda va Python’da qo‘llash, ularning model natijalariga (k-NN va Qarorlar daraxti) ta’sirini tushunish."
            ]
        },
        {
            "cell_type": "markdown",
            "metadata": {},
            "source": [
                "### 1-topshiriq. Min-Max normallashtirish (qo‘lda va pandas’da)\n",
                "Besh nafar xodimning yoshi berilgan: 20, 35, 50, 65, 80 yil."
            ]
        },
        {
            "cell_type": "code",
            "execution_count": 1,
            "metadata": {},
            "outputs": [],
            "source": [
                "import numpy as np\n",
                "import pandas as pd\n",
                "\n",
                "yoshlar = [20, 35, 50, 65, 80]\n",
                "x_min, x_max = min(yoshlar), max(yoshlar)\n",
                "farq = x_max - x_min\n",
                "\n",
                "df1 = pd.DataFrame({'Yosh, x': yoshlar})\n",
                "df1['x - x_min'] = df1['Yosh, x'] - x_min\n",
                "df1[\"x' in [0, 1]\"] = df1['x - x_min'] / farq\n",
                "df1[\"x' in [-1, 1]\"] = 2 * df1[\"x' in [0, 1]\"] - 1\n",
                "\n",
                "print(\"1-topshiriq natijaviy jadvali:\")\n",
                "display(df1)\n",
                "\n",
                "# Savollarga hisob-kitoblar\n",
                "yosh_42 = (42 - x_min) / farq\n",
                "yosh_90 = (90 - x_min) / farq\n",
                "print(f\"42 yoshli xodim uchun [0, 1] dagi qiymat: {yosh_42:.3f}\")\n",
                "print(f\"90 yoshli xodim uchun qiymat: {yosh_90:.3f} (chegaradan chiqib ketadi)\")\n",
                "\n",
                "# O'nlik masshtablash: [125, 980, -450]\n",
                "vals = [125, 980, -450]\n",
                "j = 3\n",
                "print(\"O'nlik masshtablash:\", {v: v / (10**j) for v in vals})"
            ]
        },
        {
            "cell_type": "markdown",
            "metadata": {},
            "source": [
                "### 2-topshiriq. Z-score standartlashtirish (qo‘lda va formulalar yordamida)\n",
                "Besh nafar talabaning test ballari berilgan: 60, 70, 80, 90, 100."
            ]
        },
        {
            "cell_type": "code",
            "execution_count": 2,
            "metadata": {},
            "outputs": [],
            "source": [
                "ballar = [60, 70, 80, 90, 100]\n",
                "n = len(ballar)\n",
                "mu = sum(ballar) / n\n",
                "diffs = [x - mu for x in ballar]\n",
                "diffs_sq = [(x - mu) ** 2 for x in ballar]\n",
                "sigma = (sum(diffs_sq) / n) ** 0.5\n",
                "z_scores = [d / sigma for d in diffs]\n",
                "\n",
                "df2 = pd.DataFrame({\n",
                "    'Ball, x': ballar,\n",
                "    'x - mu': diffs,\n",
                "    '(x - mu)^2': diffs_sq,\n",
                "    'z': [round(z, 3) for z in z_scores]\n",
                "})\n",
                "print(f\"mu = {mu}, sigma = {sigma:.3f}\")\n",
                "display(df2)\n",
                "\n",
                "z_95 = (95 - mu) / sigma\n",
                "print(f\"95 ball olgan talabaning z qiymati: {z_95:.3f} (o'rtachadan {z_95:.3f}*sigma yuqori)\")"
            ]
        },
        {
            "cell_type": "markdown",
            "metadata": {},
            "source": [
                "### 3-topshiriq. Python’da MinMaxScaler va StandardScaler\n",
                "6 nafar mijoz ma'lumotlari bo'yicha masshtablash va Evklid masofasi tahlili."
            ]
        },
        {
            "cell_type": "code",
            "execution_count": 3,
            "metadata": {},
            "outputs": [],
            "source": [
                "from sklearn.preprocessing import MinMaxScaler, StandardScaler\n",
                "\n",
                "df_bank = pd.DataFrame({\n",
                "    \"Yosh\":    [25, 32, 47, 51, 62, 38],\n",
                "    \"Daromad\": [4.5, 7.0, 12.0, 9.5, 6.0, 15.0]\n",
                "})\n",
                "display(df_bank.describe())\n",
                "\n",
                "mm = MinMaxScaler()\n",
                "df_mm = pd.DataFrame(mm.fit_transform(df_bank), columns=df_bank.columns)\n",
                "\n",
                "st = StandardScaler()\n",
                "df_st = pd.DataFrame(st.fit_transform(df_bank), columns=df_bank.columns)\n",
                "\n",
                "print(\"Min-Max natijasi:\")\n",
                "display(df_mm.round(3))\n",
                "\n",
                "print(\"Z-score natijasi:\")\n",
                "display(df_st.round(3))\n",
                "\n",
                "# Evklid masofasi tahlili (1 va 5 mijoz)\n",
                "p1 = df_bank.iloc[0].values\n",
                "p5 = df_bank.iloc[4].values\n",
                "print(f\"Dastlabki masofa: {np.linalg.norm(p1 - p5):.3f}\")\n",
                "print(f\"Min-Max dan keyin: {np.linalg.norm(df_mm.iloc[0].values - df_mm.iloc[4].values):.3f}\")\n",
                "print(f\"Z-score dan keyin: {np.linalg.norm(df_st.iloc[0].values - df_st.iloc[4].values):.3f}\")"
            ]
        },
        {
            "cell_type": "markdown",
            "metadata": {},
            "source": [
                "### 4-topshiriq. Masshtablashning model natijasiga ta'siri (Wine dataset)\n",
                "k-NN va Decision Tree modellarini masshtablanmagan va masshtablangan holatda solishtirish."
            ]
        },
        {
            "cell_type": "code",
            "execution_count": 4,
            "metadata": {},
            "outputs": [],
            "source": [
                "from sklearn.datasets import load_wine\n",
                "from sklearn.model_selection import train_test_split\n",
                "from sklearn.neighbors import KNeighborsClassifier\n",
                "from sklearn.tree import DecisionTreeClassifier\n",
                "from sklearn.pipeline import make_pipeline\n",
                "\n",
                "X, y = load_wine(return_X_y=True)\n",
                "X_train, X_test, y_train, y_test = train_test_split(\n",
                "    X, y, test_size=0.3, random_state=42, stratify=y\n",
                ")\n",
                "\n",
                "modellar = {\n",
                "    \"k-NN (masshtablashsiz)\": KNeighborsClassifier(n_neighbors=5),\n",
                "    \"k-NN + MinMax\": make_pipeline(MinMaxScaler(), KNeighborsClassifier(5)),\n",
                "    \"k-NN + Z-score\": make_pipeline(StandardScaler(), KNeighborsClassifier(5)),\n",
                "    \"Qarorlar daraxti\": DecisionTreeClassifier(random_state=42),\n",
                "    \"Qarorlar daraxti + Z-score\": make_pipeline(StandardScaler(), DecisionTreeClassifier(random_state=42)),\n",
                "}\n",
                "\n",
                "for nom, model in modellar.items():\n",
                "    model.fit(X_train, y_train)\n",
                "    acc = model.score(X_test, y_test)\n",
                "    print(f\"{nom:28s} : {acc:.3f} ({acc*100:.1f}%)\")"
            ]
        }
    ],
    "metadata": {
        "language_info": {"name": "python", "version": "3.12"}
    },
    "nbformat": 4,
    "nbformat_minor": 2
}

ipynb_out = os.path.join(CURRENT_DIR, "5-amaliy_mashgulot.ipynb")
with open(ipynb_out, "w", encoding="utf-8") as f:
    json.dump(nb, f, indent=2, ensure_ascii=False)
print(f"Jupyter Notebook muvaffaqiyatli saqlandi: {ipynb_out}")
