# -*- coding: utf-8 -*-
"""
5-amaliy mashg'ulot: Ma’lumotlarni normallashtirish va standartlashtirish usullarini qo‘llash
Ushbu skript topshiriqdagi barcha 4 ta vazifani, hisob-kitoblarni, scikit-learn tahlillarini va nazorat savollarini bajaradi.
"""

import os
import numpy as np
import pandas as pd
from sklearn.preprocessing import MinMaxScaler, StandardScaler
from sklearn.datasets import load_wine
from sklearn.model_selection import train_test_split
from sklearn.neighbors import KNeighborsClassifier
from sklearn.tree import DecisionTreeClassifier
from sklearn.pipeline import make_pipeline

def main():
    print("=" * 80)
    print("  5-AMALIY MASHG'ULOT: NORMALLASHTIRISH VA STANDARTLASHTIRISH TAHLILI")
    print("=" * 80)

    # -------------------------------------------------------------------------
    # 1-TOPSHIRIQ. MIN-MAX NORMALLASHTIRISH (QO'LDA VA PYTHON'DA)
    # -------------------------------------------------------------------------
    print("\n" + "=" * 80)
    print("1-TOPSHIRIQ. MIN-MAX NORMALLASHTIRISH (QO'LDA HISOBLASH)")
    print("=" * 80)

    yoshlar = [20, 35, 50, 65, 80]
    x_min = min(yoshlar)  # 20
    x_max = max(yoshlar)  # 80
    farq = x_max - x_min  # 60

    print(f"Xodimlar yoshi: {yoshlar}")
    print(f"x_min = {x_min}, x_max = {x_max}, Diapazon = {farq}")

    df_task1 = pd.DataFrame({'Yosh, x': yoshlar})
    df_task1['x - x_min'] = df_task1['Yosh, x'] - x_min
    df_task1["x' in [0, 1]"] = df_task1['x - x_min'] / farq
    df_task1["x' in [-1, 1]"] = 2 * df_task1["x' in [0, 1]"] - 1

    print("\nTo'ldirilgan jadval:")
    print(df_task1.to_string(index=False))

    yosh_42 = (42 - x_min) / farq
    yosh_90 = (90 - x_min) / farq
    print(f"\n• Yangi xodim yoshi 42 bo'lsa: x' = (42 - 20) / 60 = 22 / 60 = {yosh_42:.3f}")
    print(f"• 90 yoshli qiymat uchun: x' = (90 - 20) / 60 = 70 / 60 = {yosh_90:.3f}")
    print("  Xulosa: Natija 1 dan oshib ketadi. Min-Max faqat [20, 80] oralig'i bilan chegaralangan.")

    # O'nlik masshtablash
    vals_dec = [125, 980, -450]
    max_abs = max(abs(v) for v in vals_dec)  # 980
    j = 3  # chunki 10^3 = 1000 > 980
    dec_scaled = [v / (10 ** j) for v in vals_dec]
    print(f"\n• O'nlik masshtablash: {vals_dec}")
    print(f"  Eng katta modul = {max_abs} < 10^{j}  =>  j = {j}")
    for v, s in zip(vals_dec, dec_scaled):
        print(f"  {v:5d}  =>  {v} / 10^{j} = {s:.3f}")

    # -------------------------------------------------------------------------
    # 2-TOPSHIRIQ. Z-SCORE STANDARTLASHTIRISH (QO'LDA HISOBLASH)
    # -------------------------------------------------------------------------
    print("\n" + "=" * 80)
    print("2-TOPSHIRIQ. Z-SCORE STANDARTLASHTIRISH (QO'LDA HISOBLASH)")
    print("=" * 80)

    ballar = [60, 70, 80, 90, 100]
    n = len(ballar)
    mu = sum(ballar) / n  # 80.0
    diffs = [x - mu for x in ballar]
    diffs_sq = [(x - mu) ** 2 for x in ballar]
    sigma = (sum(diffs_sq) / n) ** 0.5  # sqrt(200) = 14.1421356...
    z_scores = [d / sigma for d in diffs]

    df_task2 = pd.DataFrame({
        'Ball, x': ballar,
        'x - mu': diffs,
        '(x - mu)^2': diffs_sq,
        'z': [round(z, 3) for z in z_scores]
    })
    print(f"Talabalar ballari: {ballar}")
    print(f"O'rtacha qiymat mu = {mu}")
    print(f"Standart chetlanish sigma = sqrt({sum(diffs_sq)} / {n}) = sqrt(200) = {sigma:.3f}\n")
    print(df_task2.to_string(index=False))

    mean_z = sum(z_scores) / n
    sigma_z = (sum((z - mean_z) ** 2 for z in z_scores) / n) ** 0.5
    print(f"\n• Tekshirish: Standartlashtirilgan qiymatlarning o'rtachasi = {mean_z:.4f}, sigma = {sigma_z:.4f}")

    z_95 = (95 - mu) / sigma
    print(f"• 95 ball olgan talabaning z qiymati: z = (95 - 80) / {sigma:.3f} = {z_95:.3f}")
    print(f"  U guruh o'rtachasidan {z_95:.3f} * sigma yuqori.")

    # -------------------------------------------------------------------------
    # 3-TOPSHIRIQ. PYTHON'DA MINMAXSCALER VA STANDARDSCALER
    # -------------------------------------------------------------------------
    print("\n" + "=" * 80)
    print("3-TOPSHIRIQ. PYTHON'DA MINMAXSCALER VA STANDARDSCALER")
    print("=" * 80)

    df_bank = pd.DataFrame({
        "Yosh":    [25, 32, 47, 51, 62, 38],
        "Daromad": [4.5, 7.0, 12.0, 9.5, 6.0, 15.0]
    })
    print("Dastlabki mijozlar ma'lumoti:")
    print(df_bank)
    print("\nDescriptive statistics:")
    print(df_bank.describe().round(3))

    mm = MinMaxScaler()
    df_mm = pd.DataFrame(mm.fit_transform(df_bank), columns=df_bank.columns)
    print("\n1) Min-Max normallashtirish natijasi:")
    print(df_mm.round(3))

    st = StandardScaler()
    df_st = pd.DataFrame(st.fit_transform(df_bank), columns=df_bank.columns)
    print(f"\n2) Z-score parametrlari: O'rtacha = {st.mean_.round(3)}, Sigma = {st.scale_.round(3)}")
    print("Z-score standartlashtirish natijasi:")
    print(df_st.round(3))

    # 3-mijozni qo'lda tekshirish
    yosh_3_mm = (47 - 25) / (62 - 25)
    dar_3_mm = (12.0 - 4.5) / (15.0 - 4.5)
    yosh_3_st = (47 - st.mean_[0]) / st.scale_[0]
    dar_3_st = (12.0 - st.mean_[1]) / st.scale_[1]
    print(f"\n• 3-mijoz (Yosh=47, Daromad=12.0) qo'lda solishtirish:")
    print(f"  MinMax:  Yosh = {yosh_3_mm:.3f}, Daromad = {dar_3_mm:.3f}  (Python: {df_mm.iloc[2, 0]:.3f}, {df_mm.iloc[2, 1]:.3f})")
    print(f"  Z-score: Yosh = {yosh_3_st:.3f}, Daromad = {dar_3_st:.3f}  (Python: {df_st.iloc[2, 0]:.3f}, {df_st.iloc[2, 1]:.3f})")

    # 1 va 5 mijoz masofasi
    p1_raw = df_bank.iloc[0].values
    p5_raw = df_bank.iloc[4].values
    d_raw = np.linalg.norm(p1_raw - p5_raw)
    d_mm = np.linalg.norm(df_mm.iloc[0].values - df_mm.iloc[4].values)
    d_st = np.linalg.norm(df_st.iloc[0].values - df_st.iloc[4].values)
    print(f"\n• 1- va 5-mijoz orasidagi Evklid masofasi:")
    print(f"  Dastlabki masofa: {d_raw:.3f} (Yosh farqi = {abs(p1_raw[0] - p5_raw[0])}, Daromad farqi = {abs(p1_raw[1] - p5_raw[1])})")
    print("  Izoh: Yosh farqi (37) masofani to'liq egallab olgan, daromad (1.5) sezilmagan.")
    print(f"  MinMax dan keyingi masofa:  {d_mm:.3f}")
    print(f"  Z-score dan keyingi masofa: {d_st:.3f}")

    # Outlier (7-mijoz)
    df7 = pd.concat([df_bank, pd.DataFrame({'Yosh': [30], 'Daromad': [60.0]})], ignore_index=True)
    df7_mm = pd.DataFrame(MinMaxScaler().fit_transform(df7), columns=df7.columns)
    df7_st = pd.DataFrame(StandardScaler().fit_transform(df7), columns=df7.columns)
    print(f"\n• 7-mijoz qo'shilgandan so'ng (Yosh=30, Daromad=60.0):")
    print("  MinMax natijasi:")
    print(df7_mm.round(3))
    print("  Z-score natijasi:")
    print(df7_st.round(3))

    # -------------------------------------------------------------------------
    # 4-TOPSHIRIQ. MODEL NATIJASIGA TA'SIRI (WINE TO'PLAMI)
    # -------------------------------------------------------------------------
    print("\n" + "=" * 80)
    print("4-TOPSHIRIQ. MODEL NATIJASIGA TA'SIRI (WINE TO'PLAMI)")
    print("=" * 80)

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

    natijalar = []
    for nom, model in modellar.items():
        model.fit(X_train, y_train)
        acc = model.score(X_test, y_test)
        natijalar.append({"Model": nom, "Test aniqligi (Accuracy)": f"{acc:.3f} ({acc*100:.1f}%)"})

    df_models = pd.DataFrame(natijalar)
    print(df_models.to_string(index=False))

    print("\n" + "=" * 80)
    print("  5-AMALIY MASHG'ULOT TAHLILI YAKUNLANDI.")
    print("=" * 80)

if __name__ == "__main__":
    main()
