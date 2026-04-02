"""
MCDM Engine Doğrulama Scripti
Elle hesaplama ile engine çıktısını karşılaştırır.

Test verisi:
         C1    C2    C3
A1        4     3     5    (C1=max, C2=min, C3=max)
A2        3     6     2
A3        5     2     4
A4        2     8     3

Eşit ağırlıklar: w = [1/3, 1/3, 1/3]
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))

import numpy as np
import pandas as pd
import mcdm_engine as me

EPS = 1e-12
np.set_printoptions(precision=6, suppress=True)
pd.set_option("display.float_format", "{:.6f}".format)
pd.set_option("display.width", 120)

# ── Test verisi ──────────────────────────────────────────────────────────────
data = pd.DataFrame(
    {"C1": [4.0, 3.0, 5.0, 2.0],
     "C2": [3.0, 6.0, 2.0, 8.0],
     "C3": [5.0, 2.0, 4.0, 3.0]},
    index=["A1", "A2", "A3", "A4"],
)
crit_types = {"C1": "max", "C2": "min", "C3": "max"}
weights    = {"C1": 1/3, "C2": 1/3, "C3": 1/3}

SEP  = "=" * 70
SEP2 = "-" * 70

def section(title):
    print(f"\n{SEP}\n{title}\n{SEP}")

def ok_fail(label, cond, got, expected):
    tag = "✅ OK" if cond else "❌ HATA"
    print(f"  {tag}  {label:30s}  engine={got:.6f}  el={expected:.6f}")

# ─────────────────────────────────────────────────────────────────────────────
# 1. TOPSIS
# ─────────────────────────────────────────────────────────────────────────────
section("1. TOPSIS")

x = data.to_numpy(float)

# Vektör normalizasyonu
col_norms = np.sqrt((x**2).sum(axis=0))
r = x / col_norms
print("Normalize matris (elle):")
print(pd.DataFrame(r, index=data.index, columns=data.columns))

# Ağırlıklı matris
w = np.array([1/3, 1/3, 1/3])
v = r * w

# PIS / NIS
pis = np.array([v[:,0].max(), v[:,1].min(), v[:,2].max()])  # C1 max, C2 min, C3 max
nis = np.array([v[:,0].min(), v[:,1].max(), v[:,2].min()])

dplus  = np.sqrt(((v - pis)**2).sum(axis=1))
dminus = np.sqrt(((v - nis)**2).sum(axis=1))
cc_manual = dminus / (dplus + dminus + EPS)

print("\nD+, D-, CC (elle):")
for i, alt in enumerate(data.index):
    print(f"  {alt}: D+={dplus[i]:.6f}  D-={dminus[i]:.6f}  CC={cc_manual[i]:.6f}")

# Engine
score_t, _ = me._rank_topsis(data, crit_types, weights)
print("\nTOPSIS engine skoru vs elle:")
all_ok = True
for i, alt in enumerate(data.index):
    match = abs(score_t[i] - cc_manual[i]) < 1e-6
    if not match: all_ok = False
    ok_fail(alt, match, score_t[i], cc_manual[i])

rank_manual = np.argsort(-cc_manual) + 1  # küçük rank = iyi
rank_engine = np.argsort(-score_t) + 1
print(f"\n  Elle   sıralama: {dict(zip(data.index, np.argsort(-cc_manual)+1))}")
print(f"  Engine sıralama: {dict(zip(data.index, np.argsort(-score_t)+1))}")
print(f"  TOPSIS {'✅ TAMAM' if all_ok else '❌ UYUŞMUYOR'}")

# ─────────────────────────────────────────────────────────────────────────────
# 2. VIKOR
# ─────────────────────────────────────────────────────────────────────────────
section("2. VIKOR  (v=0.5)")

x_v = data.to_numpy(float)

# Doğru VIKOR — signed denominator kullanır
best_v  = np.array([x_v[:,0].max(), x_v[:,1].min(), x_v[:,2].max()])  # ideal
worst_v = np.array([x_v[:,0].min(), x_v[:,1].max(), x_v[:,2].min()])  # anti-ideal

diff_correct = np.zeros_like(x_v)
for j in range(3):
    denom = best_v[j] - worst_v[j]          # signed: + for max, - for min
    if abs(denom) <= EPS:
        diff_correct[:, j] = 0.0
    else:
        diff_correct[:, j] = (1/3) * (best_v[j] - x_v[:, j]) / denom

s_correct = diff_correct.sum(axis=1)
r_correct = diff_correct.max(axis=1)

print("Doğru diff matrisi (elle, signed denom):")
print(pd.DataFrame(diff_correct, index=data.index, columns=data.columns))
print(f"\nS (elle): {dict(zip(data.index, s_correct.round(6)))}")
print(f"R (elle): {dict(zip(data.index, r_correct.round(6)))}")

# Kodun ürettiği diff (abs denominator — potansiyel hata)
diff_code = np.zeros_like(x_v)
for j in range(3):
    denom_abs = abs(best_v[j] - worst_v[j])
    if denom_abs <= EPS:
        diff_code[:, j] = 0.0
    else:
        diff_code[:, j] = (1/3) * (best_v[j] - x_v[:, j]) / denom_abs

s_code_manual = diff_code.sum(axis=1)
r_code_manual = diff_code.max(axis=1)

print("\nKodun yaptığı diff matrisi (elle, abs denom):")
print(pd.DataFrame(diff_code, index=data.index, columns=data.columns))
print(f"\nS (kod mantığıyla): {dict(zip(data.index, s_code_manual.round(6)))}")
print(f"R (kod mantığıyla): {dict(zip(data.index, r_code_manual.round(6)))}")

# Doğru Q
v_p = 0.5
s_star, s_minus = s_correct.min(), s_correct.max()
r_star, r_minus = r_correct.min(), r_correct.max()
q_correct = v_p*(s_correct-s_star)/(s_minus-s_star+EPS) + (1-v_p)*(r_correct-r_star)/(r_minus-r_star+EPS)

print(f"\nDoğru Q: {dict(zip(data.index, q_correct.round(6)))}")
print(f"Doğru sıralama: {dict(zip(data.index, (np.argsort(q_correct)+1)))}")

# Engine
score_v, det_v = me._rank_vikor(data, crit_types, weights, v_param=0.5)
vt = det_v["vikor_table"].set_index("Alternatif")
print(f"\nEngine S,R,Q:")
print(vt[["S","R","Q","Skor"]])
print(f"Engine sıralama: {dict(zip(data.index, (np.argsort(score_v)+1)))}")

# Karşılaştır
print(f"\n{SEP2}")
print("VIKOR karşılaştırması (elle=doğru formül, engine=kodun ürettiği):")
for alt in data.index:
    idx = list(data.index).index(alt)
    s_e = float(vt.loc[alt, "S"])
    s_m = s_correct[idx]
    match = abs(s_e - s_m) < 1e-4
    ok_fail(f"{alt} S-değeri", match, s_e, s_m)

# Min kriter C2 diff karşılaştırması
print(f"\n  C2 (min kriter) diff değerleri karşılaştırması:")
print(f"  {'Alt':4}  {'Doğru':>10}  {'Kod':>10}  {'Fark':>10}")
for i, alt in enumerate(data.index):
    corr = diff_correct[i, 1]
    code = diff_code[i, 1]
    print(f"  {alt:4}  {corr:10.6f}  {code:10.6f}  {corr-code:10.6f}")

# ─────────────────────────────────────────────────────────────────────────────
# 3. EDAS
# ─────────────────────────────────────────────────────────────────────────────
section("3. EDAS")

x_e = data.to_numpy(float)
av  = x_e.mean(axis=0)
print(f"Ortalama çözüm: C1={av[0]:.4f}, C2={av[1]:.4f}, C3={av[2]:.4f}")

pda_m = np.zeros_like(x_e)
nda_m = np.zeros_like(x_e)
for j, (cname, ctype) in enumerate(crit_types.items()):
    if ctype == "max":
        pda_m[:, j] = np.maximum(0.0, (x_e[:, j] - av[j]) / (abs(av[j]) + EPS))
        nda_m[:, j] = np.maximum(0.0, (av[j] - x_e[:, j]) / (abs(av[j]) + EPS))
    else:
        pda_m[:, j] = np.maximum(0.0, (av[j] - x_e[:, j]) / (abs(av[j]) + EPS))
        nda_m[:, j] = np.maximum(0.0, (x_e[:, j] - av[j]) / (abs(av[j]) + EPS))

wvec = np.array([1/3, 1/3, 1/3])
sp_m = (pda_m * wvec).sum(axis=1)
sn_m = (nda_m * wvec).sum(axis=1)
nsp_m = sp_m / (sp_m.max() + EPS)
nsn_m = 1.0 - sn_m / (sn_m.max() + EPS)
score_e_m = 0.5 * (nsp_m + nsn_m)

print("\nEDAS skor (elle):")
for i, alt in enumerate(data.index):
    print(f"  {alt}: SP={sp_m[i]:.6f}  SN={sn_m[i]:.6f}  NSP={nsp_m[i]:.6f}  NSN={nsn_m[i]:.6f}  Skor={score_e_m[i]:.6f}")

score_e, det_e = me._rank_edas(data, crit_types, weights)
print("\nEDAS engine vs elle:")
all_ok = True
for i, alt in enumerate(data.index):
    match = abs(score_e[i] - score_e_m[i]) < 1e-6
    if not match: all_ok = False
    ok_fail(alt, match, score_e[i], score_e_m[i])
print(f"  EDAS {'✅ TAMAM' if all_ok else '❌ UYUŞMUYOR'}")

# ─────────────────────────────────────────────────────────────────────────────
# 4. Entropy Ağırlıkları
# ─────────────────────────────────────────────────────────────────────────────
section("4. Entropy Ağırlıkları")

# Sum normalizasyonu (benefit yönü)
xmin = x.min(axis=0)  # [2, 2, 2]
xmax = x.max(axis=0)  # [5, 8, 5]

nsum = np.zeros_like(x)
# C1 (max): x - min
trans_c1 = (x[:,0] - xmin[0]).clip(min=0) + EPS
nsum[:,0] = trans_c1 / trans_c1.sum()
# C2 (min): max - x
trans_c2 = (xmax[1] - x[:,1]).clip(min=0) + EPS
nsum[:,1] = trans_c2 / trans_c2.sum()
# C3 (max): x - min
trans_c3 = (x[:,2] - xmin[2]).clip(min=0) + EPS
nsum[:,2] = trans_c3 / trans_c3.sum()

print("Sum normalizasyon matrisi (elle):")
print(pd.DataFrame(nsum, index=data.index, columns=data.columns))

m = 4
k = 1.0 / np.log(m)
p = nsum + EPS
entropy_m = -k * (p * np.log(p)).sum(axis=0)
divergence_m = 1.0 - entropy_m
weights_m = divergence_m / divergence_m.sum()

print(f"\nEntropy (elle): {dict(zip(data.columns, entropy_m.round(6)))}")
print(f"Divergence    : {dict(zip(data.columns, divergence_m.round(6)))}")
print(f"Ağırlıklar    : {dict(zip(data.columns, weights_m.round(6)))}")

w_e, det_w = me._weights_entropy(data, crit_types)
print(f"\nEngine ağırlıkları: {w_e}")

print("\nEntropy engine vs elle:")
all_ok = True
for c in data.columns:
    idx = list(data.columns).index(c)
    match = abs(w_e[c] - weights_m[idx]) < 1e-6
    if not match: all_ok = False
    ok_fail(c, match, w_e[c], weights_m[idx])
print(f"  Entropy {'✅ TAMAM' if all_ok else '❌ UYUŞMUYOR'}")

# ─────────────────────────────────────────────────────────────────────────────
# 5. SAW (basit doğrulama)
# ─────────────────────────────────────────────────────────────────────────────
section("5. SAW (Weighted Sum Model)")

# SAW engine _normalize_sum kullanır (entropy ile aynı norm).
# Elle: (x-min)/sum(x-min) for max, (max-x)/sum(max-x) for min
nsum_saw = np.zeros_like(x)
for j, (cname, ctype) in enumerate(crit_types.items()):
    if ctype == "max":
        tr = (x[:,j] - x[:,j].min()).clip(min=0) + EPS
    else:
        tr = (x[:,j].max() - x[:,j]).clip(min=0) + EPS
    nsum_saw[:,j] = tr / tr.sum()

saw_score_m = (nsum_saw * wvec).sum(axis=1)
print(f"SAW skor (elle, sum-norm): {dict(zip(data.index, saw_score_m.round(6)))}")

score_saw, _ = me._rank_saw(data, crit_types, weights)
print(f"SAW engine                : {dict(zip(data.index, score_saw.round(6)))}")
all_ok = all(abs(score_saw[i]-saw_score_m[i]) < 1e-6 for i in range(4))
print(f"  SAW {'✅ TAMAM' if all_ok else '❌ UYUŞMUYOR'}")

# ─────────────────────────────────────────────────────────────────────────────
# 6. ÖZET
# ─────────────────────────────────────────────────────────────────────────────
section("ÖZET")

topsis_ok = all(abs(score_t[i]-cc_manual[i])<1e-6 for i in range(4))
edas_ok   = all(abs(score_e[i]-score_e_m[i])<1e-6 for i in range(4))
ent_ok    = all(abs(w_e[c]-weights_m[list(data.columns).index(c)])<1e-6 for c in data.columns)
saw_ok2   = all(abs(score_saw[i]-saw_score_m[i])<1e-6 for i in range(4))

# VIKOR bug testi: min kriter C2'nin diff değeri negatif mi?
vikor_c2_negative = any(diff_code[:, 1] < -EPS)
vt_s = np.array([float(vt.loc[alt, "S"]) for alt in data.index])
vikor_s_matches_correct = all(abs(vt_s[i]-s_correct[i])<1e-4 for i in range(4))

print(f"  TOPSIS       {'✅' if topsis_ok else '❌'}")
print(f"  EDAS         {'✅' if edas_ok else '❌'}")
print(f"  Entropy      {'✅' if ent_ok else '❌'}")
print(f"  SAW          {'✅' if saw_ok2 else '❌'}")
print()
print(f"  VIKOR (min kriter diff negatif mi?): {'❌ EVET — bug var' if vikor_c2_negative else '✅ Hayır'}")
print(f"  VIKOR S değerleri doğru formülle eşleşiyor mu?: {'✅ Evet' if vikor_s_matches_correct else '❌ Hayır — fark var'}")

if vikor_c2_negative:
    print()
    print("  VIKOR BUG DETAYI:")
    print("  Kodda: denom = abs(best - worst)  →  min kriter için (min-x) / (max-min)")
    print("  Bu negatif diff değerleri üretir; S ve R hesaplamalarını bozar.")
    print("  Doğru: denom = best - worst  (işaretli)  →  her durumda 0..w aralığında")
    print()
    print("  Min kriter C2 için:")
    print(f"  {'Alt':4}  {'x':>6}  {'Kod diff':>10}  {'Doğru diff':>12}  {'Fark':>10}")
    for i, alt in enumerate(data.index):
        print(f"  {alt:4}  {x[i,1]:6.1f}  {diff_code[i,1]:10.6f}  {diff_correct[i,1]:12.6f}  {diff_code[i,1]-diff_correct[i,1]:10.6f}")

    print()
    print("  Doğru ranking : ", {k: v for k,v in zip(data.index, np.argsort(q_correct)+1)})
    score_v_arr = np.array([float(vt.loc[alt,"Q"]) for alt in data.index])
    print("  Engine ranking: ", {k: v for k,v in zip(data.index, np.argsort(score_v_arr)+1)})
    rank_match = all(
        (np.argsort(q_correct)+1)[i] == (np.argsort(score_v_arr)+1)[i]
        for i in range(4)
    )
    print(f"  Sıralama uyuşuyor mu? {'✅ Evet' if rank_match else '❌ Hayır'}")
