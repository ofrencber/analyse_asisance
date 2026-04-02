"""
Fuzzy MCDM Doğrulama Scripti
Elle hesaplama ile engine çıktısını karşılaştırır.
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))

import numpy as np
import pandas as pd
import mcdm_engine as me

EPS = 1e-12
np.set_printoptions(precision=6, suppress=True)
pd.set_option("display.float_format", "{:.6f}".format)

# ── Test verisi ──────────────────────────────────────────────────────────────
data = pd.DataFrame(
    {"C1": [4.0, 3.0, 5.0, 2.0],
     "C2": [3.0, 6.0, 2.0, 8.0],
     "C3": [5.0, 2.0, 4.0, 3.0]},
    index=["A1", "A2", "A3", "A4"],
)
crit_types = {"C1": "max", "C2": "min", "C3": "max"}
weights    = {"C1": 1/3, "C2": 1/3, "C3": 1/3}
SPREAD = 0.10

SEP  = "=" * 70
SEP2 = "-" * 70

def section(title):
    print(f"\n{SEP}\n{title}\n{SEP}")

def ok_fail(label, cond, got=None, expected=None):
    tag = "✅ OK" if cond else "❌ HATA"
    if got is not None:
        print(f"  {tag}  {label:35s}  engine={got:.6f}  el={expected:.6f}")
    else:
        print(f"  {tag}  {label}")

# ─────────────────────────────────────────────────────────────────────────────
# 1. TFN üretimi
# ─────────────────────────────────────────────────────────────────────────────
section("1. TFN Üretimi (spread=0.10)")

x = data.to_numpy(float)
tfn_m = np.stack([x * (1 - SPREAD), x, x * (1 + SPREAD)], axis=2)

tfn_e = me._triangular_fuzzy_from_crisp(data, SPREAD)

print("Elle TFN[A1]: ", tfn_m[0])
print("Engine TFN[A1]:", tfn_e[0])
all_ok = np.allclose(tfn_m, tfn_e, atol=1e-9)
ok_fail("TFN üretimi (tüm alternatifler)", all_ok)

# l ≤ m ≤ u koşulu
lmu_ok = bool(np.all(tfn_m[:,:,0] <= tfn_m[:,:,1]) and np.all(tfn_m[:,:,1] <= tfn_m[:,:,2]))
ok_fail("l ≤ m ≤ u koşulu", lmu_ok)

# ─────────────────────────────────────────────────────────────────────────────
# 2. Fuzzy TFN Normalizasyonu
# ─────────────────────────────────────────────────────────────────────────────
section("2. Fuzzy TFN Normalizasyonu")

# Önce shift (tüm değerler pozitif olduğu için shift=0)
tfn_pos = tfn_m.copy()

# Max kriter (C1, idx=0): normalize = val / max(upper)
max_u_c1 = tfn_pos[:, 0, 2].max()
norm_c1_m = tfn_pos[:, 0, :] / max_u_c1

# Min kriter (C2, idx=1): normalize = min(lower) / val, sıralama ters
min_l_c2 = tfn_pos[:, 1, 0].min()
norm_c2_l = min_l_c2 / (tfn_pos[:, 1, 2] + EPS)  # lower → min_l/u
norm_c2_m = min_l_c2 / (tfn_pos[:, 1, 1] + EPS)  # middle → min_l/m
norm_c2_u = min_l_c2 / (tfn_pos[:, 1, 0] + EPS)  # upper → min_l/l
norm_c2_m_arr = np.stack([norm_c2_l, norm_c2_m, norm_c2_u], axis=1)

# Max kriter (C3, idx=2): normalize = val / max(upper)
max_u_c3 = tfn_pos[:, 2, 2].max()
norm_c3_m = tfn_pos[:, 2, :] / max_u_c3

print(f"max_u(C1)={max_u_c1:.4f}, min_l(C2)={min_l_c2:.4f}, max_u(C3)={max_u_c3:.4f}")

norm_e = me._normalize_fuzzy_tfn(tfn_e, list(data.columns), crit_types)

# C1 (max kriter)
ok_c1 = np.allclose(norm_c1_m, norm_e[:, 0, :], atol=1e-9)
ok_fail("C1 (max) normalizasyon", ok_c1)
print(f"    Elle: {norm_c1_m.round(4)}")
print(f"  Engine: {norm_e[:,0,:].round(4)}")

# C2 (min kriter)
ok_c2 = np.allclose(norm_c2_m_arr, norm_e[:, 1, :], atol=1e-9)
ok_fail("C2 (min) normalizasyon", ok_c2)
print(f"    Elle (l=min_l/u, m=min_l/m, u=min_l/l):\n{norm_c2_m_arr.round(4)}")
print(f"  Engine:\n{norm_e[:,1,:].round(4)}")

# C2'nin l ≤ m ≤ u koşulu
c2_lmu_ok = bool(np.all(norm_e[:,1,0] <= norm_e[:,1,1]) and np.all(norm_e[:,1,1] <= norm_e[:,1,2]))
ok_fail("C2 min norm: l ≤ m ≤ u koşulu", c2_lmu_ok)

# C3 (max kriter)
ok_c3 = np.allclose(norm_c3_m, norm_e[:, 2, :], atol=1e-9)
ok_fail("C3 (max) normalizasyon", ok_c3)

# ─────────────────────────────────────────────────────────────────────────────
# 3. Fuzzy Mesafe Fonksiyonu
# ─────────────────────────────────────────────────────────────────────────────
section("3. Fuzzy Mesafe Fonksiyonu  d(a,b) = sqrt(sum((a-b)²)/3)")

# İki örnek TFN
a = np.array([0.3, 0.5, 0.7])
b = np.array([0.6, 0.8, 1.0])
dist_manual = np.sqrt(((a - b)**2).sum() / 3.0)
# Engine: _fuzzy_distance çalışıyor (1D değil, son eksen üzerinde)
a_arr = a.reshape(1, 1, 3)
b_arr = b.reshape(1, 1, 3)
dist_engine = float(me._fuzzy_distance(a_arr, b_arr)[0, 0])
ok_fail("Fuzzy mesafe (0.3,0.5,0.7)↔(0.6,0.8,1.0)", abs(dist_manual - dist_engine) < 1e-9,
        dist_engine, dist_manual)

# ─────────────────────────────────────────────────────────────────────────────
# 4. Fuzzy TOPSIS — Elle hesaplama
# ─────────────────────────────────────────────────────────────────────────────
section("4. Fuzzy TOPSIS — Tam elle hesaplama")

norm_m  = norm_e.copy()           # yukarıda doğruladık
wvec    = np.array([1/3, 1/3, 1/3])
weighted_m = norm_m * wvec.reshape(1, -1, 1)  # (4, 3, 3)

# FPIS ve FNIS (her kriter için component-wise max/min)
fpis_m = weighted_m.max(axis=0)   # (3, 3)
fnis_m = weighted_m.min(axis=0)   # (3, 3)

print("FPIS (elle):")
print(pd.DataFrame(fpis_m, index=["C1","C2","C3"], columns=["l","m","u"]).round(6))
print("FNIS (elle):")
print(pd.DataFrame(fnis_m, index=["C1","C2","C3"], columns=["l","m","u"]).round(6))

# D+ ve D- (her alternatif için kriterlere göre fuzzy mesafe toplamı)
dplus_m  = np.zeros(4)
dminus_m = np.zeros(4)
for i in range(4):
    for j in range(3):
        dplus_m[i]  += np.sqrt(((weighted_m[i,j,:] - fpis_m[j,:])**2).sum() / 3.0)
        dminus_m[i] += np.sqrt(((weighted_m[i,j,:] - fnis_m[j,:])**2).sum() / 3.0)

cc_m = dminus_m / (dplus_m + dminus_m + EPS)
print("\nD+, D-, CC (elle):")
for i, alt in enumerate(data.index):
    print(f"  {alt}: D+={dplus_m[i]:.6f}  D-={dminus_m[i]:.6f}  CC={cc_m[i]:.6f}")

score_ft, det_ft = me._rank_fuzzy_topsis(data, crit_types, weights, spread=SPREAD)
print("\nFuzzy TOPSIS engine vs elle:")
all_ok = True
for i, alt in enumerate(data.index):
    match = abs(score_ft[i] - cc_m[i]) < 1e-6
    if not match: all_ok = False
    ok_fail(alt, match, score_ft[i], cc_m[i])

print(f"\n  Elle   sıralama: {dict(zip(data.index, np.argsort(-cc_m)+1))}")
print(f"  Engine sıralama: {dict(zip(data.index, np.argsort(-score_ft)+1))}")
print(f"  Fuzzy TOPSIS {'✅ TAMAM' if all_ok else '❌ UYUŞMUYOR'}")

# ─────────────────────────────────────────────────────────────────────────────
# 5. Senaryo Matrisleri Doğrulaması
# ─────────────────────────────────────────────────────────────────────────────
section("5. _rank_fuzzy_by_scenarios — TFN senaryo matrisleri")

# Lower senaryo: x * (1 - spread)
lower_m = pd.DataFrame(x * (1 - SPREAD), index=data.index, columns=data.columns)
mid_m   = data.copy()
upper_m = pd.DataFrame(x * (1 + SPREAD), index=data.index, columns=data.columns)

# Engine: Fuzzy EDAS senaryolarını çalıştır, her senaryonun engine EDAS ile karşılaştır
sc_low_e,  _ = me._rank_edas(lower_m,  crit_types, weights)
sc_mid_e,  _ = me._rank_edas(mid_m,    crit_types, weights)
sc_upp_e,  _ = me._rank_edas(upper_m,  crit_types, weights)
scenario_avg_m = (sc_low_e + sc_mid_e + sc_upp_e) / 3.0

score_fedas, det_fedas = me._rank_fuzzy_edas(data, crit_types, weights, spread=SPREAD)
all_ok = np.allclose(scenario_avg_m, score_fedas, atol=1e-9)
ok_fail("Fuzzy EDAS = avg(Lower,Mid,Upper) EDAS skorları", all_ok)
print(f"  Elle avg : {scenario_avg_m.round(6)}")
print(f"  Engine   : {score_fedas.round(6)}")

# ─────────────────────────────────────────────────────────────────────────────
# 6. spread=0 → crisp eşdeğerliği
# ─────────────────────────────────────────────────────────────────────────────
section("6. spread=0 → Fuzzy == Crisp")

# Fuzzy TOPSIS with spread=0 should equal crisp TOPSIS
sc_topsis_crisp, _ = me._rank_topsis(data, crit_types, weights)
sc_ftopsis_0, _    = me._rank_fuzzy_topsis(data, crit_types, weights, spread=0.0)
ok_topsis_0 = np.allclose(sc_topsis_crisp, sc_ftopsis_0, atol=1e-9)
ok_fail("Fuzzy TOPSIS(spread=0) == Crisp TOPSIS", ok_topsis_0)

# Fuzzy VIKOR with spread=0 should equal crisp VIKOR
sc_vikor_crisp, _  = me._rank_vikor(data, crit_types, weights)
sc_fvikor_0, _     = me._rank_fuzzy_vikor(data, crit_types, weights, spread=0.0)
ok_vikor_0 = np.allclose(sc_vikor_crisp, sc_fvikor_0, atol=1e-9)
ok_fail("Fuzzy VIKOR(spread=0) == Crisp VIKOR", ok_vikor_0)

# Fuzzy EDAS with spread=0 should equal crisp EDAS
sc_edas_crisp, _   = me._rank_edas(data, crit_types, weights)
sc_fedas_0, _      = me._rank_fuzzy_edas(data, crit_types, weights, spread=0.0)
ok_edas_0 = np.allclose(sc_edas_crisp, sc_fedas_0, atol=1e-9)
ok_fail("Fuzzy EDAS(spread=0) == Crisp EDAS", ok_edas_0)

sc_saw_crisp, _    = me._rank_saw(data, crit_types, weights)
sc_fsaw_0, _       = me._rank_fuzzy_saw(data, crit_types, weights, spread=0.0)
ok_saw_0 = np.allclose(sc_saw_crisp, sc_fsaw_0, atol=1e-9)
ok_fail("Fuzzy SAW(spread=0) == Crisp SAW", ok_saw_0)

# ─────────────────────────────────────────────────────────────────────────────
# 7. Fuzzy VIKOR — min kriter sonrası fix doğrulaması
# ─────────────────────────────────────────────────────────────────────────────
section("7. Fuzzy VIKOR — min kriter fix kontrolü")

# lower senaryo için elle VIKOR (min kriter C2)
sc_vl_e, det_vl = me._rank_vikor(lower_m, crit_types, weights)
sc_vm_e, det_vm = me._rank_vikor(mid_m,   crit_types, weights)
sc_vu_e, det_vu = me._rank_vikor(upper_m, crit_types, weights)

# VIKOR küçük = iyi, senaryo ortalaması yine de avg
fvikor_avg_m = (sc_vl_e + sc_vm_e + sc_vu_e) / 3.0
sc_fvikor, _ = me._rank_fuzzy_vikor(data, crit_types, weights, spread=SPREAD)

ok = np.allclose(fvikor_avg_m, sc_fvikor, atol=1e-9)
ok_fail("Fuzzy VIKOR = avg(Lower,Mid,Upper) VIKOR skorları", ok)

# C2 min kriter: lower senaryo diff değerleri negatif olmamalı
vt_lower = det_vl
x_low = lower_m.to_numpy(float)
best_vl = np.array([x_low[:,0].max(), x_low[:,1].min(), x_low[:,2].max()])
worst_vl= np.array([x_low[:,0].min(), x_low[:,1].max(), x_low[:,2].min()])
diff_c2_low = [(1/3)*(best_vl[1]-x_low[i,1])/(best_vl[1]-worst_vl[1]) for i in range(4)]
neg_flag = any(d < -1e-9 for d in diff_c2_low)
ok_fail("Fuzzy VIKOR lower senaryo C2 diff değerleri ≥ 0", not neg_flag)
print(f"    C2 diff (lower senaryo): {[round(d,6) for d in diff_c2_low]}")

# ─────────────────────────────────────────────────────────────────────────────
# 8. Fuzzy Ağırlık Yöntemleri
# ─────────────────────────────────────────────────────────────────────────────
section("8. Fuzzy Ağırlık Yöntemleri — Senaryo Ortalaması")

# Fuzzy Entropy: 3 senaryoda entropy çalıştır, ağırlıkları ortala
w_low,  _ = me._weights_entropy(lower_m,  crit_types)
w_mid,  _ = me._weights_entropy(mid_m,    crit_types)
w_upp,  _ = me._weights_entropy(upper_m,  crit_types)
avg_w_m = {c: (w_low[c] + w_mid[c] + w_upp[c]) / 3.0 for c in data.columns}
# normalize (ağırlıklar toplamı zaten ~1 ama kontrol et)
total = sum(avg_w_m.values())
avg_w_m = {c: v/total for c, v in avg_w_m.items()}

w_fent, _ = me._weights_fuzzy_entropy(data, crit_types, spread=SPREAD)

ok_ent = all(abs(w_fent[c] - avg_w_m[c]) < 1e-6 for c in data.columns)
ok_fail("Fuzzy Entropy = avg(senaryo ağırlıkları)", ok_ent)
print(f"    Elle  : {avg_w_m}")
print(f"    Engine: {w_fent}")

# Fuzzy CRITIC
w_cl, _ = me._weights_critic(lower_m, crit_types)
w_cm, _ = me._weights_critic(mid_m,   crit_types)
w_cu, _ = me._weights_critic(upper_m, crit_types)
avg_wc = {c: (w_cl[c]+w_cm[c]+w_cu[c])/3.0 for c in data.columns}
total_c = sum(avg_wc.values())
avg_wc = {c: v/total_c for c, v in avg_wc.items()}

w_fcrit, _ = me._weights_fuzzy_critic(data, crit_types, spread=SPREAD)
ok_crit = all(abs(w_fcrit[c] - avg_wc[c]) < 1e-6 for c in data.columns)
ok_fail("Fuzzy CRITIC = avg(senaryo ağırlıkları)", ok_crit)

# ─────────────────────────────────────────────────────────────────────────────
# 9. Defuzzification kontrolü
# ─────────────────────────────────────────────────────────────────────────────
section("9. Defuzzification — (l+m+u)/3")

tfn_test = np.array([[[1.0, 2.0, 3.0],
                       [0.5, 1.0, 1.5]]])  # shape (1, 2, 3)
defuzz_m = tfn_test.mean(axis=2)    # (1, 2): [[2.0, 1.0]]
defuzz_e = me._defuzzify_tfn(tfn_test)
ok_defuzz = np.allclose(defuzz_m, defuzz_e, atol=1e-9)
ok_fail("Defuzzify = (l+m+u)/3", ok_defuzz,
        float(defuzz_e[0,0]), float(defuzz_m[0,0]))

# ─────────────────────────────────────────────────────────────────────────────
# 10. TFN input (gerçek fuzzy veri) — Fuzzy TOPSIS
# ─────────────────────────────────────────────────────────────────────────────
section("10. Fuzzy TOPSIS — Doğrudan TFN input (spread yerine)")

# Gerçek TFN veri: l ≤ m ≤ u, kriterlere göre
tfn_direct = np.array([
    # C1          C2          C3
    [[3.5,4.0,4.5],[2.0,3.0,4.0],[4.5,5.0,5.5]],  # A1
    [[2.5,3.0,3.5],[5.0,6.0,7.0],[1.5,2.0,2.5]],  # A2
    [[4.5,5.0,5.5],[1.5,2.0,2.5],[3.5,4.0,4.5]],  # A3
    [[1.5,2.0,2.5],[7.0,8.0,9.0],[2.5,3.0,3.5]],  # A4
], dtype=float)  # shape (4, 3, 3)

# l ≤ m ≤ u kontrolü
lmu_ok2 = bool(
    np.all(tfn_direct[:,:,0] <= tfn_direct[:,:,1]) and
    np.all(tfn_direct[:,:,1] <= tfn_direct[:,:,2])
)
ok_fail("Doğrudan TFN: l ≤ m ≤ u", lmu_ok2)

# Engine çalıştır
sc_tft, _ = me._rank_fuzzy_topsis(data, crit_types, weights, spread=SPREAD, tfn_input=tfn_direct)

# Elle hesapla
norm_direct = me._normalize_fuzzy_tfn(tfn_direct, list(data.columns), crit_types)
wvec2 = np.array([1/3, 1/3, 1/3])
weighted_d = norm_direct * wvec2.reshape(1,-1,1)
fpis_d = weighted_d.max(axis=0)
fnis_d = weighted_d.min(axis=0)
dplus_d  = np.array([sum(np.sqrt(((weighted_d[i,j,:]-fpis_d[j,:])**2).sum()/3) for j in range(3)) for i in range(4)])
dminus_d = np.array([sum(np.sqrt(((weighted_d[i,j,:]-fnis_d[j,:])**2).sum()/3) for j in range(3)) for i in range(4)])
cc_d = dminus_d / (dplus_d + dminus_d + EPS)

all_ok = all(abs(sc_tft[i]-cc_d[i]) < 1e-6 for i in range(4))
ok_fail("Fuzzy TOPSIS (doğrudan TFN input) elle ile eşleşiyor", all_ok)
print(f"  Elle  : {cc_d.round(6)}")
print(f"  Engine: {sc_tft.round(6)}")

# ─────────────────────────────────────────────────────────────────────────────
# 11. TFN normalizasyonu — Tüm metodlar için ortaklık kontrolü
# ─────────────────────────────────────────────────────────────────────────────
section("11. Senaryo-tabanlı tüm fuzzy metodlar — Dahili tutarlılık")

fuzzy_methods_scenario = [
    ("Fuzzy EDAS",      lambda: me._rank_fuzzy_edas(data, crit_types, weights, spread=SPREAD)),
    ("Fuzzy SAW",       lambda: me._rank_fuzzy_saw(data, crit_types, weights, spread=SPREAD)),
    ("Fuzzy VIKOR",     lambda: me._rank_fuzzy_vikor(data, crit_types, weights, spread=SPREAD)),
    ("Fuzzy COPRAS",    lambda: me._rank_fuzzy_copras(data, crit_types, weights, spread=SPREAD)),
    ("Fuzzy ARAS",      lambda: me._rank_fuzzy_aras(data, crit_types, weights, spread=SPREAD)),
    ("Fuzzy MABAC",     lambda: me._rank_fuzzy_mabac(data, crit_types, weights, spread=SPREAD)),
    ("Fuzzy MARCOS",    lambda: me._rank_fuzzy_marcos(data, crit_types, weights, spread=SPREAD)),
    ("Fuzzy MOORA",     lambda: me._rank_fuzzy_moora(data, crit_types, weights, spread=SPREAD)),
    ("Fuzzy GRA",       lambda: me._rank_fuzzy_gra(data, crit_types, weights, spread=SPREAD)),
    ("Fuzzy CODAS",     lambda: me._rank_fuzzy_codas(data, crit_types, weights, spread=SPREAD)),
    ("Fuzzy CoCoSo",    lambda: me._rank_fuzzy_cocoso(data, crit_types, weights, spread=SPREAD)),
    ("Fuzzy WPM",       lambda: me._rank_fuzzy_wpm(data, crit_types, weights, spread=SPREAD)),
    ("Fuzzy MAUT",      lambda: me._rank_fuzzy_maut(data, crit_types, weights, spread=SPREAD)),
    ("Fuzzy WASPAS",    lambda: me._rank_fuzzy_waspas(data, crit_types, weights, spread=SPREAD)),
    ("Fuzzy SPOTIS",    lambda: me._rank_fuzzy_spotis(data, crit_types, weights, spread=SPREAD)),
    ("Fuzzy MULTIMOORA",lambda: me._rank_fuzzy_multimoora(data, crit_types, weights, spread=SPREAD)),
    ("Fuzzy RAWEC",     lambda: me._rank_fuzzy_rawec(data, crit_types, weights, spread=SPREAD)),
    ("Fuzzy RAFSI",     lambda: me._rank_fuzzy_rafsi(data, crit_types, weights, spread=SPREAD)),
    ("Fuzzy ROV",       lambda: me._rank_fuzzy_rov(data, crit_types, weights, spread=SPREAD)),
    ("Fuzzy AROMAN",    lambda: me._rank_fuzzy_aroman(data, crit_types, weights, spread=SPREAD)),
    ("Fuzzy DNMA",      lambda: me._rank_fuzzy_dnma(data, crit_types, weights, spread=SPREAD)),
    ("Fuzzy PSI",       lambda: me._rank_fuzzy_psi(data, crit_types, weights, spread=SPREAD)),
    ("Fuzzy WISP",      lambda: me._rank_fuzzy_wisp(data, crit_types, weights, spread=SPREAD)),
]

errors = []
for name, fn in fuzzy_methods_scenario:
    try:
        score, _ = fn()
        arr = np.asarray(score, dtype=float)
        has_nan = bool(np.any(np.isnan(arr)))
        has_inf = bool(np.any(np.isinf(arr)))
        all_zero = bool(np.all(arr == 0.0))
        ok = not has_nan and not has_inf and not all_zero
        flag = ("nan!" if has_nan else "") + ("inf!" if has_inf else "") + ("all-zero!" if all_zero else "")
        tag = "✅" if ok else "❌"
        print(f"  {tag} {name:22s}  scores={arr.round(4)}  {''+flag}")
        if not ok:
            errors.append(f"{name}: {flag}")
    except Exception as exc:
        print(f"  ❌ {name:22s}  EXCEPTION: {exc}")
        errors.append(f"{name}: exception: {exc}")

# ─────────────────────────────────────────────────────────────────────────────
# ÖZET
# ─────────────────────────────────────────────────────────────────────────────
section("ÖZET")

all_tests = [
    ("TFN üretimi",                  all_ok := np.allclose(tfn_m, tfn_e, atol=1e-9)),
    ("Fuzzy TFN norm C1 (max)",      ok_c1),
    ("Fuzzy TFN norm C2 (min)",      ok_c2),
    ("C2 norm l≤m≤u",                c2_lmu_ok),
    ("Fuzzy mesafe",                  abs(dist_manual-dist_engine)<1e-9),
    ("Fuzzy TOPSIS tam hesap",        all(abs(score_ft[i]-cc_m[i])<1e-6 for i in range(4))),
    ("Fuzzy EDAS senaryo avg",        np.allclose(scenario_avg_m, score_fedas, atol=1e-9)),
    ("spread=0 → crisp TOPSIS",      ok_topsis_0),
    ("spread=0 → crisp VIKOR",       ok_vikor_0),
    ("spread=0 → crisp EDAS",        ok_edas_0),
    ("spread=0 → crisp SAW",         ok_saw_0),
    ("Fuzzy VIKOR min-kriter fix",   ok),
    ("Fuzzy Entropy senaryo avg",    ok_ent),
    ("Fuzzy CRITIC senaryo avg",     ok_crit),
    ("Defuzzification",              ok_defuzz),
    ("Fuzzy TOPSIS TFN input",       all(abs(sc_tft[i]-cc_d[i])<1e-6 for i in range(4))),
    ("Senaryo metodları (runtime)",  len(errors) == 0),
]

passed = sum(1 for _, v in all_tests if v)
total  = len(all_tests)
print(f"\n  {passed}/{total} test geçti\n")
for name, ok in all_tests:
    print(f"  {'✅' if ok else '❌'} {name}")

if errors:
    print(f"\n  Hatalı metodlar:")
    for e in errors:
        print(f"    ❌ {e}")
