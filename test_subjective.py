import numpy as np
from mcdm_subjective import (
    calc_ahp,
    calc_bwm,
    calc_swara,
    calc_dematel,
    calc_smart,
    calc_fuzzy_ahp,
    calc_fuzzy_bwm,
    calc_fuzzy_swara,
    calc_fuzzy_dematel,
    calc_fuzzy_smart,
)


def require(condition, message):
    if not condition:
        raise AssertionError(message)

def print_result(name, res):
    print(f"\n--- {name} Sonuçları ---")
    print(f"Ağırlıklar: {[round(w, 4) for w in res['weights']]}")
    if 'cr' in res:
        print(f"CR: {res['cr']:.4f} (Tutarlı mı? {res.get('is_consistent')})")
    if 'xi' in res:
         print(f"Xi: {res['xi']:.4f}")
    if 'prominence' in res:
         print(f"Prominence: {[round(w, 2) for w in res['prominence']]}")
         print(f"Relation (Sebep>0, Sonuç<0): {[round(w, 2) for w in res['relation']]}")
    print("-" * 40)


def assert_weight_vector(name, res):
    arr = np.asarray(res["weights"], dtype=float)
    require(np.isfinite(arr).all(), f"{name}: non-finite weights.")
    require(np.all(arr >= -1e-10), f"{name}: negative weights.")
    require(np.isclose(arr.sum(), 1.0, atol=1e-8), f"{name}: weights do not sum to 1.")

def main():
    print("MCDM Subjektif Değerlendirme Modülü Testleri Başlıyor...\n")

    # 1. AHP Testi
    # Saaty'nin klasik 3x3 matrisi
    ahp_matrix = np.array([
        [1, 3, 5],
        [1/3, 1, 3],
        [1/5, 1/3, 1]
    ])
    res_ahp = calc_ahp(ahp_matrix)
    print_result("AHP", res_ahp)
    assert_weight_vector("AHP", res_ahp)
    res_fuzzy_ahp = calc_fuzzy_ahp(ahp_matrix, spread=0.15)
    print_result("Fuzzy AHP", res_fuzzy_ahp)
    assert_weight_vector("Fuzzy AHP", res_fuzzy_ahp)

    # 2. BWM Testi
    # 5 kriter. En iyi olan index 0, En kötü olan index 4
    # best_to_others: En iyinin (0) diğerlerine üstünlük derecesi
    # others_to_worst: Diğerlerinin En kötüye (4) üstünlük derecesi
    best_to_others = np.array([1, 2, 4, 3, 8])
    others_to_worst = np.array([8, 4, 2, 3, 1])
    res_bwm = calc_bwm(best_to_others, others_to_worst)
    print_result("BWM", res_bwm)
    assert_weight_vector("BWM", res_bwm)
    res_fuzzy_bwm = calc_fuzzy_bwm(best_to_others, others_to_worst, spread=0.15)
    print_result("Fuzzy BWM", res_fuzzy_bwm)
    assert_weight_vector("Fuzzy BWM", res_fuzzy_bwm)

    # 3. SWARA Testi
    # 4 kriter. Önem sırasına göre ZATEN dizilmiş kabul ediliyor.
    # Uzman s_j puanları (kendisinden bir öncekine kıyasla önemi)
    s_j = np.array([0.0, 0.4, 0.6, 0.2])
    res_swara = calc_swara(s_j)
    print_result("SWARA", res_swara)
    assert_weight_vector("SWARA", res_swara)
    res_fuzzy_swara = calc_fuzzy_swara(s_j, spread=0.15)
    print_result("Fuzzy SWARA", res_fuzzy_swara)
    assert_weight_vector("Fuzzy SWARA", res_fuzzy_swara)
    
    # 4. DEMATEL Testi
    # 4x4 doğrudan ilişki (etkileşim) matrisi
    dem_matrix = np.array([
        [0, 2, 3, 1],
        [1, 0, 2, 2],
        [0, 1, 0, 3],
        [2, 0, 1, 0]
    ])
    res_dematel = calc_dematel(dem_matrix)
    print_result("DEMATEL", res_dematel)
    assert_weight_vector("DEMATEL", res_dematel)
    res_fuzzy_dematel = calc_fuzzy_dematel(dem_matrix, spread=0.15)
    print_result("Fuzzy DEMATEL", res_fuzzy_dematel)
    assert_weight_vector("Fuzzy DEMATEL", res_fuzzy_dematel)
    
    # 5. SMART Testi
    smart_points = np.array([100, 80, 50, 20])
    res_smart = calc_smart(smart_points)
    print_result("SMART", res_smart)
    assert_weight_vector("SMART", res_smart)
    res_fuzzy_smart = calc_fuzzy_smart(smart_points, spread=0.15)
    print_result("Fuzzy SMART", res_fuzzy_smart)
    assert_weight_vector("Fuzzy SMART", res_fuzzy_smart)

    print("\nTüm klasik ve fuzzy subjektif testler başarıyla çalıştırıldı.")

if __name__ == "__main__":
    main()
