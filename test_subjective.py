import numpy as np
import json
from mcdm_subjective import calc_ahp, calc_bwm, calc_swara, calc_dematel, calc_smart

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

    # 2. BWM Testi
    # 5 kriter. En iyi olan index 0, En kötü olan index 4
    # best_to_others: En iyinin (0) diğerlerine üstünlük derecesi
    # others_to_worst: Diğerlerinin En kötüye (4) üstünlük derecesi
    best_to_others = np.array([1, 2, 4, 3, 8])
    others_to_worst = np.array([8, 4, 2, 3, 1])
    res_bwm = calc_bwm(best_to_others, others_to_worst)
    print_result("BWM", res_bwm)

    # 3. SWARA Testi
    # 4 kriter. Önem sırasına göre ZATEN dizilmiş kabul ediliyor.
    # Uzman s_j puanları (kendisinden bir öncekine kıyasla önemi)
    s_j = np.array([0.0, 0.4, 0.6, 0.2])
    res_swara = calc_swara(s_j)
    print_result("SWARA", res_swara)
    
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
    
    # 5. SMART Testi
    smart_points = np.array([100, 80, 50, 20])
    res_smart = calc_smart(smart_points)
    print_result("SMART", res_smart)

    print("\nTüm izole testler başarıyla çalıştırıldı.")

if __name__ == "__main__":
    main()
