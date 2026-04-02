import numpy as np
from scipy.optimize import minimize
from typing import Dict, Any

EPS = 1e-12


def _normalize_weights(values: np.ndarray) -> np.ndarray:
    arr = np.asarray(values, dtype=float)
    arr = np.where(np.isfinite(arr), arr, 0.0)
    arr = np.clip(arr, 0.0, None)
    total = float(arr.sum())
    if total <= EPS:
        return np.ones(len(arr), dtype=float) / max(len(arr), 1)
    return arr / total


def _scenario_triplets(values: np.ndarray, spread: float, *, low: float | None = None, high: float | None = None) -> tuple[np.ndarray, np.ndarray, np.ndarray]:
    middle = np.asarray(values, dtype=float).copy()
    delta = np.abs(middle) * max(float(spread), 0.0)
    lower = middle - delta
    upper = middle + delta
    if low is not None:
        lower = np.maximum(lower, low)
        middle = np.maximum(middle, low)
        upper = np.maximum(upper, low)
    if high is not None:
        lower = np.minimum(lower, high)
        middle = np.minimum(middle, high)
        upper = np.minimum(upper, high)
    return lower, middle, upper


def _aggregate_scenario_weights(results: list[Dict[str, Any]]) -> np.ndarray:
    weight_mat = np.asarray([res["weights"] for res in results], dtype=float)
    return _normalize_weights(weight_mat.mean(axis=0))


def _build_fuzzy_ahp_scenarios(matrix: np.ndarray, spread: float) -> tuple[np.ndarray, np.ndarray, np.ndarray]:
    n = matrix.shape[0]
    lower = np.ones((n, n), dtype=float)
    middle = np.ones((n, n), dtype=float)
    upper = np.ones((n, n), dtype=float)

    for r in range(n):
        lower[r, r] = 1.0
        middle[r, r] = 1.0
        upper[r, r] = 1.0
        for c in range(r + 1, n):
            modal = float(matrix[r, c]) if np.isfinite(matrix[r, c]) else 1.0
            if modal <= EPS:
                modal = 1.0
            if modal >= 1.0:
                l_val = max(1.0 / 9.0, modal * (1.0 - spread))
                u_val = min(9.0, modal * (1.0 + spread))
            else:
                inv_modal = 1.0 / modal
                inv_lower = max(1.0 / 9.0, inv_modal * (1.0 - spread))
                inv_upper = min(9.0, inv_modal * (1.0 + spread))
                l_val = 1.0 / inv_upper
                u_val = 1.0 / inv_lower

            lower[r, c] = l_val
            middle[r, c] = modal
            upper[r, c] = u_val
            lower[c, r] = 1.0 / u_val
            middle[c, r] = 1.0 / modal
            upper[c, r] = 1.0 / l_val

    return lower, middle, upper

def calc_ahp(matrix: np.ndarray) -> Dict[str, Any]:
    """
    Analytic Hierarchy Process (AHP)
    matrix: n x n ikili karşılaştırma matrisi
    """
    n = matrix.shape[0]
    eigenvalues, eigenvectors = np.linalg.eig(matrix)
    max_idx = np.argmax(np.real(eigenvalues))
    principal_eigenvalue = np.real(eigenvalues[max_idx])
    principal_eigenvector = np.real(eigenvectors[:, max_idx])
    
    weights = principal_eigenvector / np.sum(principal_eigenvector)
    weights = np.clip(weights, 0, None)
    
    # Consistency Index (CI) and Consistency Ratio (CR)
    ci = (principal_eigenvalue - n) / (n - 1) if n > 1 else 0
    ri_values = {1: 0.00, 2: 0.00, 3: 0.58, 4: 0.90, 5: 1.12, 
                 6: 1.24, 7: 1.32, 8: 1.41, 9: 1.45, 10: 1.49}
    ri = ri_values.get(n, 1.49)
    cr = ci / ri if ri > 0 else 0
    
    return {
        "weights": weights.tolist(),
        "cr": float(cr),
        "lambda_max": float(principal_eigenvalue),
        "is_consistent": cr <= 0.10,
        "details": {
            "ci": float(ci),
            "ri": float(ri)
        }
    }


def calc_fuzzy_ahp(matrix: np.ndarray, spread: float = 0.15) -> Dict[str, Any]:
    """
    Fuzzy AHP
    matrix: n x n ikili karşılaştırma matrisi (modal değerler)
    spread: modal değerin alt/üst bulanık sapması
    """
    lower, middle, upper = _build_fuzzy_ahp_scenarios(np.asarray(matrix, dtype=float), spread)
    scenario_defs = [("Lower", lower), ("Middle", middle), ("Upper", upper)]
    scenario_results = []
    for scenario_name, scenario_matrix in scenario_defs:
        res = calc_ahp(scenario_matrix)
        scenario_results.append(
            {
                "scenario": scenario_name,
                "weights": res["weights"],
                "cr": res["cr"],
                "lambda_max": res["lambda_max"],
                "is_consistent": res["is_consistent"],
            }
        )

    weights = _aggregate_scenario_weights(scenario_results)
    cr_vals = np.asarray([r["cr"] for r in scenario_results], dtype=float)
    lambda_vals = np.asarray([r["lambda_max"] for r in scenario_results], dtype=float)
    return {
        "weights": weights.tolist(),
        "cr": float(cr_vals.mean()),
        "lambda_max": float(lambda_vals.mean()),
        "is_consistent": bool(all(r["is_consistent"] for r in scenario_results)),
        "details": {
            "spread": float(spread),
            "scenario_results": scenario_results,
            "ci_proxy": float(cr_vals.mean()),
        },
    }

def calc_bwm(best_to_others: np.ndarray, others_to_worst: np.ndarray) -> Dict[str, Any]:
    """
    Best-Worst Method (BWM) - Linear Model
    best_to_others: 1xN vector (Best kriterin diğerlerine üstünlüğü, 1-9 arası)
    others_to_worst: 1xN vector (Diğerlerinin Worst kritere üstünlüğü, 1-9 arası)
    Not: best_to_others[best_idx] = 1, others_to_worst[worst_idx] = 1 olmalıdır.
    """
    n = len(best_to_others)
    best_idx = np.argmin(best_to_others) 
    worst_idx = np.argmin(others_to_worst)
    
    def objective(x):
        return x[-1] # minimize xi
    
    def constraint_sum(x):
        return np.sum(x[:-1]) - 1.0 # sum(w) = 1
        
    bounds = [(0, 1) for _ in range(n)] + [(0, None)] # w_j >= 0, xi >= 0
    
    constraints = [{'type': 'eq', 'fun': constraint_sum}]
    
    for j in range(n):
        if j != best_idx:
            constraints.append({'type': 'ineq', 'fun': lambda x, j=j: x[-1] - (x[best_idx] - best_to_others[j] * x[j])})
            constraints.append({'type': 'ineq', 'fun': lambda x, j=j: x[-1] + (x[best_idx] - best_to_others[j] * x[j])})
            
        if j != worst_idx:
            constraints.append({'type': 'ineq', 'fun': lambda x, j=j: x[-1] - (x[j] - others_to_worst[j] * x[worst_idx])})
            constraints.append({'type': 'ineq', 'fun': lambda x, j=j: x[-1] + (x[j] - others_to_worst[j] * x[worst_idx])})
            
    x0 = np.ones(n+1) / n
    x0[-1] = 0.1
    
    res = minimize(objective, x0, method='SLSQP', bounds=bounds, constraints=constraints)
    
    weights = res.x[:-1]
    xi = res.x[-1]
    
    # Consistency computation
    a_bw = float(best_to_others[worst_idx])
    ci_values = {1: 0.0, 2: 0.44, 3: 1.00, 4: 1.63, 5: 2.30, 6: 3.00, 7: 3.73, 8: 4.47, 9: 5.23}
    # Round a_bw to closest integer key or fallback to max
    a_bw_int = int(round(a_bw))
    if a_bw_int > 9: a_bw_int = 9
    ci = ci_values.get(a_bw_int, 5.23)
    
    cr = xi / ci if ci > 0 else 0
    
    return {
        "weights": weights.tolist(),
        "xi": float(xi),
        "cr": float(cr),
        "is_consistent": cr <= 0.10,
        "details": {
            "a_bw": float(a_bw),
            "ci": float(ci)
        }
    }


def calc_fuzzy_bwm(best_to_others: np.ndarray, others_to_worst: np.ndarray, spread: float = 0.15) -> Dict[str, Any]:
    """
    Fuzzy BWM
    best_to_others / others_to_worst: modal 1-9 vektörleri
    """
    bto = np.asarray(best_to_others, dtype=float)
    otw = np.asarray(others_to_worst, dtype=float)
    lower_bto, middle_bto, upper_bto = _scenario_triplets(bto, spread, low=1.0, high=9.0)
    lower_otw, middle_otw, upper_otw = _scenario_triplets(otw, spread, low=1.0, high=9.0)
    lower_bto = np.where(np.isclose(bto, 1.0), 1.0, lower_bto)
    middle_bto = np.where(np.isclose(bto, 1.0), 1.0, middle_bto)
    upper_bto = np.where(np.isclose(bto, 1.0), 1.0, upper_bto)
    lower_otw = np.where(np.isclose(otw, 1.0), 1.0, lower_otw)
    middle_otw = np.where(np.isclose(otw, 1.0), 1.0, middle_otw)
    upper_otw = np.where(np.isclose(otw, 1.0), 1.0, upper_otw)

    scenario_defs = [
        ("Lower", lower_bto, lower_otw),
        ("Middle", middle_bto, middle_otw),
        ("Upper", upper_bto, upper_otw),
    ]
    scenario_results = []
    for scenario_name, s_bto, s_otw in scenario_defs:
        res = calc_bwm(s_bto, s_otw)
        scenario_results.append(
            {
                "scenario": scenario_name,
                "weights": res["weights"],
                "xi": res["xi"],
                "cr": res["cr"],
                "is_consistent": res["is_consistent"],
            }
        )

    weights = _aggregate_scenario_weights(scenario_results)
    xi_vals = np.asarray([r["xi"] for r in scenario_results], dtype=float)
    cr_vals = np.asarray([r["cr"] for r in scenario_results], dtype=float)
    return {
        "weights": weights.tolist(),
        "xi": float(xi_vals.mean()),
        "cr": float(cr_vals.mean()),
        "is_consistent": bool(all(r["is_consistent"] for r in scenario_results)),
        "details": {
            "spread": float(spread),
            "scenario_results": scenario_results,
        },
    }

def calc_swara(s_j: np.ndarray) -> Dict[str, Any]:
    """
    SWARA (Step-wise Weight Assessment Ratio Analysis)
    s_j: Uzmanlar tarafından belirlenen göreli önem (relative importance) değerleri.
         İlk eleman 0.0 olmalıdır (En önemli kriterin kendine göreceği bir üstünlük yoktur).
         Girdi kriterler önem sırasına göre ZATEN SIRALANMIŞ kabul edilir.
    """
    n = len(s_j)
    k_j = np.zeros(n)
    q_j = np.zeros(n)
    
    k_j[0] = 1.0
    q_j[0] = 1.0
    
    for j in range(1, n):
        k_j[j] = s_j[j] + 1.0
        q_j[j] = q_j[j-1] / k_j[j]
        
    weights = q_j / np.sum(q_j)
    
    return {
        "weights": weights.tolist(),
        "details": {
            "s_j": s_j.tolist(),
            "k_j": k_j.tolist(),
            "q_j": q_j.tolist()
        }
    }


def calc_fuzzy_swara(s_j: np.ndarray, spread: float = 0.15) -> Dict[str, Any]:
    """
    Fuzzy SWARA
    s_j: modal göreli önem azaltım değerleri
    """
    modal = np.asarray(s_j, dtype=float)
    lower, middle, upper = _scenario_triplets(modal, spread, low=0.0)
    if len(modal) > 0:
        lower[0] = middle[0] = upper[0] = 0.0
    scenario_defs = [("Lower", lower), ("Middle", middle), ("Upper", upper)]
    scenario_results = []
    for scenario_name, scenario_values in scenario_defs:
        res = calc_swara(scenario_values)
        scenario_results.append({"scenario": scenario_name, "weights": res["weights"], "details": res["details"]})

    weights = _aggregate_scenario_weights(scenario_results)
    return {
        "weights": weights.tolist(),
        "details": {
            "spread": float(spread),
            "scenario_results": scenario_results,
        },
    }

def calc_dematel(matrix: np.ndarray) -> Dict[str, Any]:
    """
    DEMATEL (Decision Making Trial and Evaluation Laboratory)
    matrix: Kriterlerin birbirini etkileme derecelerini gösteren direkt ilişki matrisi 
            (Genellikle 0-4 arası uzman puanlarıyla doldurulur).
    """
    n = matrix.shape[0]
    
    # Normalizasyon The max(row_sum, col_sum) is used
    row_sums = np.sum(matrix, axis=1)
    col_sums = np.sum(matrix, axis=0)
    max_sum = max(np.max(row_sums), np.max(col_sums))
    
    X = matrix / max_sum if max_sum > 0 else matrix
    
    # Toplam İlişki Matrisi (T = X * (I - X)^-1)
    I = np.eye(n)
    _singular_matrix = False
    try:
        T = np.dot(X, np.linalg.inv(I - X))
    except np.linalg.LinAlgError:
        T = np.zeros_like(X)
        _singular_matrix = True
    
    R = np.sum(T, axis=1) # Row sums (Etkileyen)
    C = np.sum(T, axis=0) # Col sums (Etkilenen)
    
    D_plus_R = R + C # Prominence (Önem - Merkezilik)
    D_minus_R = R - C # Relation (Sebep/Sonuç - Cause/Effect)
    
    # Kriter ağırlıkları olarak Prominence (D+R) kullanılması yaygındır
    weights = D_plus_R / np.sum(D_plus_R) if np.sum(D_plus_R) > 0 else np.ones(n)/n
    
    return {
        "weights": weights.tolist(),
        "prominence": D_plus_R.tolist(),
        "relation": D_minus_R.tolist(),
        "singular_warning": _singular_matrix,
        "details": {
            "R": R.tolist(),
            "C": C.tolist(),
            "T": T.tolist()
        }
    }


def calc_fuzzy_dematel(matrix: np.ndarray, spread: float = 0.15) -> Dict[str, Any]:
    """
    Fuzzy DEMATEL
    matrix: modal etki matrisi (0-4 arası)
    """
    modal = np.asarray(matrix, dtype=float)
    lower, middle, upper = _scenario_triplets(modal, spread, low=0.0, high=4.0)
    np.fill_diagonal(lower, 0.0)
    np.fill_diagonal(middle, 0.0)
    np.fill_diagonal(upper, 0.0)

    scenario_defs = [("Lower", lower), ("Middle", middle), ("Upper", upper)]
    scenario_results = []
    any_singular = False
    for scenario_name, scenario_matrix in scenario_defs:
        res = calc_dematel(scenario_matrix)
        any_singular = any_singular or bool(res.get("singular_warning"))
        scenario_results.append(
            {
                "scenario": scenario_name,
                "weights": res["weights"],
                "prominence": res["prominence"],
                "relation": res["relation"],
                "singular_warning": res.get("singular_warning", False),
            }
        )

    weights = _aggregate_scenario_weights(scenario_results)
    prominence = np.mean(np.asarray([r["prominence"] for r in scenario_results], dtype=float), axis=0)
    relation = np.mean(np.asarray([r["relation"] for r in scenario_results], dtype=float), axis=0)
    return {
        "weights": weights.tolist(),
        "prominence": prominence.tolist(),
        "relation": relation.tolist(),
        "singular_warning": any_singular,
        "details": {
            "spread": float(spread),
            "scenario_results": scenario_results,
        },
    }

def calc_smart(points: np.ndarray) -> Dict[str, Any]:
    """
    SMART (Simple Multi-Attribute Rating Technique)
    points: Kriterlere verilen doğrudan uzman puanları (Genellikle 10-100 arası vs).
    """
    weights = points / np.sum(points) if np.sum(points) > 0 else np.ones(len(points))/len(points)
    return {
        "weights": weights.tolist(),
        "details": {
            "points": points.tolist()
        }
    }


def calc_fuzzy_smart(points: np.ndarray, spread: float = 0.15) -> Dict[str, Any]:
    """
    Fuzzy SMART
    points: modal doğrudan uzman puanları
    """
    modal = np.asarray(points, dtype=float)
    lower, middle, upper = _scenario_triplets(modal, spread, low=0.0)
    scenario_defs = [("Lower", lower), ("Middle", middle), ("Upper", upper)]
    scenario_results = []
    for scenario_name, scenario_points in scenario_defs:
        res = calc_smart(scenario_points)
        scenario_results.append(
            {
                "scenario": scenario_name,
                "weights": res["weights"],
                "points": scenario_points.tolist(),
            }
        )

    weights = _aggregate_scenario_weights(scenario_results)
    return {
        "weights": weights.tolist(),
        "details": {
            "spread": float(spread),
            "scenario_results": scenario_results,
        },
    }
