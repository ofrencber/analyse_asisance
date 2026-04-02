import numpy as np
from scipy.optimize import minimize
from typing import Any, Dict

EPS = 1e-12

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
    try:
        T = np.dot(X, np.linalg.inv(I - X))
    except np.linalg.LinAlgError:
        T = np.zeros_like(X)
    
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
        "details": {
            "R": R.tolist(),
            "C": C.tolist(),
            "T": T.tolist()
        }
    }


def defuzzify_triangular_matrix(triangular_matrix: np.ndarray, method: str = "graded_mean") -> np.ndarray:
    """
    Convert an n x n x 3 triangular fuzzy matrix to a crisp direct-relation matrix.

    Supported methods:
    - graded_mean: (l + 2m + u) / 4
    - centroid: (l + m + u) / 3
    """
    tfn = np.asarray(triangular_matrix, dtype=float)
    if tfn.ndim != 3 or tfn.shape[-1] != 3:
        raise ValueError("Triangular fuzzy matrix must have shape (n, n, 3).")
    if not np.isfinite(tfn).all():
        raise ValueError("Triangular fuzzy matrix contains non-finite values.")
    if np.any(tfn[..., 0] > tfn[..., 1]) or np.any(tfn[..., 1] > tfn[..., 2]):
        raise ValueError("Each triangular fuzzy number must satisfy l <= m <= u.")
    if np.any(tfn < 0):
        raise ValueError("DEMATEL direct-relation values must be non-negative.")

    if method == "graded_mean":
        return (tfn[..., 0] + (2.0 * tfn[..., 1]) + tfn[..., 2]) / 4.0
    if method == "centroid":
        return np.mean(tfn, axis=2)
    raise ValueError(f"Unsupported defuzzification method: {method}")


def calc_fuzzy_dematel(triangular_matrix: np.ndarray, defuzz_method: str = "graded_mean") -> Dict[str, Any]:
    """
    Fuzzy DEMATEL using triangular fuzzy numbers followed by explicit defuzzification.
    """
    crisp_matrix = defuzzify_triangular_matrix(triangular_matrix, method=defuzz_method)
    result = calc_dematel(crisp_matrix)
    result["details"]["input_type"] = "fuzzy"
    result["details"]["defuzzification"] = defuzz_method
    result["details"]["fuzzy_matrix"] = np.asarray(triangular_matrix, dtype=float).tolist()
    result["details"]["defuzzified_matrix"] = crisp_matrix.tolist()
    return result

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
