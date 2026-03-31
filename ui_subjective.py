import streamlit as st
import pandas as pd
import numpy as np
import hashlib
import mcdm_subjective as sub_engine

def tt(tr, en):
    return en if st.session_state.get("ui_lang", "TR") == "EN" else tr

def _normalize_dematel_label(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    text = " ".join(str(value).strip().split())
    if not text or text.lower() == "nan":
        return ""
    return text.casefold()

def _has_meaningful_dematel_labels(labels) -> bool:
    for label in labels:
        norm = _normalize_dematel_label(label)
        if norm and not norm.replace(".", "", 1).isdigit():
            return True
    return False

def _read_dematel_upload(uploaded_file) -> pd.DataFrame:
    file_name = str(getattr(uploaded_file, "name", "") or "").lower()
    try:
        uploaded_file.seek(0)
    except Exception:
        pass

    if file_name.endswith(".csv"):
        df = pd.read_csv(uploaded_file, header=None)
    elif file_name.endswith(".xlsx"):
        df = pd.read_excel(uploaded_file, header=None)
    else:
        raise ValueError(
            tt(
                "DEMATEL içe aktarımı için yalnızca CSV veya XLSX desteklenir.",
                "Only CSV or XLSX files are supported for DEMATEL import.",
            )
        )

    df = df.dropna(how="all").dropna(axis=1, how="all").reset_index(drop=True)
    if df.empty:
        raise ValueError(tt("Yüklenen DEMATEL dosyası boş görünüyor.", "The uploaded DEMATEL file appears to be empty."))
    return df

def _try_parse_fuzzy_dematel(raw_df: pd.DataFrame) -> pd.DataFrame | None:
    if raw_df.shape[0] < 3 or raw_df.shape[1] < 4:
        return None

    subheader = raw_df.iloc[1].fillna("").astype(str).str.strip().str.casefold()
    triplets: list[tuple[int, object]] = []
    col = 1
    while col + 2 < raw_df.shape[1]:
        trio = subheader.iloc[col:col + 3].tolist()
        header = raw_df.iloc[0, col]
        if trio == ["l", "m", "u"] and _normalize_dematel_label(header):
            triplets.append((col, header))
            col += 3
            continue
        col += 1

    if not triplets:
        return None

    rows: list[list[float]] = []
    row_labels: list[object] = []
    for ridx in range(2, raw_df.shape[0]):
        row_label = raw_df.iloc[ridx, 0]
        if not _normalize_dematel_label(row_label):
            continue
        parsed_row: list[float] = []
        for start_col, _header in triplets:
            vals = pd.to_numeric(raw_df.iloc[ridx, start_col:start_col + 3], errors="coerce")
            if vals.isna().all():
                parsed_row = []
                break
            if vals.isna().any():
                raise ValueError(
                    tt(
                        f"Satır '{row_label}' için eksik bulanık üçlü bulundu. Her ilişki için l, m ve u değerlerinin tamamı gerekir.",
                        f"An incomplete fuzzy triplet was found on row '{row_label}'. Each relation needs l, m, and u values.",
                    )
                )
            parsed_row.append(float(vals.mean()))
        if parsed_row:
            row_labels.append(row_label)
            rows.append(parsed_row)

    if not rows:
        return None

    col_labels = [header for _, header in triplets]
    return pd.DataFrame(rows, index=row_labels, columns=col_labels, dtype=float)

def _try_parse_crisp_dematel(raw_df: pd.DataFrame) -> pd.DataFrame | None:
    numeric = raw_df.apply(pd.to_numeric, errors="coerce")
    if raw_df.shape[0] == raw_df.shape[1] and numeric.notna().all().all():
        return pd.DataFrame(numeric.to_numpy(dtype=float))

    if raw_df.shape[0] >= 2 and raw_df.shape[1] >= 2:
        body = raw_df.iloc[1:, 1:]
        body_num = body.apply(pd.to_numeric, errors="coerce")
        if body.shape[0] == body.shape[1] and body_num.notna().all().all():
            return pd.DataFrame(
                body_num.to_numpy(dtype=float),
                index=raw_df.iloc[1:, 0].tolist(),
                columns=raw_df.iloc[0, 1:].tolist(),
            )

        body = raw_df.iloc[:, 1:]
        body_num = body.apply(pd.to_numeric, errors="coerce")
        if body.shape[0] == body.shape[1] and body_num.notna().all().all():
            return pd.DataFrame(body_num.to_numpy(dtype=float), index=raw_df.iloc[:, 0].tolist())

        body = raw_df.iloc[1:, :]
        body_num = body.apply(pd.to_numeric, errors="coerce")
        if body.shape[0] == body.shape[1] and body_num.notna().all().all():
            return pd.DataFrame(body_num.to_numpy(dtype=float), columns=raw_df.iloc[0, :].tolist())

    return None

def _align_dematel_matrix(matrix_df: pd.DataFrame, criteria: list[str]) -> pd.DataFrame:
    n_crit = len(criteria)
    if matrix_df.shape != (n_crit, n_crit):
        raise ValueError(
            tt(
                f"Yüklenen DEMATEL matrisi {n_crit}x{n_crit} boyutunda olmalı; mevcut dosya {matrix_df.shape[0]}x{matrix_df.shape[1]}.",
                f"The uploaded DEMATEL matrix must be {n_crit}x{n_crit}; the current file is {matrix_df.shape[0]}x{matrix_df.shape[1]}.",
            )
        )

    label_map = {_normalize_dematel_label(name): name for name in criteria}
    row_named = _has_meaningful_dematel_labels(matrix_df.index)
    col_named = _has_meaningful_dematel_labels(matrix_df.columns)

    if row_named or col_named:
        mapped_rows = [label_map.get(_normalize_dematel_label(value)) for value in matrix_df.index]
        mapped_cols = [label_map.get(_normalize_dematel_label(value)) for value in matrix_df.columns]
        if any(value is None for value in mapped_rows) or any(value is None for value in mapped_cols):
            raise ValueError(
                tt(
                    "Yüklenen DEMATEL matrisindeki kriter adları, uygulamadaki mevcut kriterlerle eşleşmiyor.",
                    "The criterion names in the uploaded DEMATEL matrix do not match the current criteria in the app.",
                )
            )
        matrix_df = matrix_df.copy()
        matrix_df.index = mapped_rows
        matrix_df.columns = mapped_cols
        if len(set(matrix_df.index)) != n_crit or len(set(matrix_df.columns)) != n_crit:
            raise ValueError(
                tt(
                    "Yüklenen DEMATEL matrisinde yinelenen kriter etiketleri bulundu.",
                    "Duplicate criterion labels were found in the uploaded DEMATEL matrix.",
                )
            )
        matrix_df = matrix_df.loc[criteria, criteria]
    else:
        matrix_df = matrix_df.copy()
        matrix_df.index = criteria
        matrix_df.columns = criteria

    numeric_df = matrix_df.apply(pd.to_numeric, errors="coerce")
    if numeric_df.isna().any().any():
        raise ValueError(
            tt(
                "Yüklenen DEMATEL matrisi sayısal olmayan veya eksik değerler içeriyor.",
                "The uploaded DEMATEL matrix contains missing or non-numeric values.",
            )
        )
    if (numeric_df.to_numpy(dtype=float) < 0).any():
        raise ValueError(
            tt(
                "Yüklenen DEMATEL matrisi negatif değer içeremez.",
                "The uploaded DEMATEL matrix cannot contain negative values.",
            )
        )
    return numeric_df.astype(float)

def _parse_dematel_upload(uploaded_file, criteria: list[str]) -> tuple[pd.DataFrame, str]:
    raw_df = _read_dematel_upload(uploaded_file)
    parse_note = ""

    matrix_df = _try_parse_fuzzy_dematel(raw_df)
    if matrix_df is not None:
        parse_note = tt(
            "Bulanık DEMATEL biçimi algılandı; l-m-u üçlüleri ağırlık merkezi yöntemiyle tek değere dönüştürüldü.",
            "A fuzzy DEMATEL layout was detected; l-m-u triplets were defuzzified into single values using the centroid method.",
        )
    else:
        matrix_df = _try_parse_crisp_dematel(raw_df)
        if matrix_df is None:
            raise ValueError(
                tt(
                    "Dosya DEMATEL matrisi olarak çözülemedi. Kare sayısal matris veya l-m-u sütunlarından oluşan bulanık DEMATEL biçimi bekleniyor.",
                    "The file could not be parsed as a DEMATEL matrix. Expected a square numeric matrix or a fuzzy DEMATEL layout with l-m-u columns.",
                )
            )

    matrix_df = _align_dematel_matrix(matrix_df, criteria)
    matrix = matrix_df.to_numpy(dtype=float)

    if not np.allclose(np.diag(matrix), 0.0):
        np.fill_diagonal(matrix, 0.0)
        matrix_df = pd.DataFrame(matrix, index=criteria, columns=criteria)
        diag_note = tt(
            "Köşegen değerler DEMATEL gereği otomatik olarak 0 yapıldı.",
            "Diagonal values were automatically set to 0 as required by DEMATEL.",
        )
        parse_note = f"{parse_note} {diag_note}".strip()

    return matrix_df, parse_note

def render_subjective_component(criteria: list[str]):
    n_crit = len(criteria)
    if n_crit < 2:
        st.warning(tt("En az 2 kriter olmalıdır.", "At least 2 criteria required."))
        return

    st.markdown(f"#### {tt('Uzman Görüşü Hesaplama Paneli', 'Expert Judgment Panel')}")
    
    num_experts = st.number_input(
        tt("Değerlendirme Yapacak Uzman Sayısı", "Number of Experts"),
        min_value=1, max_value=100, value=1, step=1,
        help=tt("Maksimum 100 uzmana kadar desteklenir. Değerler Geometrik Ortalama ile (AIP yöntemi) birleştirilip Normalize edilir.",
                "Up to 100 experts supported. Weights are aggregated via Geometric Mean (AIP method) and Normalized.")
    )
    
    st.info(tt(
        "Bu alanda yapacağınız hesaplamalar otomatik olarak 'Manuel Ağırlık' sistemine aktarılır. "
        "Uzman sayısı en fazla **100** olabilir.",
        "Calculations made here are automatically transferred to the 'Manual Weight' system. "
        "Maximum number of experts is **100**.",
    ))

    # ── Global Klasik / Fuzzy seçimi ──────────────────────────────────────────
    _calc_mode = st.radio(
        tt("Hesaplama Modu", "Calculation Mode"),
        [tt("🟦 Klasik", "🟦 Classical"), tt("🟪 Fuzzy", "🟪 Fuzzy")],
        horizontal=True,
        key="subjective_calc_mode",
        help=tt(
            "Klasik: kesin sayısal değerlerle hesaplama. Fuzzy: belirsizlik spread'i ile üçgen bulanık sayı (TFN) senaryoları üretilir.",
            "Classical: calculation with crisp values. Fuzzy: triangular fuzzy number (TFN) scenarios generated from uncertainty spread.",
        ),
    )
    _is_fuzzy_sub = "Fuzzy" in _calc_mode or "🟪" in _calc_mode

    fuzzy_spread = float(st.session_state.get("fuzzy_spread", 0.15))
    if _is_fuzzy_sub:
        fuzzy_spread = st.slider(
            tt("Bulanık Belirsizlik Genişliği", "Fuzzy Uncertainty Spread"),
            min_value=0.05,
            max_value=0.40,
            value=fuzzy_spread,
            step=0.01,
            help=tt(
                "Fuzzy yöntemlerde modal uzman değerinin etrafında oluşturulacak alt-orta-üst belirsizlik bandını belirler.",
                "Defines the lower-middle-upper uncertainty band around the modal expert value in fuzzy methods.",
            ),
        )
        st.session_state["fuzzy_spread"] = float(fuzzy_spread)

    tabs = st.tabs(["AHP", "BWM", "SWARA", "DEMATEL", "SMART"])
    
    def apply_weights(weights_list, method_name):
        weights_dict = {c: float(w) for c, w in zip(criteria, weights_list)}
        st.session_state["subjective_manual_weights"] = weights_dict
        st.session_state["subjective_method_name"] = method_name
        if num_experts > 1:
            st.success(f"{method_name} " + tt(f"ağırlıkları {num_experts} uzmanın görüşleri Geometrik Ortalama ile birleştirilerek sisteme aktarıldı!", 
                                              f"weights for {num_experts} experts aggregated via Geometric Mean and loaded!"))
        else:
            st.success(f"{method_name} " + tt("ağırlıkları sisteme başarıyla aktarıldı!", "weights successfully loaded!"))
        st.rerun()

    def build_expert_tabs(prefix_key: str, default_df_func):
        if num_experts == 1:
            exp_tabs = [st.container()]
        else:
            exp_tabs = st.tabs([f"{tt('Uzman', 'Expert')} {i}" for i in range(1, num_experts + 1)])
        
        dfs = []
        for i in range(num_experts):
            with exp_tabs[i]:
                k = f"{prefix_key}_{n_crit}_exp_{i}"
                if k not in st.session_state:
                    st.session_state[k] = default_df_func()
                dfs.append(st.data_editor(st.session_state[k], key=f"edit_{k}", use_container_width=True))
        return dfs

    def aggregate_weights(all_weights):
        # all_weights: list of numpy arrays, shape: (num_experts, n_crit)
        w_mat = np.array(all_weights)
        geom = np.prod(w_mat, axis=0) ** (1.0 / num_experts)
        normalized = geom / np.sum(geom)
        return normalized

    def render_consistency(results_list, method):
        if num_experts == 1:
            res = results_list[0]
            if "cr" in res:
                cols = st.columns(3)
                cols[0].metric("CR", f"{res['cr']:.4f}")
                cols[1].metric(tt("Tutarlı Mı?", "Is Consistent?"), tt("Evet","Yes") if res['is_consistent'] else tt("Hayır","No"))
                if "AHP" in method:
                    cols[2].metric("Lambda Max", f"{res.get('lambda_max',0):.4f}")
                elif "BWM" in method:
                    cols[2].metric("Xi Değeri", f"{res.get('xi',0):.4f}")
                if not res['is_consistent']:
                    st.warning(tt("CR > 0.10! Matris tutarsız.", "CR > 0.10! Matrix is inconsistent."))
        else:
            cr_vals = [r['cr'] for r in results_list if 'cr' in r]
            if cr_vals:
                avg_cr = np.mean(cr_vals)
                max_cr = np.max(cr_vals)
                all_consistent = all([r['is_consistent'] for r in results_list])
                cols = st.columns(3)
                cols[0].metric(tt("Ortalama CR", "Avg CR"), f"{avg_cr:.4f}")
                cols[1].metric(tt("Maksimum CR", "Max CR"), f"{max_cr:.4f}")
                cols[2].metric(tt("Tümü Tutarlı Mı?", "All Consistent?"), tt("Evet","Yes") if all_consistent else tt("Hayır","No"))
                if not all_consistent:
                    st.warning(tt("En az bir uzmanın hesabı tutarsız (CR > 0.10).", "At least one expert is inconsistent (CR > 0.10)."))

    # --- AHP ---
    with tabs[0]:
        st.subheader("AHP (Analitik Hiyerarşi Prosesi)")
        with st.expander(tt("📌 Yöntem Hakkında Bilgi & Örnek Anket Formu", "📌 Method Info & Sample Survey Form")):
            st.markdown(tt("""
            **🌟 Yöntemin Güçlü Yanları:**
            AHP, kriterleri birbirleriyle ikili olarak (pairwise) karşılaştırmaya dayanır. En büyük gücü, uzmanların verdiği yanıtların kendi içindeki **Matematiksel Tutarlılığını (CR)** ölçebilmesidir. Eğer uzman çelişkili yanıtlar vermişse (Örn: A > B, B > C ama C > A dediyse), CR değeri 0.10'un üzerine çıkarak uyarı verir.
            
            **📝 Örnek Gönderilecek Anket Formu:**
            Lütfen Kriter İkililerini birbirlerine göre 1 ile 9 arasındaki Saaty skalasını kullanarak değerlendirin:
            - **[1]** Eşit Önemde                 
            - **[3]** Biraz Daha Önemli
            - **[5]** Güçlü Şekilde Önemli        
            - **[7]** Çok Güçlü Önemli  
            - **[9]** Kesin / Mutlak Üstün
            
            *(Örnek Soru: "Üretim Kalitesi" kriteri, "Maliyet" kriterine göre ne derece üstündür?)*
            """,
            """
            **🌟 Strengths of the Method:**
            AHP is based on pairwise comparisons. Its greatest strength lies in its ability to measure the **Mathematical Consistency (CR)** of the expert's answers.
            
            **📝 Sample Survey Form:**
            Please evaluate Criterion pairs using the 1-9 Saaty scale:
            - [1] Equal Importance
            - [3] Moderate Importance
            - [5] Strong Importance
            ...
            """))

        st.write(tt("İkili Karşılaştırma Matrisi (Üst üçgeni doldurmanız yeterlidir, alt üçgen otomatik hesaplanır):", 
                    "Pairwise Comparison Matrix (Fill upper triangle, lower is auto calculated):"))
        if _is_fuzzy_sub:
            st.caption(tt(
                f"Fuzzy AHP aynı matrisi kullanır; {fuzzy_spread:.2f} belirsizlik genişliği ile alt-orta-üst senaryolar oluşturur.",
                f"Fuzzy AHP uses the same matrix and builds lower-middle-upper scenarios with uncertainty spread {fuzzy_spread:.2f}.",
            ))

        dfs_ahp = build_expert_tabs("ahp", lambda: pd.DataFrame(np.ones((n_crit, n_crit)), columns=criteria, index=criteria))

        if not _is_fuzzy_sub:
            if st.button(tt("AHP Hesapla & Uygula", "Calculate & Apply AHP"), key="ahp_btn"):
                try:
                    all_weights = []
                    all_res = []
                    for df in dfs_ahp:
                        m_copy = df.to_numpy(dtype=float).copy()
                        for r in range(n_crit):
                            for c in range(r+1, n_crit):
                                val = m_copy[r, c]
                                if val == 0: val = 1.0
                                m_copy[r, c] = val
                                m_copy[c, r] = 1.0 / val
                        res = sub_engine.calc_ahp(m_copy)
                        all_weights.append(res['weights'])
                        all_res.append(res)
                    final_w = aggregate_weights(all_weights)
                    render_consistency(all_res, "AHP")
                    apply_weights(final_w.tolist(), "AHP")
                except Exception as e:
                    st.error(f"Hata: {e}")
        else:
            if st.button(tt("Fuzzy AHP Hesapla & Uygula", "Calculate & Apply Fuzzy AHP"), key="fuzzy_ahp_btn"):
                try:
                    all_weights = []
                    all_res = []
                    for df in dfs_ahp:
                        m_copy = df.to_numpy(dtype=float).copy()
                        for r in range(n_crit):
                            for c in range(r+1, n_crit):
                                val = m_copy[r, c]
                                if val == 0: val = 1.0
                                m_copy[r, c] = val
                                m_copy[c, r] = 1.0 / val
                        res = sub_engine.calc_fuzzy_ahp(m_copy, spread=fuzzy_spread)
                        all_weights.append(res["weights"])
                        all_res.append(res)
                    final_w = aggregate_weights(all_weights)
                    render_consistency(all_res, "Fuzzy AHP")
                    apply_weights(final_w.tolist(), "Fuzzy AHP")
                except Exception as e:
                    st.error(f"Hata: {e}")

    # --- BWM ---
    with tabs[1]:
        st.subheader("BWM (Best-Worst Method)")
        with st.expander(tt("📌 Yöntem Hakkında Bilgi & Örnek Anket Formu", "📌 Method Info & Sample Survey Form")):
            st.markdown(tt("""
            **🌟 Yöntemin Güçlü Yanları:**
            BWM, AHP'ye göre çok daha az sayıda soru (kıyaslama) sorulmasını gerektirir ($2n-3$ vs $\dfrac{n(n-1)}{2}$). Karar vericinin sadece belirlediği En İyi (Best) ve En Kötü (Worst) kritere göre diğerlerini değerlendirmesini ister. Uzmanlar için bilişsel yükü düşüktür ve tutarlılığı genellikle daha yüksektir.
            
            **📝 Örnek Gönderilecek Anket Formu:**
            - **Aşama 1:** Lütfen elinizdeki kriterler listesinden **SİZCE EN ÖNEMLİ (Best)** olan 1 kriteri seçin. Ardından **EN ÖNEMSİZ (Worst)** olan 1 kriteri seçin.
            - **Aşama 2 (En İyinin Üstünlüğü):** Seçtiğiniz EN İYİ kriter, diğer kalan kriterlere göre 1 ile 9 arasında ne derece üstündür? (Kendisine olan derecesi 1'dir).
            - **Aşama 3 (En Kötünün Zayıflığı):** Diğer kalan kriterlerin tümü, seçtiğiniz EN KÖTÜ kritere göre 1 ile 9 arasında ne derece üstündür? 
            """,
            """
            **🌟 Strengths of the Method:**
            BWM requires very few comparisons compared to AHP. Identifying the Best and Worst criteria makes it highly consistent.
            
            **📝 Sample Survey Form:**
            - Step 1: Identify the MOST Important and LEAST Important criteria.
            - Step 2: Rate the Best criterion against all others (1-9).
            - Step 3: Rate all others against the Worst criterion (1-9).
            """))

        st.markdown(tt("1. **En iyi** ve **en kötü** kriteri belirleyin.\n2. 1-9 arası puanlayın.", "1. Best/Worst criteria.\n2. Score 1-9."))
        if _is_fuzzy_sub:
            st.caption(tt(
                f"Fuzzy BWM, aynı modal puanlardan {fuzzy_spread:.2f} genişliğinde alt-orta-üst üstünlük senaryoları üretir.",
                f"Fuzzy BWM derives lower-middle-upper dominance scenarios from the same modal scores using spread {fuzzy_spread:.2f}.",
            ))

        dfs_bwm = build_expert_tabs("bwm", lambda: pd.DataFrame({
                "Best-to-Others (1-9)": [1.0 for _ in range(n_crit)],
                "Others-to-Worst (1-9)": [3.0 for _ in range(n_crit)]
            }, index=criteria))

        if not _is_fuzzy_sub:
            if st.button(tt("BWM Hesapla & Uygula", "Calculate & Apply BWM"), key="bwm_btn"):
                try:
                    all_weights = []
                    all_res = []
                    for df in dfs_bwm:
                        bto = df["Best-to-Others (1-9)"].to_numpy(dtype=float)
                        otw = df["Others-to-Worst (1-9)"].to_numpy(dtype=float)
                        res = sub_engine.calc_bwm(bto, otw)
                        all_weights.append(res['weights'])
                        all_res.append(res)
                    final_w = aggregate_weights(all_weights)
                    render_consistency(all_res, "BWM")
                    apply_weights(final_w.tolist(), "BWM")
                except Exception as e:
                    st.error(f"Hata: {e}")
        else:
            if st.button(tt("Fuzzy BWM Hesapla & Uygula", "Calculate & Apply Fuzzy BWM"), key="fuzzy_bwm_btn"):
                try:
                    all_weights = []
                    all_res = []
                    for df in dfs_bwm:
                        bto = df["Best-to-Others (1-9)"].to_numpy(dtype=float)
                        otw = df["Others-to-Worst (1-9)"].to_numpy(dtype=float)
                        res = sub_engine.calc_fuzzy_bwm(bto, otw, spread=fuzzy_spread)
                        all_weights.append(res["weights"])
                        all_res.append(res)
                    final_w = aggregate_weights(all_weights)
                    render_consistency(all_res, "Fuzzy BWM")
                    apply_weights(final_w.tolist(), "Fuzzy BWM")
                except Exception as e:
                    st.error(f"Hata: {e}")

    # --- SWARA ---
    with tabs[2]:
        st.subheader("SWARA")
        with st.expander(tt("📌 Yöntem Hakkında Bilgi & Örnek Anket Formu", "📌 Method Info & Sample Survey Form")):
            st.markdown(tt("""
            **🌟 Yöntemin Güçlü Yanları:**
            SWARA (Step-wise Weight Assessment Ratio Analysis), karmaşık ikili kıyaslamalar yerine uzmanların konuyu sezgisel olarak derecelendirmesine odaklanır. Kullanımı ve öğrenilmesi inanılmaz kolaydır.
            
            **📝 Örnek Gönderilecek Anket Formu:**
            Lütfen anket formundaki adımları izleyin:
            1. Kriterlerinizi, size göre **en önemliden en önemsize doğru 1., 2., 3. şeklinde sıralayın.**
            2. En baştaki (1.) kritere herhangi bir değer atamayın (Değeri boş kalacak, veya 0 olacak).
            3. İkinci sıradaki kriter, birinciye göre YÜZDE KAÇ DAHA AZ ÖNEMLİDİR? (Örneğin %20 oranında daha az önemli olduğunu düşünüyorsanız boşluğa **0.20** yazınız.)
            4. Bu şekilde her kriteri, bir üstündeki kritere göre yüzde kaç eksik hissettirdiğini (örneğin 0.15, 0.30 gibi) değerlendirerek aşağıya doğru ilerleyin.
            """,
            """
            **🌟 Strengths of the Method:**
            SWARA provides a step-by-step intuitive ranking with an incredibly fast and straightforward assessment.
            
            **📝 Sample Survey Form:**
            1. Rank the criteria from most important to least important.
            2. The first gets 0.
            3. Rate the consecutive items by how much less important they are compared to the one above it (e.g., 0.20 for 20% less).
            """))

        st.info(tt("Her bir kriterin, bir üstündeki kritere kıyasla yüzde kaç daha az önemli olduğunu ondalık değer (örn: 0.15) olarak giriniz. İlk kriter her zaman 0 olmalıdır.",
                   "Enter how much less important each criterion is compared to the one above it as a decimal (e.g. 0.15). The first value must be 0."))
        if _is_fuzzy_sub:
            st.caption(tt(
                f"Fuzzy SWARA, bu azalış oranlarını {fuzzy_spread:.2f} bandında bulanıklaştırarak senaryo ortalaması alır.",
                f"Fuzzy SWARA fuzzifies these reduction ratios with spread {fuzzy_spread:.2f} and averages the scenario weights.",
            ))

        dfs_swara = build_expert_tabs("swara", lambda: pd.DataFrame({
                "s_j (İlk değer 0)": [0.0] + [0.2 for _ in range(n_crit-1)]
            }, index=criteria))

        if not _is_fuzzy_sub:
            if st.button(tt("SWARA Hesapla & Uygula", "Calculate & Apply SWARA"), key="swara_btn"):
                try:
                    all_weights = []
                    for df in dfs_swara:
                        sj = df["s_j (İlk değer 0)"].to_numpy(dtype=float)
                        res = sub_engine.calc_swara(sj)
                        all_weights.append(res['weights'])
                    final_w = aggregate_weights(all_weights)
                    apply_weights(final_w.tolist(), "SWARA")
                except Exception as e:
                    st.error(f"Hata: {e}")
        else:
            if st.button(tt("Fuzzy SWARA Hesapla & Uygula", "Calculate & Apply Fuzzy SWARA"), key="fuzzy_swara_btn"):
                try:
                    all_weights = []
                    for df in dfs_swara:
                        sj = df["s_j (İlk değer 0)"].to_numpy(dtype=float)
                        res = sub_engine.calc_fuzzy_swara(sj, spread=fuzzy_spread)
                        all_weights.append(res["weights"])
                    final_w = aggregate_weights(all_weights)
                    apply_weights(final_w.tolist(), "Fuzzy SWARA")
                except Exception as e:
                    st.error(f"Hata: {e}")

    # --- DEMATEL ---
    with tabs[3]:
        st.subheader("DEMATEL")
        with st.expander(tt("📌 Yöntem Hakkında Bilgi & Örnek Anket Formu", "📌 Method Info & Sample Survey Form")):
            st.markdown(tt("""
            **🌟 Yöntemin Güçlü Yanları:**
            DEMATEL, sadece kriter ağırlığı değil, kriterlerin **Birbirleri Üzerindeki Etkileşimini (Sebep-Sonuç İlişkisini)** bulmaya yarar. Eğer kriter ağınızda karmaşık bir iç etkileşim mekanizması varsa en iyi alternatiftir.
            
            **📝 Örnek Gönderilecek Anket Formu:**
            Lütfen Kriter İkililerinin birbirini Etkileme (Tetikleme) derecesini 0'dan 4'e kadar puanlayınız:
            - **[0]** Hiç Etkilemez
            - **[1]** Düşük Derecede Etkiler
            - **[2]** Orta Derecede Etkiler
            - **[3]** Yüksek Derecede Etkiler
            - **[4]** Çok Yüksek Derecede Etkiler
            
            *(Örnek Soru: Üretim Hızı faktörü arttığında / değiştiğinde, Üretim Maliyetini doğrudan ve tek başına ne derece tetikler?)*
            """,
            """
            **🌟 Strengths of the Method:**
            DEMATEL captures cause-and-effect relationships among criteria effectively.
            
            **📝 Sample Survey Form:**
            Rate how strongly Criterion i influences Criterion j (0 to 4 scale).
            - [0] No influence -> [4] Very high influence
            """))
        st.info(tt(
            "Manuel girişte 0-4 ölçeğini kullanın. CSV/XLSX yüklemede kare DEMATEL matrisleri ve l-m-u sütunlu bulanık DEMATEL dosyaları da desteklenir.",
            "Use the 0-4 scale for manual entry. CSV/XLSX upload also supports square DEMATEL matrices and fuzzy DEMATEL files with l-m-u columns."
        ))
        if _is_fuzzy_sub:
            st.caption(tt(
                f"Fuzzy DEMATEL, girilen modal etki değerlerinden {fuzzy_spread:.2f} genişliğinde alt-orta-üst etki senaryoları kurar.",
                f"Fuzzy DEMATEL builds lower-middle-upper influence scenarios with spread {fuzzy_spread:.2f} from the entered modal influence values.",
            ))
        
        def dem_def():
            dmat = np.ones((n_crit, n_crit)) * 2 # Varsayılan olarak Orta etki
            np.fill_diagonal(dmat, 0) # Köşegen daima 0
            return pd.DataFrame(dmat, columns=criteria, index=criteria)
        if num_experts == 1:
            dem_tabs = [st.container()]
        else:
            dem_tabs = st.tabs([f"{tt('Uzman', 'Expert')} {i}" for i in range(1, num_experts + 1)])

        dfs_dem = []
        for i in range(num_experts):
            with dem_tabs[i]:
                dem_key = f"dem_{n_crit}_exp_{i}"
                if dem_key not in st.session_state:
                    st.session_state[dem_key] = dem_def()

                upload = st.file_uploader(
                    tt("DEMATEL matrisi yükle (CSV/XLSX)", "Upload DEMATEL matrix (CSV/XLSX)"),
                    type=["csv", "xlsx"],
                    key=f"dematel_upload_{n_crit}_{i}",
                    help=tt(
                        "Kare sayısal DEMATEL matrisi veya l-m-u üçlülerinden oluşan bulanık DEMATEL dosyası yükleyebilirsiniz.",
                        "Upload a square numeric DEMATEL matrix or a fuzzy DEMATEL file with l-m-u triplets.",
                    ),
                )
                upload_sig_key = f"dematel_upload_sig_{n_crit}_{i}"
                if upload is not None:
                    upload_sig = hashlib.sha1(upload.getvalue()).hexdigest()
                    if st.session_state.get(upload_sig_key) != upload_sig:
                        try:
                            parsed_df, parse_note = _parse_dematel_upload(upload, criteria)
                        except Exception as exc:
                            st.error(str(exc))
                        else:
                            st.session_state[dem_key] = parsed_df
                            st.session_state[upload_sig_key] = upload_sig
                            st.success(tt("DEMATEL matrisi yüklendi ve tabloya aktarıldı.", "The DEMATEL matrix was loaded into the table."))
                            if parse_note:
                                st.caption(parse_note)

                dfs_dem.append(st.data_editor(st.session_state[dem_key], key=f"edit_{dem_key}", use_container_width=True))
        
        if not _is_fuzzy_sub:
            if st.button(tt("DEMATEL Hesapla & Uygula", "Calculate & Apply DEMATEL"), key="dem_btn"):
                try:
                    all_weights = []
                    any_singular = False
                    for df in dfs_dem:
                        dem_mat = df.to_numpy(dtype=float)
                        res = sub_engine.calc_dematel(dem_mat)
                        all_weights.append(res['weights'])
                        if res.get('singular_warning'):
                            any_singular = True
                    final_w = aggregate_weights(all_weights)
                    apply_weights(final_w.tolist(), "DEMATEL")
                    if any_singular:
                        st.warning(tt(
                            "⚠️ Etki matrisi tekil (singular) — (I–X) tersi alınamadı. Sonuçlar güvenilir olmayabilir.",
                            "⚠️ Influence matrix is singular — (I–X) could not be inverted. Results may be unreliable.",
                        ))
                except Exception as e:
                    st.error(f"Hata: {e}")
        else:
            if st.button(tt("Fuzzy DEMATEL Hesapla & Uygula", "Calculate & Apply Fuzzy DEMATEL"), key="fuzzy_dem_btn"):
                try:
                    all_weights = []
                    any_singular = False
                    for df in dfs_dem:
                        dem_mat = df.to_numpy(dtype=float)
                        res = sub_engine.calc_fuzzy_dematel(dem_mat, spread=fuzzy_spread)
                        all_weights.append(res["weights"])
                        if res.get("singular_warning"):
                            any_singular = True
                    final_w = aggregate_weights(all_weights)
                    apply_weights(final_w.tolist(), "Fuzzy DEMATEL")
                    if any_singular:
                        st.warning(tt(
                            "⚠️ En az bir bulanık DEMATEL senaryosunda matris tekil bulundu; sonuçları dikkatle yorumlayın.",
                            "⚠️ At least one fuzzy DEMATEL scenario produced a singular matrix; interpret results with caution.",
                        ))
                except Exception as e:
                    st.error(f"Hata: {e}")

    # --- SMART ---
    with tabs[4]:
        st.subheader("SMART")
        with st.expander(tt("📌 Yöntem Hakkında Bilgi & Örnek Anket Formu", "📌 Method Info & Sample Survey Form")):
            st.markdown(tt("""
            **🌟 Yöntemin Güçlü Yanları:**
            SMART (Simple Multi-Attribute Rating Technique), adı üstünde en "Basit" ancak en yaygın oransal değerlendirme yöntemlerinden biridir. Uzmanlara karmaşık matematiksel yük bindirmez, herkes tarafından kolayca uygulanıp anlaşılabilir. İkili kıyaslama yapılmaz, puanlama direkt yapılır.
            
            **📝 Örnek Gönderilecek Anket Formu:**
            Uygulamada genel yaklaşım 10'dan 100'e kadar doğrudan oran puanı vermektir:
            - Araştırmanın amacı doğrultusunda saptanan kriterlere **bağımsız olarak değerlerine (önemlerine) göre 10 ile 100 arasında doğrudan puanlar** veriniz. 
            - En önemsediğiniz kritere 100, diğerlerine ona olan oranları kadar puan atayınız.
            """,
            """
            **🌟 Strengths of the Method:**
            SMART is exceptionally straightforward. It requires direct scoring with no complex matrices.
            
            **📝 Sample Survey Form:**
            Directly score your criteria by assigning points between 10 and 100 (e.g., rate your most important as 100).
            """))
        st.info(tt("Kriterlere direkt olarak (10 ile 100 vb.) önem puanları verin.", "Assign direct importance points (e.g., 10 to 100)."))
        if _is_fuzzy_sub:
            st.caption(tt(
                f"Fuzzy SMART, bu doğrudan puanları {fuzzy_spread:.2f} belirsizlik bandıyla alt-orta-üst puan senaryolarına dönüştürür.",
                f"Fuzzy SMART converts these direct scores into lower-middle-upper scoring scenarios using spread {fuzzy_spread:.2f}.",
            ))

        dfs_smart = build_expert_tabs("smart", lambda: pd.DataFrame({
                "Puan (10-100 vb.)": [50.0 for _ in range(n_crit)]
            }, index=criteria))

        if not _is_fuzzy_sub:
            if st.button(tt("SMART Hesapla & Uygula", "Calculate & Apply SMART"), key="smart_btn"):
                try:
                    all_weights = []
                    for df in dfs_smart:
                        pts = df["Puan (10-100 vb.)"].to_numpy(dtype=float)
                        res = sub_engine.calc_smart(pts)
                        all_weights.append(res['weights'])
                    final_w = aggregate_weights(all_weights)
                    apply_weights(final_w.tolist(), "SMART")
                except Exception as e:
                    st.error(f"Hata: {e}")
        else:
            if st.button(tt("Fuzzy SMART Hesapla & Uygula", "Calculate & Apply Fuzzy SMART"), key="fuzzy_smart_btn"):
                try:
                    all_weights = []
                    for df in dfs_smart:
                        pts = df["Puan (10-100 vb.)"].to_numpy(dtype=float)
                        res = sub_engine.calc_fuzzy_smart(pts, spread=fuzzy_spread)
                        all_weights.append(res["weights"])
                    final_w = aggregate_weights(all_weights)
                    apply_weights(final_w.tolist(), "Fuzzy SMART")
                except Exception as e:
                    st.error(f"Hata: {e}")
