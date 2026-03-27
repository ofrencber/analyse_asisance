import streamlit as st
import pandas as pd
import numpy as np
import mcdm_subjective as sub_engine

def tt(tr, en):
    return en if st.session_state.get("ui_lang", "TR") == "EN" else tr

def render_subjective_component(criteria: list[str]):
    n_crit = len(criteria)
    if n_crit < 2:
        st.warning(tt("En az 2 kriter olmalıdır.", "At least 2 criteria required."))
        return

    st.markdown(f"#### {tt('Uzman Görüşü Hesaplama Paneli', 'Expert Judgment Panel')}")
    
    num_experts = st.number_input(
        tt("Değerlendirme Yapacak Uzman Sayısı", "Number of Experts"), 
        min_value=1, max_value=50, value=1, step=1,
        help=tt("Maksimum 50 uzmana kadar desteklenir. Değerler Geometrik Ortalama ile (AIP yöntemi) birleştirilip Normalize edilir.",
                "Up to 50 experts supported. Weights are aggregated via Geometric Mean (AIP method) and Normalized.")
    )
    
    st.info(tt("Bu alanda yapacağınız hesaplamalar otomatik olarak 'Manuel Ağırlık' sistemine aktarılır.",
               "Calculations made here are automatically transferred to the 'Manual Weight' system."))
    
    tabs = st.tabs(["AHP", "BWM", "SWARA", "DEMATEL", "SMART"])
    
    def apply_weights(weights_list, method_name):
        weights_dict = {c: float(w) for c, w in zip(criteria, weights_list)}
        st.session_state["subjective_manual_weights"] = weights_dict
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
                if method == "AHP":
                    cols[2].metric("Lambda Max", f"{res.get('lambda_max',0):.4f}")
                elif method == "BWM":
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
        
        dfs_ahp = build_expert_tabs("ahp", lambda: pd.DataFrame(np.ones((n_crit, n_crit)), columns=criteria, index=criteria))
        
        if st.button(tt("AHP Hesapla & Uygula", "Calculate & Apply AHP"), key="ahp_btn"):
            try:
                all_weights = []
                all_res = []
                for df in dfs_ahp:
                    m_copy = df.to_numpy(dtype=float).copy()
                    for r in range(n_crit):
                        for c in range(r+1, n_crit):
                            val = m_copy[r, c]
                            if val == 0: val = 1.0 # prevent div by zero
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
        
        dfs_bwm = build_expert_tabs("bwm", lambda: pd.DataFrame({
                "Best-to-Others (1-9)": [1.0 for _ in range(n_crit)],
                "Others-to-Worst (1-9)": [3.0 for _ in range(n_crit)]
            }, index=criteria))
        
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
        
        dfs_swara = build_expert_tabs("swara", lambda: pd.DataFrame({
                "s_j (İlk değer 0)": [0.0] + [0.2 for _ in range(n_crit-1)]
            }, index=criteria))
        
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
        st.info(tt("Sıfırdan Dörde (0-4) kadar numaralarla etkileşimi girin.", "Scale 0-4 for influences."))
        
        def dem_def():
            dmat = np.ones((n_crit, n_crit)) * 2 # Varsayılan olarak Orta etki
            np.fill_diagonal(dmat, 0) # Köşegen daima 0
            return pd.DataFrame(dmat, columns=criteria, index=criteria)
        dfs_dem = build_expert_tabs("dem", dem_def)
        
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
                        "⚠️ Etki matrisi tekil (singular) — (I–X) tersi alınamadı. "
                        "Ağırlıklar sıfır matrisinden hesaplandı; sonuçlar güvenilir olmayabilir. "
                        "Lütfen matris değerlerini kontrol edin.",
                        "⚠️ Influence matrix is singular — (I–X) could not be inverted. "
                        "Weights were computed from a zero matrix; results may be unreliable. "
                        "Please review your matrix values."
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
        
        dfs_smart = build_expert_tabs("smart", lambda: pd.DataFrame({
                "Puan (10-100 vb.)": [50.0 for _ in range(n_crit)]
            }, index=criteria))
        
        if st.button(tt("SMART Hesapla & Uygula", "Calculate & Apply SMART")):
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
