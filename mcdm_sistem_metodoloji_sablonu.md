# MCDM Sistem Metodoloji Referans Dökümanı
# Makale Üretim Motoru İçin Tam Şablon

Bu döküman sistemin makale üretim katmanına girdi sağlar.
Kullanıcının seçtiği yöntem kombinasyonuna göre ilgili bölümler
dinamik olarak seçilir ve IMRAD yapısına yerleştirilir.

---

## BÖLÜM 0 — GENEL MİMARİ VE KULLANIM MANTIĞI

Sistem üç katmandan oluşur:

**Katman 1 — Ağırlıklandırma**
Kullanıcı objektif (Entropy, CRITIC, SD, MEREC, LOPCOW, PCA, CILOS, IDOCRIW,
Fuzzy IDOCRIW) veya subjektif (AHP, BWM, SWARA, DEMATEL, SMART) yöntemlerden
birini seçer. Çıktı her zaman crisp ağırlık vektörü $\mathbf{w} = (w_1, \ldots, w_n)$'dir.

**Katman 2 — Sıralama**
Kullanıcı klasik (24 yöntem) veya bulanık (24 yöntem) bir sıralama yöntemi seçer.
Bulanık yöntemlerde $\tilde{x}_{ij} = (x_{ij}(1-s),\, x_{ij},\, x_{ij}(1+s))$
dönüşümü otomatik uygulanır.

**Katman 3 — Sağlamlık**
Lokal duyarlılık analizi (±%10, ±%20 pertürbasyon) ve Monte Carlo simülasyonu
($N$ iterasyon, log-normal gürültü, $\sigma$ parametresi) otomatik çalışır.
Yöntem karşılaştırması Spearman sıra korelasyonu ile raporlanır.

---

## BÖLÜM 1 — AĞIRLIK YÖNTEMLERİ

### 1.A OBJEKTİF YÖNTEMLER

#### 1.A.1 Entropy (Shannon Bilgi Entropisi)

**Dayanak:** Zeleny (1982); Wang & Lee (2009)

**Akış:**

Adım 1. Ham karar matrisi $X_{m \times n}$ alınır.

Adım 2. Sum normalizasyonu uygulanır.
Fayda kriteri ($c_j = \text{max}$):
$$p_{ij} = \frac{x_{ij} - \min_i x_{ij}}{\displaystyle\sum_{i=1}^{m}(x_{ij} - \min_i x_{ij}) + \varepsilon}$$

Maliyet kriteri ($c_j = \text{min}$):
$$p_{ij} = \frac{\max_i x_{ij} - x_{ij}}{\displaystyle\sum_{i=1}^{m}(\max_i x_{ij} - x_{ij}) + \varepsilon}$$

Adım 3. Shannon entropisi hesaplanır:
$$H_j = -k \sum_{i=1}^{m} p_{ij} \ln(p_{ij}), \quad k = \frac{1}{\ln(m)}$$

Adım 4. Sapma ve ağırlık:
$$d_j = 1 - H_j, \quad w_j = \frac{d_j}{\displaystyle\sum_{j=1}^{n} d_j}$$

**Yorumlama:** $H_j \to 1$ ise kriter alternatifleri ayırt etmez; $d_j \to 0$,
ağırlık küçülür. Tüm alternatiflerde eşit değer alan kriter sıfır ağırlık alır.

---

#### 1.A.2 CRITIC (Criteria Importance Through Intercriteria Correlation)

**Dayanak:** Diakoulaki et al. (1995)

**Akış:**

Adım 1. Min-max normalizasyonu:
$$r_{ij} = \frac{x_{ij} - \min_i x_{ij}}{\max_i x_{ij} - \min_i x_{ij}}\ (\text{fayda}), \quad
r_{ij} = \frac{\max_i x_{ij} - x_{ij}}{\max_i x_{ij} - \min_i x_{ij}}\ (\text{maliyet})$$

Adım 2. Standart sapma ve Pearson korelasyon matrisi:
$$\sigma_j = \text{std}(r_{\cdot j}), \quad \rho_{jk} = \text{corr}(r_{\cdot j}, r_{\cdot k})$$

Adım 3. Bilgi içeriği ve ağırlık:
$$C_j = \sigma_j \sum_{k=1}^{n}(1 - \rho_{jk}), \quad w_j = \frac{C_j}{\displaystyle\sum_{j=1}^{n} C_j}$$

**Yorumlama:** $C_j$, hem yüksek varyansa hem de diğer kriterlerle düşük korelasyona sahip kriterleri ön plana çıkarır.

---

#### 1.A.3 Standart Sapma (SD)

**Dayanak:** Diakoulaki et al. (1992)

**Akış:**

Adım 1. Min-max normalizasyonu:
$$r_{ij} = \frac{x_{ij} - \min_i x_{ij}}{\max_i x_{ij} - \min_i x_{ij}}\ (\text{fayda}), \quad
r_{ij} = \frac{\max_i x_{ij} - x_{ij}}{\max_i x_{ij} - \min_i x_{ij}}\ (\text{maliyet})$$

Adım 2. Her kriter için standart sapma:
$$\sigma_j = \sqrt{\frac{1}{m-1}\sum_{i=1}^{m}(r_{ij} - \bar{r}_j)^2}$$

Adım 3. Ağırlık:
$$w_j = \frac{\sigma_j}{\displaystyle\sum_{j=1}^{n} \sigma_j}$$

**Yorumlama:** Yalnızca dağılım genişliği dikkate alınır; kriterler arası ilişki göz ardı edilir.

---

#### 1.A.4 MEREC (Method Based on the Removal Effects of Criteria)

**Dayanak:** Keshavarz-Ghorabaee et al. (2021)

**Akış:**

Adım 1. MEREC normalizasyonu:
$$n_{ij} = \frac{\min_i x_{ij}}{x_{ij} + \varepsilon}\ (\text{fayda}), \quad
n_{ij} = \frac{x_{ij}}{\max_i x_{ij} + \varepsilon}\ (\text{maliyet})$$

Adım 2. Toplam performans skoru:
$$S_i = \ln\!\left(1 + \frac{1}{n}\sum_{j=1}^{n}\left|\ln(n_{ij} + \varepsilon)\right|\right)$$

Adım 3. Kriter $j$ çıkarıldığında ($n-1$ kriterle):
$$S_i^{(j)} = \ln\!\left(1 + \frac{1}{n-1}\sum_{k \neq j}\left|\ln(n_{ik} + \varepsilon)\right|\right)$$

Adım 4. Çıkarım etkisi ve ağırlık:
$$E_j = \sum_{i=1}^{m}\left|S_i^{(j)} - S_i\right|, \quad w_j = \frac{E_j}{\displaystyle\sum_{j=1}^{n} E_j}$$

**Yorumlama:** Kriter çıkarıldığında toplam performans ne kadar değişiyorsa o kriter o kadar önemlidir.

---

#### 1.A.5 LOPCOW (Logarithmic Percentage Change-driven Objective Weighting)

**Dayanak:** Ecer & Pamucar (2022)

**Akış:**

Adım 1. Min-max normalizasyonu.

Adım 2. Kök ortalama kare ve standart sapma:
$$\text{RMS}_j = \sqrt{\frac{1}{m}\sum_{i=1}^{m} r_{ij}^2}, \quad \sigma_j = \text{std}(r_{\cdot j})$$

Adım 3. Logaritmik yüzde değişim gücü:
$$\text{PV}_j = \left|\ln\!\left(\frac{\text{RMS}_j}{\sigma_j + \varepsilon}\right) \times 100\right|$$

Adım 4. Ağırlık:
$$w_j = \frac{\text{PV}_j}{\displaystyle\sum_{j=1}^{n} \text{PV}_j}$$

---

#### 1.A.6 PCA (Principal Component Analysis)

**Dayanak:** Jolliffe (2002)

**Akış:**

Adım 1. Min-max normalizasyonu → StandardScaler ile z-dönüşümü.

Adım 2. PCA uygulanır; Kaiser kriteri ($\lambda_k > 1$) ile bileşen seti $\mathcal{S}$ seçilir.

Adım 3. Seçilen bileşenlerin yük değerleri ve açıklanan varyans oranları kullanılarak ham ağırlık:
$$q_j = \sum_{k \in \mathcal{S}} |\ell_{jk}| \cdot \text{VarOran}_k$$

Adım 4. Normalize ağırlık:
$$w_j = \frac{q_j}{\displaystyle\sum_{j=1}^{n} q_j}$$

---

#### 1.A.7 CILOS (Criterion Impact LOSs)

**Dayanak:** Zavadskas & Podvezko (2016)

**Akış:**

Adım 1. Fayda matrisi oluşturulur:
$$b_{ij} = \frac{x_{ij}}{\max_i x_{ij} + \varepsilon}\ (\text{fayda}), \quad
b_{ij} = \frac{\min_i x_{ij}}{x_{ij} + \varepsilon}\ (\text{maliyet})$$

Adım 2. Her kriter $j$ için en iyi alternatif: $i^*(j) = \arg\max_i b_{ij}$

Adım 3. Etki kaybı matrisi:
$$p_{jk} = \max\!\left(0,\; \frac{b_{i^*(k),k} - b_{i^*(k),j}}{b_{i^*(k),k} + \varepsilon}\right)$$

Adım 4. $F$ matrisi kurulur:
$$F_{jj} = -\sum_{k \neq j} p_{kj}, \quad F_{jk} = p_{kj}\ (j \neq k)$$

Son satır $\sum w_j = 1$ kısıtıyla değiştirilerek $\mathbf{F}\mathbf{w} = \mathbf{e}_n$ lineer sistemi çözülür.

---

#### 1.A.8 IDOCRIW (Integrated Determination of Objective CRIteria Weights)

**Dayanak:** Zavadskas et al. (2021)

**Akış:**

Adım 1. Entropy ağırlıkları $\mathbf{w}^E$ ve CILOS ağırlıkları $\mathbf{w}^C$ ayrı ayrı hesaplanır.

Adım 2. Eleman çarpımı ve normalizasyon:
$$q_j = w_j^E \cdot w_j^C, \quad w_j = \frac{q_j}{\displaystyle\sum_{j=1}^{n} q_j}$$

**Yorumlama:** Hem bilgi içeriği (Entropy) hem de kriter etki kaybı (CILOS) bütünleştirilir.

---

#### 1.A.9 Fuzzy IDOCRIW

**Dayanak:** Mevcut sistem uyarlaması

**Akış:**

Adım 1. Crisp matris $X$ üç senaryoya bölünür:
$$X^L = X(1-s), \quad X^M = X, \quad X^U = X(1+s)$$

Adım 2. Her senaryoda IDOCRIW çalıştırılır:
$$q_j^{(k)} = w_j^{E(k)} \cdot w_j^{C(k)},\quad k \in \{L, M, U\}$$

Adım 3. Senaryo ortalaması normalize edilir:
$$w_j = \text{normalize}\!\left(\frac{1}{3}\sum_{k} q_j^{(k)}\right)$$

---

### 1.B SUBJEKTİF YÖNTEMLER

#### 1.B.1 AHP (Analytic Hierarchy Process)

**Dayanak:** Saaty (1980)

**Akış:**

Adım 1. $n \times n$ ikili karşılaştırma matrisi $A$ uzman tarafından doldurulur
(Saaty 1–9 ölçeği: 1 = eşit önem, 9 = mutlak üstünlük).

Adım 2. Ana özdeğer $\lambda_{\max}$ ve ana özvektör $\mathbf{v}$ hesaplanır:
$$A\mathbf{v} = \lambda_{\max}\mathbf{v}$$

Adım 3. Normalize ağırlık:
$$w_j = \frac{v_j}{\displaystyle\sum_{j=1}^{n} v_j}, \quad v_j \geq 0$$

Adım 4. Tutarlılık kontrolü:
$$CI = \frac{\lambda_{\max} - n}{n - 1}, \quad CR = \frac{CI}{RI_n}$$

$CR \leq 0{,}10$ tutarlı kabul edilir. $RI_n$ değerleri: $RI_3 = 0{,}58$, $RI_4 = 0{,}90$,
$RI_5 = 1{,}12$, $RI_6 = 1{,}24$, $RI_7 = 1{,}32$, $RI_8 = 1{,}41$, $RI_9 = 1{,}45$.

---

#### 1.B.2 BWM (Best-Worst Method)

**Dayanak:** Rezaei (2015)

**Akış:**

Adım 1. En iyi (Best, $B$) ve en kötü (Worst, $W$) kriterler belirlenir.

Adım 2. Vektörler oluşturulur:
- $\mathbf{a}_{B} = (a_{B1}, \ldots, a_{Bn})$: Best'in diğerlerine göre üstünlüğü (1–9)
- $\mathbf{a}_{W} = (a_{1W}, \ldots, a_{nW})$: Diğerlerinin Worst'a göre üstünlüğü (1–9)

Not: $a_{BB} = 1$ ve $a_{WW} = 1$.

Adım 3. Doğrusal optimizasyon (SLSQP):
$$\min \xi \quad \text{s.t.}$$
$$|w_B - a_{Bj} w_j| \leq \xi \quad \forall j$$
$$|w_j - a_{jW} w_W| \leq \xi \quad \forall j$$
$$\sum_j w_j = 1, \quad w_j \geq 0$$

Adım 4. Tutarlılık oranı:
$$CR = \frac{\xi^*}{CI_{a_{BW}}}$$

$CI$ tablosu: $a_{BW} = 1: 0$; $3: 1{,}00$; $5: 2{,}30$; $7: 3{,}73$; $9: 5{,}23$.
$CR \leq 0{,}10$ tutarlı.

---

#### 1.B.3 SWARA (Stepwise Weight Assessment Ratio Analysis)

**Dayanak:** Keršuliene et al. (2010)

**Akış:**

Adım 1. Uzman kriterleri önem sırasına göre dizer: $c_1 \succ c_2 \succ \ldots \succ c_n$.

Adım 2. İlk kriter hariç her kriter için göreli önem $s_j$ ($s_1 = 0$) belirlenir.

Adım 3. Karşılaştırma katsayısı:
$$k_j = \begin{cases}1 & j=1 \\ s_j + 1 & j > 1\end{cases}$$

Adım 4. Göreli ağırlık:
$$q_j = \begin{cases}1 & j=1 \\ q_{j-1}/k_j & j > 1\end{cases}$$

Adım 5. Normalize ağırlık:
$$w_j = \frac{q_j}{\displaystyle\sum_{j=1}^{n} q_j}$$

---

#### 1.B.4 DEMATEL (Decision Making Trial and Evaluation Laboratory)

**Dayanak:** Fontela & Gabus (1976)

**Akış:**

Adım 1. $n \times n$ direkt ilişki matrisi $Z$ uzman tarafından doldurulur (0–4 ölçeği).

Adım 2. Normalizasyon:
$$X = \frac{Z}{\max\!\left(\max_i \sum_j z_{ij},\; \max_j \sum_i z_{ij}\right)}$$

Adım 3. Toplam ilişki matrisi:
$$T = X(I - X)^{-1}$$

Adım 4. Satır ve sütun toplamları:
$$R_i = \sum_{j=1}^{n} t_{ij}\ (\text{etkileyen}), \quad C_j = \sum_{i=1}^{n} t_{ij}\ (\text{etkilenen})$$

Adım 5. Prominence (önem) skoru ve ağırlık:
$$D_j + R_j = R_j + C_j, \quad w_j = \frac{D_j + R_j}{\displaystyle\sum_{j=1}^{n}(D_j + R_j)}$$

**Yorumlama:** $R_j - C_j > 0$: kriter neden (cause); $< 0$: kriter sonuç (effect).

---

#### 1.B.5 SMART (Simple Multi-Attribute Rating Technique)

**Dayanak:** Edwards (1971)

**Akış:**

Adım 1. Uzman her kritere $0$–$100$ arası bir puan $p_j$ atar.

Adım 2. Normalize ağırlık:
$$w_j = \frac{p_j}{\displaystyle\sum_{j=1}^{n} p_j}$$

---

## BÖLÜM 2 — SIRALAMA YÖNTEMLERİ

### 2.A KLASİK SIRALAMA YÖNTEMLERİ

Her yöntem $m$ alternatif, $n$ kriter ve ağırlık vektörü $\mathbf{w}$ alır.
$i = 1, \ldots, m$ tüm alternatifleri, $j = 1, \ldots, n$ tüm kriterleri kapsar.

---

#### 2.A.1 TOPSIS

**Dayanak:** Hwang & Yoon (1981)

**Akış:**

Adım 1. Vektör normalizasyonu (tüm alternatifler $i$, tüm kriterler $j$):
$$r_{ij} = \frac{x_{ij}}{\sqrt{\displaystyle\sum_{i=1}^{m} x_{ij}^2} + \varepsilon}$$

Adım 2. Ağırlıklı normalize matris:
$$v_{ij} = w_j \cdot r_{ij}$$

Adım 3. İdeal çözümler:
$$v_j^+ = \begin{cases}\max_i v_{ij} & c_j = \text{max}\\\min_i v_{ij} & c_j = \text{min}\end{cases}, \quad
v_j^- = \begin{cases}\min_i v_{ij} & c_j = \text{max}\\\max_i v_{ij} & c_j = \text{min}\end{cases}$$

Adım 4. Öklid uzaklıkları (tüm $i$ için):
$$D_i^+ = \sqrt{\sum_{j=1}^{n}(v_{ij} - v_j^+)^2}, \quad D_i^- = \sqrt{\sum_{j=1}^{n}(v_{ij} - v_j^-)^2}$$

Adım 5. Yakınlık katsayısı ve sıralama:
$$CC_i = \frac{D_i^-}{D_i^+ + D_i^- + \varepsilon}, \quad CC_i \in [0, 1]$$

Büyük $CC_i$ → üst sıra.

---

#### 2.A.2 VIKOR

**Dayanak:** Opricovic (1998)

**Akış:**

Adım 1. Her kriter için en iyi ve en kötü değer:
$$f_j^* = \begin{cases}\max_i x_{ij} & c_j = \text{max}\\\min_i x_{ij} & c_j = \text{min}\end{cases}, \quad
f_j^- = \begin{cases}\min_i x_{ij} & c_j = \text{max}\\\max_i x_{ij} & c_j = \text{min}\end{cases}$$

Adım 2. Maksimum grup faydası ve minimum bireysel pişmanlık:
$$S_i = \sum_{j=1}^{n} w_j \frac{f_j^* - x_{ij}}{f_j^* - f_j^- + \varepsilon}$$
$$R_i = \max_{j}\left(w_j \frac{f_j^* - x_{ij}}{f_j^* - f_j^- + \varepsilon}\right)$$

Adım 3. Uzlaşı indeksi:
$$Q_i = v\frac{S_i - S^*}{S^- - S^*} + (1-v)\frac{R_i - R^*}{R^- - R^*}$$

Burada $S^* = \min_i S_i$, $S^- = \max_i S_i$, $R^* = \min_i R_i$, $R^- = \max_i R_i$; $v = 0{,}5$ varsayılan.

Düşük $Q_i$ → üst sıra.

**Uzlaşı koşulları:**
- C1 (Kabul edilebilir avantaj): $Q(A^{(2)}) - Q(A^{(1)}) \geq 1/(m-1)$
- C2 (Kararlı kabul): $A^{(1)}$ aynı zamanda $S$ veya $R$ sıralamasında da birinci.

---

#### 2.A.3 EDAS

**Dayanak:** Keshavarz-Ghorabaee et al. (2016)

**Akış:**

Adım 1. Ortalama çözüm:
$$\text{AV}_j = \frac{1}{m}\sum_{i=1}^{m} x_{ij}$$

Adım 2. Pozitif ve negatif mesafe (tüm $i$, $j$):
Fayda ($c_j = \text{max}$):
$$\text{PDA}_{ij} = \frac{\max(0,\, x_{ij} - \text{AV}_j)}{|\text{AV}_j| + \varepsilon}, \quad
\text{NDA}_{ij} = \frac{\max(0,\, \text{AV}_j - x_{ij})}{|\text{AV}_j| + \varepsilon}$$
Maliyet ($c_j = \text{min}$): PDA ve NDA ters hesaplanır.

Adım 3. Ağırlıklı toplam mesafeler:
$$SP_i = \sum_{j=1}^{n} w_j \cdot \text{PDA}_{ij}, \quad SN_i = \sum_{j=1}^{n} w_j \cdot \text{NDA}_{ij}$$

Adım 4. Normalize ve sıralama skoru:
$$NSP_i = \frac{SP_i}{\max_i SP_i}, \quad NSN_i = 1 - \frac{SN_i}{\max_i SN_i}$$
$$\text{AS}_i = \frac{NSP_i + NSN_i}{2}$$

Büyük $\text{AS}_i$ → üst sıra.

---

#### 2.A.4 CODAS

**Dayanak:** Keshavarz-Ghorabaee et al. (2016b)

**Akış:**

Adım 1. Normalizasyon (MEREC tipi) ve ağırlıklı matris $v_{ij}$.

Adım 2. Negatif ideal çözüm: $\text{NIS}_j = \min_i v_{ij}$

Adım 3. Öklid ve Manhattan uzaklıkları (tüm $i$):
$$E_i = \sqrt{\sum_{j=1}^{n}(v_{ij} - \text{NIS}_j)^2}, \quad T_i = \sum_{j=1}^{n}|v_{ij} - \text{NIS}_j|$$

Adım 4. Eşik fonksiyonu:
$$\Psi(x) = \begin{cases}1 & |x| \geq \tau \\ 0 & |x| < \tau\end{cases}, \quad \tau = 0{,}02$$

Adım 5. Değerlendirme skoru:
$$h_i = \sum_{k=1}^{m}\left[(E_i - E_k) + \Psi(E_i - E_k)(T_i - T_k)\right]$$

Büyük $h_i$ → üst sıra.

---

#### 2.A.5 COPRAS

**Dayanak:** Zavadskas & Kaklauskas (1996)

**Akış:**

Adım 1. Sütun toplamı normalizasyonu:
$$\bar{x}_{ij} = \frac{x_{ij}}{\displaystyle\sum_{i=1}^{m} x_{ij} + \varepsilon}$$

Adım 2. Ağırlıklı normalize matris: $\hat{x}_{ij} = w_j \bar{x}_{ij}$

Adım 3. Fayda ve maliyet toplamları:
$$S_i^+ = \sum_{j: c_j = \text{max}} \hat{x}_{ij}, \quad S_i^- = \sum_{j: c_j = \text{min}} \hat{x}_{ij}$$

Adım 4. Göreli önem:
$$Q_i = S_i^+ + \frac{\min_i S_i^- \cdot \displaystyle\sum_{i=1}^{m} S_i^-}{S_i^- \cdot \displaystyle\sum_{i=1}^{m}\frac{\min_i S_i^-}{S_i^-}}$$

Büyük $Q_i$ → üst sıra.

---

#### 2.A.6 OCRA

**Dayanak:** Parkan (1994)

**Akış:**

Adım 1. Fayda bileşeni (tüm $i$):
$$I_i = \sum_{j: c_j = \text{max}} w_j \frac{x_{ij} - \min_i x_{ij}}{\max_i x_{ij} - \min_i x_{ij} + \varepsilon}$$

Adım 2. Maliyet bileşeni:
$$O_i = \sum_{j: c_j = \text{min}} w_j \frac{\max_i x_{ij} - x_{ij}}{\max_i x_{ij} - \min_i x_{ij} + \varepsilon}$$

Adım 3. Toplam skor:
$$\text{OCRA}_i = I_i + O_i$$

---

#### 2.A.7 ARAS

**Dayanak:** Zavadskas & Turskis (2010)

**Akış:**

Adım 1. İdeal alternatif $A_0$ oluşturulur:
$$x_{0j} = \begin{cases}\max_i x_{ij} & c_j = \text{max}\\\min_i x_{ij} & c_j = \text{min}\end{cases}$$

Adım 2. Genişletilmiş matrise ($A_0$ dahil) sum normalizasyonu.

Adım 3. Ağırlıklı normalize matris: $\hat{x}_{ij} = w_j r_{ij}$

Adım 4. Optimality fonksiyon değeri:
$$S_i = \sum_{j=1}^{n} \hat{x}_{ij}$$

Adım 5. Fayda derecesi:
$$K_i = \frac{S_i}{S_0}$$

Büyük $K_i$ → üst sıra.

---

#### 2.A.8 SAW

**Dayanak:** MacCrimmon (1968)

**Akış:**

Adım 1. Sum normalizasyonu.

Adım 2. Ağırlıklı toplam:
$$\text{SAW}_i = \sum_{j=1}^{n} w_j r_{ij}$$

---

#### 2.A.9 WPM

**Dayanak:** Miller & Starr (1969)

**Akış:**

Adım 1. Min-max normalizasyonu.

Adım 2. Ağırlıklı çarpımsal:
$$\text{WPM}_i = \prod_{j=1}^{n}(r_{ij} + \varepsilon)^{w_j}$$

---

#### 2.A.10 MAUT

**Dayanak:** Keeney & Raiffa (1976)

**Akış:**

Adım 1. Min-max normalizasyonu (fayda fonksiyonu olarak doğrusal).

Adım 2. Toplam beklenen fayda:
$$U_i = \sum_{j=1}^{n} w_j r_{ij}$$

---

#### 2.A.11 WASPAS

**Dayanak:** Zavadskas et al. (2012)

**Akış:**

Adım 1. Min-max normalizasyonu.

Adım 2. WSM skoru:
$$Q_i^{(1)} = \sum_{j=1}^{n} w_j r_{ij}$$

Adım 3. WPM skoru:
$$Q_i^{(2)} = \prod_{j=1}^{n}(r_{ij} + \varepsilon)^{w_j}$$

Adım 4. Bütünleşik skor:
$$\text{WASPAS}_i = \lambda Q_i^{(1)} + (1 - \lambda) Q_i^{(2)}, \quad \lambda \in [0,1]$$

---

#### 2.A.12 MOORA

**Dayanak:** Brauers & Zavadskas (2006)

**Akış:**

Adım 1. Vektör normalizasyonu.

Adım 2. Net skor:
$$y_i = \sum_{j: c_j = \text{max}} w_j r_{ij} - \sum_{j: c_j = \text{min}} w_j r_{ij}$$

---

#### 2.A.13 MULTIMOORA

**Dayanak:** Brauers & Zavadskas (2010)

**Akış:**

Vektör normalizasyonu ardından üç bileşen hesaplanır.

Bileşen 1 — Oran sistemi:
$$\text{RS}_i = \sum_{j: \text{max}} w_j r_{ij} - \sum_{j: \text{min}} w_j r_{ij}$$

Bileşen 2 — Referans nokta (tüm $j$ için $r_j^* = $ ideal):
$$\text{RP}_i = \max_{j}\left(w_j |r_{ij} - r_j^*|\right)$$

Bileşen 3 — Tam çarpımsal form:
$$\text{FMF}_i = \frac{\displaystyle\prod_{j: \text{max}}(r_{ij} + \varepsilon)^{w_j}}{\displaystyle\prod_{j: \text{min}}(r_{ij} + \varepsilon)^{w_j}}$$

Borda toplamı (düşük = iyi):
$$B_i = \text{Sıra}(\text{RS}_i) + \text{Sıra}(\text{RP}_i) + \text{Sıra}(\text{FMF}_i)$$

---

#### 2.A.14 MABAC

**Dayanak:** Pamučar & Ćirović (2015)

**Akış:**

Adım 1. Min-max normalizasyonu.

Adım 2. Ağırlıklı matris:
$$v_{ij} = w_j(r_{ij} + 1)$$

Adım 3. Sınır yaklaşım alanı:
$$g_j = \left(\prod_{i=1}^{m} v_{ij}\right)^{1/m}$$

Adım 4. Uzaklık matrisi ve skor:
$$q_{ij} = v_{ij} - g_j, \quad \text{MABAC}_i = \sum_{j=1}^{n} q_{ij}$$

---

#### 2.A.15 MARCOS

**Dayanak:** Stević et al. (2020)

**Akış:**

Adım 1. Anti-ideal ($\text{AAI}$) ve ideal ($\text{AI}$) nesneler oluşturulur.

Adım 2. Genişletilmiş matris normalize edilir:
$$r_{ij} = \frac{x_{ij}}{x_{0j}^{AI}}\ (\text{fayda}), \quad r_{ij} = \frac{x_{0j}^{AI}}{x_{ij}}\ (\text{maliyet})$$

Adım 3. Ağırlıklı skor: $S_i = \sum_j w_j r_{ij}$

Adım 4. Fayda dereceleri ve yararlılık fonksiyonu:
$$k_i^- = \frac{S_i}{S_{\text{AAI}}}, \quad k_i^+ = \frac{S_i}{S_{\text{AI}}}$$
$$f(k_i^+) = \frac{k_i^+}{k_i^+ + k_i^-}, \quad f(k_i^-) = \frac{k_i^-}{k_i^+ + k_i^-}$$
$$U_i = \frac{k_i^+ + k_i^-}{1 + \frac{1 - f(k_i^+)}{f(k_i^+)} + \frac{1 - f(k_i^-)}{f(k_i^-)}}$$

---

#### 2.A.16 CoCoSo

**Dayanak:** Yazdani et al. (2019)

**Akış:**

Adım 1. Min-max normalizasyonu.

Adım 2. Toplamsal ve üstel skor:
$$S_i = \sum_{j=1}^{n} w_j r_{ij}, \quad P_i = \sum_{j=1}^{n}(r_{ij} + \varepsilon)^{w_j}$$

Adım 3. Üç bileşik strateji:
$$K_i^a = \frac{S_i + P_i}{\displaystyle\sum_{i=1}^{m}(S_i + P_i)}, \quad
K_i^b = \frac{S_i}{\displaystyle\sum S_i} + \frac{P_i}{\displaystyle\sum P_i}$$
$$K_i^c = \frac{\lambda S_i + (1-\lambda)P_i}{\lambda \max_i S_i + (1-\lambda)\max_i P_i}$$

Adım 4. Bütünleşik skor:
$$\text{CoCoSo}_i = \frac{\sqrt[3]{K_i^a K_i^b K_i^c} + \frac{1}{3}(K_i^a + K_i^b + K_i^c)}{1}$$

---

#### 2.A.17 PROMETHEE II

**Dayanak:** Brans & Vincke (1985)

**Akış:**

Adım 1. Tüm ikili çiftler $(a, b)$ için kriter farkı ($\text{dir}_j = +1$ fayda, $-1$ maliyet):
$$d_j(a,b) = (x_{aj} - x_{bj}) \cdot \text{dir}_j$$

Adım 2. Tercih fonksiyonu (lineer varsayılan, $q$ ve $p$ parametreleri):
$$P_j(a,b) = \begin{cases}0 & d_j \leq q \\ \frac{d_j - q}{p - q} & q < d_j < p \\ 1 & d_j \geq p\end{cases}$$

Adım 3. Ağırlıklı ağ akışları:
$$\pi(a,b) = \sum_{j=1}^{n} w_j P_j(a,b)$$
$$\phi^+(a) = \frac{1}{m-1}\sum_{b \neq a}\pi(a,b), \quad
\phi^-(a) = \frac{1}{m-1}\sum_{b \neq a}\pi(b,a)$$
$$\phi(a) = \phi^+(a) - \phi^-(a)$$

Büyük $\phi(a)$ → üst sıra.

---

#### 2.A.18 GRA (Grey Relational Analysis)

**Dayanak:** Deng (1989)

**Akış:**

Adım 1. Min-max normalizasyonu; referans dizi $r_j^* = 1$ (tüm $j$).

Adım 2. Mutlak fark:
$$\Delta_{ij} = |r_j^* - r_{ij}| = |1 - r_{ij}|$$

Adım 3. Gri ilişkisel katsayı ($\rho = 0{,}5$ varsayılan):
$$\xi_{ij} = \frac{\Delta_{\min} + \rho \Delta_{\max}}{\Delta_{ij} + \rho \Delta_{\max}},\quad
\Delta_{\min} = \min_{i,j}\Delta_{ij},\quad \Delta_{\max} = \max_{i,j}\Delta_{ij}$$

Adım 4. İlişkisellik derecesi:
$$\Gamma_i = \sum_{j=1}^{n} w_j \xi_{ij}$$

---

#### 2.A.19 SPOTIS

**Dayanak:** Dezert et al. (2020)

**Akış:**

Adım 1. Sabit ideal nokta belirlenir:
$$d_j^* = \begin{cases}\max_i x_{ij} & c_j = \text{max}\\\min_i x_{ij} & c_j = \text{min}\end{cases}$$

Adım 2. Normalleştirilmiş uzaklık:
$$\text{SPOTIS}_i = \sum_{j=1}^{n} w_j \frac{|x_{ij} - d_j^*|}{\max_i x_{ij} - \min_i x_{ij} + \varepsilon}$$

Düşük $\text{SPOTIS}_i$ → üst sıra. İdeal nokta yeni alternatif eklenmesiyle değişmez.

---

#### 2.A.20 RAWEC

**Dayanak:** Sotoudeh-Anvari (2023)

**Akış:**

Adım 1. Sum normalizasyonu; ağırlıklı matris $v_{ij} = w_j \bar{x}_{ij}$.

Adım 2. Sütun bazlı sıralama matrisi $\text{rk}_{ij}$ (her kriter için ayrı, büyük değer 1. sıra).

Adım 3. Ağırlıklı harmonik sıra toplamı:
$$\text{RAWEC}_i = \sum_{j=1}^{n} \frac{w_j}{\text{rk}_{ij}}$$

---

#### 2.A.21 RAFSI

**Dayanak:** Žižović et al. (2020)

**Akış:**

Adım 1. Sabit referans aralığına $[r_1, r_2] = [1, 9]$ doğrusal eşleme:
$$h_{ij} = r_1 + (r_2 - r_1)\frac{x_{ij} - \min_i x_{ij}}{\max_i x_{ij} - \min_i x_{ij} + \varepsilon}\ (\text{fayda})$$
$$h_{ij} = r_1 + (r_2 - r_1)\frac{\max_i x_{ij} - x_{ij}}{\max_i x_{ij} - \min_i x_{ij} + \varepsilon}\ (\text{maliyet})$$

Adım 2. Ağırlıklı toplam:
$$\text{RAFSI}_i = \sum_{j=1}^{n} w_j h_{ij}$$

---

#### 2.A.22 ROV

**Dayanak:** Yakowitz et al. (1993)

**Akış:**

Adım 1. Min-max normalizasyonu.

Adım 2. Ağırlıklı toplam:
$$\text{ROV}_i = \sum_{j=1}^{n} w_j r_{ij}$$

---

#### 2.A.23 AROMAN

**Dayanak:** Dimitrijević et al. (2022)

**Akış:**

Adım 1. Sum normalizasyonu $r_{ij}^{(1)}$ ve min-max normalizasyonu $r_{ij}^{(2)}$.

Adım 2. Geometrik bileşim:
$$h_{ij} = \sqrt{r_{ij}^{(1)} \cdot r_{ij}^{(2)} + \varepsilon} - \sqrt{\varepsilon}$$

Adım 3. Ağırlıklı toplam:
$$\text{AROMAN}_i = \sum_{j=1}^{n} w_j h_{ij}$$

---

#### 2.A.24 DNMA

**Dayanak:** Liu & Zhu (2021)

**Akış:**

Adım 1. Min-max normalizasyonu $r_{ij}^{(1)}$ ve sum normalizasyonu $r_{ij}^{(2)}$.

Adım 2. Ayrı ağırlıklı toplamlar:
$$S_i^{(1)} = \sum_{j} w_j r_{ij}^{(1)}, \quad S_i^{(2)} = \sum_{j} w_j r_{ij}^{(2)}$$

Adım 3. Bütünleşik skor ($\alpha = 0{,}5$ varsayılan):
$$\text{DNMA}_i = \alpha S_i^{(1)} + (1 - \alpha) S_i^{(2)}$$

---

### 2.B BULANIK SIRALAMA YÖNTEMLERİ

#### Ortak Altyapı

**TFN dönüşümü** ($s = \text{spread}$, varsayılan $0{,}10$):
$$\tilde{x}_{ij} = (l_{ij},\, m_{ij},\, u_{ij}) = \bigl(x_{ij}(1-s),\; x_{ij},\; x_{ij}(1+s)\bigr)$$

**TFN normalizasyonu:**
Fayda ($u_j^* = \max_i u_{ij}$):
$$\tilde{r}_{ij} = \left(\frac{l_{ij}}{u_j^*},\; \frac{m_{ij}}{u_j^*},\; \frac{u_{ij}}{u_j^*}\right)$$
Maliyet ($l_j^* = \min_i l_{ij}$):
$$\tilde{r}_{ij} = \left(\frac{l_j^*}{u_{ij}},\; \frac{l_j^*}{m_{ij}},\; \frac{l_j^*}{l_{ij}}\right)$$

**Vertex uzaklık formülü:**
$$d(\tilde{a}, \tilde{b}) = \sqrt{\frac{(l_a - l_b)^2 + (m_a - m_b)^2 + (u_a - u_b)^2}{3}}$$

**Senaryo sarmalayıcısı** (Fuzzy TOPSIS hariç tüm bulanık yöntemlerde):
$$\text{Skor}_i^{\text{Fuzzy}} = \frac{1}{3}\bigl(\text{Skor}_i^L + \text{Skor}_i^M + \text{Skor}_i^U\bigr)$$

---

#### 2.B.1 Fuzzy TOPSIS

Doğrudan TFN üzerinde çalışır (senaryo sarmalayıcısı kullanmaz).

Adım 1. Ağırlıklı bulanık matris:
$$\tilde{v}_{ij} = w_j \otimes \tilde{r}_{ij} = (w_j l_{ij},\; w_j m_{ij},\; w_j u_{ij})$$

Adım 2. FPIS ve FNIS:
$$\tilde{v}_j^+ = \max_i \tilde{v}_{ij}, \quad \tilde{v}_j^- = \min_i \tilde{v}_{ij}$$

Adım 3. Vertex uzaklıkları ve skor:
$$D_i^+ = \sum_{j=1}^{n} d(\tilde{v}_{ij}, \tilde{v}_j^+), \quad D_i^- = \sum_{j=1}^{n} d(\tilde{v}_{ij}, \tilde{v}_j^-)$$
$$CC_i = \frac{D_i^-}{D_i^+ + D_i^- + \varepsilon}$$

---

#### 2.B.2–2.B.24 Diğer Bulanık Yöntemler

Her yöntem için üç senaryo ($X^L$, $X^M$, $X^U$) oluşturulur.
Her senaryoda klasik versiyonun tüm adımları çalıştırılır.
Skor ortalaması alınarak nihai sıralama elde edilir.

| Yöntem | Sıralama yönü | Temel skor değişkeni |
|---|---|---|
| Fuzzy VIKOR | Düşük = iyi | $Q_i$ |
| Fuzzy EDAS | Yüksek = iyi | $\text{AS}_i$ |
| Fuzzy CODAS | Yüksek = iyi | $h_i$ |
| Fuzzy COPRAS | Yüksek = iyi | $Q_i$ |
| Fuzzy OCRA | Yüksek = iyi | $\text{OCRA}_i$ |
| Fuzzy ARAS | Yüksek = iyi | $K_i$ |
| Fuzzy SAW | Yüksek = iyi | $\text{SAW}_i$ |
| Fuzzy WPM | Yüksek = iyi | $\text{WPM}_i$ |
| Fuzzy MAUT | Yüksek = iyi | $U_i$ |
| Fuzzy WASPAS | Yüksek = iyi | $\text{WASPAS}_i$ |
| Fuzzy MOORA | Yüksek = iyi | $y_i$ |
| Fuzzy MULTIMOORA | Düşük = iyi | $B_i$ |
| Fuzzy MABAC | Yüksek = iyi | $\text{MABAC}_i$ |
| Fuzzy MARCOS | Yüksek = iyi | $U_i$ |
| Fuzzy CoCoSo | Yüksek = iyi | $\text{CoCoSo}_i$ |
| Fuzzy PROMETHEE | Yüksek = iyi | $\phi_i$ |
| Fuzzy GRA | Yüksek = iyi | $\Gamma_i$ |
| Fuzzy SPOTIS | Düşük = iyi | $\text{SPOTIS}_i$ |
| Fuzzy RAWEC | Yüksek = iyi | $\text{RAWEC}_i$ |
| Fuzzy RAFSI | Yüksek = iyi | $\text{RAFSI}_i$ |
| Fuzzy ROV | Yüksek = iyi | $\text{ROV}_i$ |
| Fuzzy AROMAN | Yüksek = iyi | $\text{AROMAN}_i$ |
| Fuzzy DNMA | Yüksek = iyi | $\text{DNMA}_i$ |

---

## BÖLÜM 3 — SAĞLAMLIK VE DAYANIKLILIK ANALİZİ

### 3.1 Lokal Duyarlılık Analizi

**Amaç:** Tek bir ağırlığın değiştirilmesinin sıralama üzerindeki etkisini ölçmek.

**Mekanizma:**

Baz ağırlık vektörü $\mathbf{w}^0 = (w_1^0, \ldots, w_n^0)$ için
$j$. kritere $\delta \in \{-0{,}20,\, -0{,}10,\, +0{,}10,\, +0{,}20\}$ pertürbasyonu:

$$w_j^{(\delta)} = \max(\varepsilon,\; w_j^0(1 + \delta))$$

Yeni ağırlık vektörü normalize edilir:
$$\tilde{w}_k^{(\delta)} = \frac{w_k^{(\delta)}}{\displaystyle\sum_{k=1}^{n} w_k^{(\delta)}}$$

Pertürbe edilmiş ağırlıklarla sıralama yeniden hesaplanır.
Baz sıralama $\mathbf{r}^0$ ile pertürbe sıralama $\mathbf{r}^{(\delta)}$ arasında
Spearman sıra korelasyonu:
$$\rho_s = 1 - \frac{6\displaystyle\sum_{i=1}^{m}(r_i^0 - r_i^{(\delta)})^2}{m(m^2 - 1)}$$

Toplam lokal senaryo sayısı: $n \times 4$ (her kriter için 4 pertürbasyon düzeyi).

---

### 3.2 Monte Carlo Simülasyonu

**Amaç:** Ağırlık vektörünün tüm olası pertürbasyonları altında sıralama kararlılığını istatistiksel olarak ölçmek.

**Mekanizma:**

$N$ iterasyon (varsayılan $N = 200$) için:

Adım 1. Log-normal gürültü çekimi:
$$\boldsymbol{\varepsilon}^{(t)} \sim \mathcal{N}(\mathbf{0},\, \sigma^2 \mathbf{I}), \quad \sigma = 0{,}12\ (\text{varsayılan})$$

Adım 2. Pertürbe edilmiş ağırlıklar (lognormal çarpan):
$$\tilde{w}_j^{(t)} = w_j^0 \cdot e^{\varepsilon_j^{(t)}}$$

Adım 3. Normalize:
$$w_j^{(t)} = \frac{\tilde{w}_j^{(t)}}{\displaystyle\sum_{k=1}^{n} \tilde{w}_k^{(t)}}$$

Adım 4. Sıralama hesaplanır; lider alternatif kaydedilir.

**Çıktı istatistikleri:**

Birincililik oranı (Top-1 kararlılık):
$$\text{Stab}(A_i) = \frac{\#\{t : A_i \text{ birinci sırada}\}}{N}$$

Ortalama sıra:
$$\bar{r}_i = \frac{1}{N}\sum_{t=1}^{N} r_i^{(t)}$$

**Yorumlama eşikleri:**
- $\text{Stab} \geq 0{,}80$: Yüksek kararlılık — sonuç güvenle raporlanabilir.
- $0{,}60 \leq \text{Stab} < 0{,}80$: Orta kararlılık — duyarlılık tablosuyla birlikte sunulmalı.
- $\text{Stab} < 0{,}60$: Düşük kararlılık — kesin öneri yapılmamalı.

---

### 3.3 Yöntem Karşılaştırması (Spearman Uyum Analizi)

**Amaç:** Farklı sıralama yöntemlerinin aynı veri üzerinde ne kadar tutarlı sonuç ürettiğini ölçmek.

**Mekanizma:**

$K$ yöntemin her biri için sıralama vektörü $\mathbf{r}^{(k)}$ hesaplanır.

İkili Spearman korelasyonu:
$$\rho_s^{(k,l)} = 1 - \frac{6\displaystyle\sum_{i=1}^{m}(r_i^{(k)} - r_i^{(l)})^2}{m(m^2 - 1)}$$

Ortalama uyum:
$$\bar{\rho} = \frac{2}{K(K-1)}\sum_{k < l} \rho_s^{(k,l)}$$

Minimum uyum:
$$\rho_{\min} = \min_{k < l} \rho_s^{(k,l)}$$

**Yorumlama:**
- $\bar{\rho} \geq 0{,}85$: Yüksek metodolojik uyum — bulgular yöntem seçiminden bağımsız.
- $0{,}70 \leq \bar{\rho} < 0{,}85$: Orta uyum — yöntem gerekçesi açıklanmalı.
- $\bar{\rho} < 0{,}70$: Düşük uyum — uzlaşı odaklı yöntem tercih edilmeli.

---

### 3.4 Karar Güven Skoru

**Amaç:** Tüm sağlamlık sinyallerini tek bir güven düzeyine ($\text{high}$, $\text{medium}$, $\text{low}$) bütünleştirmek.

**Mekanizma:**

Dört sinyal değerlendirilir:

1. Yöntem uyumu: $\bar{\rho} \geq 0{,}85 \to +1$
2. MC kararlılık: $\text{Stab} \geq 0{,}80 \to +1$
3. Göreli ayrışma:
$$\text{RelGap} = \frac{|\text{Skor}_1 - \text{Skor}_2|}{\text{SkorAralığı} + \varepsilon} \geq 0{,}15 \to +1$$
4. VIKOR uzlaşı koşulları (yalnızca VIKOR seçiliyse): C1 ∧ C2 $\to +1$

Güven düzeyi:
- $\sum \geq 2$ pozitif sinyal ve hiç düşük sinyal yok $\to$ **Yüksek**
- $\sum = 1$ pozitif sinyal veya karışık $\to$ **Orta**
- Herhangi bir düşük sinyal $\to$ **Düşük**

---

## BÖLÜM 4 — IMRAD MAKALE ŞABLONU

Bu bölüm dinamik olarak doldurulur. Köşeli parantez içindeki ifadeler
kullanıcının seçimine göre sistem tarafından otomatik yerleştirilir.

---

### BAŞLIK

[AĞIRLIK_YÖNTEMİ] ve [SIRALAMA_YÖNTEMİ] Entegrasyonuyla [ALAN]'da [PROBLEM] İçin Çok Kriterli Karar Analizi

---

### ÖZET

Bu çalışmada [ALAN] alanında [PROBLEM] için [ALTERNATİF_SAYISI] alternatif ve [KRİTER_SAYISI] kriter içeren bir karar matrisi oluşturulmuş; kriter ağırlıkları [AĞIRLIK_YÖNTEMİ] yöntemiyle belirlenmiş ve [SIRALAMA_YÖNTEMİ] ile alternatifler sıralanmıştır. Sağlamlık [MONTE_CARLO_N] iterasyonlu Monte Carlo simülasyonu ve lokal duyarlılık analizi ile doğrulanmıştır. Sonuçlar, [LİDER_ALTERNATİF] alternatifinin %[STABILIITE_ORANI] oranında kararlı biçimde üst sırada yer aldığını ve [ORTALAMA_SPEARMAN] ortalama Spearman uyum katsayısının yüksek metodolojik tutarlılığa işaret ettiğini ortaya koymaktadır.

**Anahtar Kelimeler:** Çok Kriterli Karar Verme, [AĞIRLIK_YÖNTEMİ], [SIRALAMA_YÖNTEMİ], Sağlamlık Analizi, [ALAN]

---

### 1. GİRİŞ

[ALAN] alanında karar vericiler, çoğunlukla birbiriyle çelişen kriterler altında en uygun alternatifi seçmek durumundadır. Bu tür problemlerin çözümünde Çok Kriterli Karar Verme (ÇKKV) yöntemleri yaygın biçimde kullanılmakta; hem kriter ağırlıklarının nesnel biçimde belirlenmesi hem de alternatiflerin sistematik sıralanması mümkün kılınmaktadır.

Literatürde [ALAN] alanına yönelik ÇKKV çalışmaları incelendiğinde [LİTERATÜR_ÖZETİ]. Mevcut çalışmalar [ARALIK_YIL] dönemini kapsamakta; ancak [ARAŞTIRMA_BOŞLUĞU] konusundaki eksiklik dikkat çekmektedir.

Bu çalışmada [AĞIRLIK_YÖNTEMİ] ile objektif/subjektif ağırlıklandırma ve [SIRALAMA_YÖNTEMİ] ile sıralama birleştirilerek [PROBLEM] ele alınmaktadır. Çalışmanın [ALAN] literatürüne katkısı: (i) [KATKI_1], (ii) [KATKI_2], (iii) kapsamlı sağlamlık analizi.

---

### 2. YÖNTEM

#### 2.1 Karar Matrisi

Bu çalışmada [ALTERNATİF_SAYISI] alternatif ($A_1, \ldots, A_m$) ve [KRİTER_SAYISI] kriter ($C_1, \ldots, C_n$) dikkate alınmıştır. Kriterler [VERİ_KAYNAĞI] kaynağından elde edilmiştir. Karar matrisi $X_{m \times n}$ olarak formüle edilmiş; kriter yönleri [KRİTER_YÖNLERİ] şeklinde belirlenmiştir.

#### 2.2 Kriter Ağırlıklarının Belirlenmesi

##### [AĞIRLIK_YÖNTEMİ == Entropy ise]
Kriter ağırlıkları Shannon bilgi entropisi yöntemiyle (Zeleny, 1982; Wang & Lee, 2009) belirlenmiştir. Yöntem, kriter içi düzensizlik azaldıkça ayırt ediciliğin arttığı ilkesine dayanmakta ve veri üzerinden otomatik ağırlık üretmektedir.

Ham karar matrisi önce sum normalizasyonuna tabi tutularak olasılık dağılımı $p_{ij}$ elde edilmiştir. Her $j$ kriteri için Shannon entropisi:
$$H_j = -k \sum_{i=1}^{m} p_{ij} \ln(p_{ij}), \quad k = \frac{1}{\ln(m)}$$
formülüyle hesaplanmıştır. Sapma $d_j = 1 - H_j$ ve ağırlık $w_j = d_j / \sum d_j$ olarak türetilmiştir.

##### [AĞIRLIK_YÖNTEMİ == CRITIC ise]
Kriter ağırlıkları CRITIC yöntemiyle (Diakoulaki et al., 1995) belirlenmiştir. Yöntem hem kriter varyansını hem de kriterler arası çatışma yapısını bilgi içeriğine dönüştürmektedir.

Bilgi içeriği: $C_j = \sigma_j \sum_{k=1}^{n}(1 - \rho_{jk})$; ağırlık: $w_j = C_j / \sum C_j$.

##### [AĞIRLIK_YÖNTEMİ == AHP ise]
Kriter ağırlıkları Analitik Hiyerarşi Süreci (AHP; Saaty, 1980) yöntemiyle belirlenmiştir. [UZMAN_BİLGİSİ] tarafından Saaty 1–9 ölçeğinde doldurulan ikili karşılaştırma matrisi üzerinden ana özdeğer $\lambda_{\max}$ hesaplanmış, ana özvektör normalize edilerek ağırlık vektörü elde edilmiştir. Tutarlılık oranı $CR = CI/RI = [CR_DEĞERİ] < 0{,}10$ olup kabul edilebilir düzeydedir.

##### [AĞIRLIK_YÖNTEMİ == BWM ise]
Kriter ağırlıkları En İyi-En Kötü Yöntemi (BWM; Rezaei, 2015) ile belirlenmiştir. En iyi kriter [EN_İYİ_KRİTER], en kötü kriter [EN_KÖTÜ_KRİTER] olarak belirlenmiş; SLSQP optimizasyon modeli çözülmüş ve $CR = [CR_DEĞERİ] < 0{,}10$ tutarlılık düzeyinde $\xi^* = [KSİ_DEĞERİ]$ elde edilmiştir.

#### 2.3 Alternatif Sıralaması

##### [SIRALAMA_YÖNTEMİ == Fuzzy TOPSIS ise]
Alternatiflerin sıralanmasında Bulanık TOPSIS yöntemi (Chen, 2000) kullanılmıştır. Yöntem, belirsizliği üçgensel bulanık sayılarla (TFN) modelleyerek pozitif ve negatif ideal çözümlere yakınlık katsayısı üzerinden sıralama yapmaktadır.

Crisp karar matrisi $s = [SPREAD]$ spread parametresiyle TFN'ye dönüştürülmüştür:
$$\tilde{x}_{ij} = (x_{ij}(1-s),\; x_{ij},\; x_{ij}(1+s))$$

TFN normalizasyonu, ağırlıklı bulanık matris ve vertex uzaklık formülü uygulanarak yakınlık katsayısı $CC_i \in [0,1]$ hesaplanmıştır. Yüksek $CC_i$ değeri pozitif ideale yakınlığı ifade etmektedir.

#### 2.4 Sağlamlık Analizi

Sonuçların güvenilirliği iki aşamalı sağlamlık analiziyle sınanmıştır.

**Lokal duyarlılık:** Her kriter ağırlığına $\delta \in \{-0{,}20, -0{,}10, +0{,}10, +0{,}20\}$ pertürbasyonu uygulanmış; her senaryoda Spearman sıra korelasyonu $\rho_s$ hesaplanmıştır.

**Monte Carlo simülasyonu:** $N = [MC_N]$ iterasyon boyunca ağırlık vektörüne $\sigma = [SIGMA]$ standart sapmalı log-normal gürültü enjekte edilmiş; lider alternatifin birincililik oranı $\text{Stab}([LİDER]) = [STABILIITE_ORANI]$ olarak belirlenmiştir.

**Yöntem karşılaştırması:** [KARŞILAŞTIRMA_YÖNTEMLERİ] yöntemleriyle çapraz doğrulama yapılmış; ortalama Spearman uyum katsayısı $\bar{\rho} = [ORTALAMA_SPEARMAN]$ olarak hesaplanmıştır.

---

### 3. BULGULAR

#### 3.1 Kriter Ağırlıkları

[AĞIRLIK_TABLOSU] — En yüksek ağırlık $w_{[EN_AGIRLI_KRİTER]} = [EN_YÜKSEK_AGIRLIK]$ ile [EN_AGIRLI_KRİTER] kriterine aittir. Bu bulgu [ALAN] bağlamında [YORUM] anlamına gelmektedir.

#### 3.2 Sıralama Sonuçları

[SIRALAMA_TABLOSU] — Analiz sonucunda alternatiflerin sıralaması $[LİDER] \succ [İKİNCİ] \succ \ldots$ şeklinde belirlenmiştir. Lider alternatif $CC_{[LİDER]} = [CC_DEĞERİ]$ yakınlık katsayısıyla birinci sıraya yerleşmiştir.

#### 3.3 Sağlamlık Bulguları

Lokal duyarlılık analizi tüm $n \times 4 = [TOPLAM_SENARYO]$ senaryoda Spearman $\rho_s \geq [MIN_SPEARMAN]$ değeri üretmiş; sıralamada değişim gözlemlenmemiştir / [KAÇ] senaryoda kısmi değişim gözlemlenmiştir.

Monte Carlo simülasyonu $\text{Stab}([LİDER]) = [STABILIITE_ORANI]$ birincilik oranıyla lider alternatifin kararlılığını [YÜKSEK/ORTA/DÜŞÜK] düzeyde teyit etmiştir. Ortalama sıra $\bar{r}_{[LİDER]} = [ORTALAMA_SIRA]$ olarak hesaplanmıştır.

Yöntem karşılaştırmasında $\bar{\rho} = [ORTALAMA_SPEARMAN]$ değeri metodolojik tutarlılığın [YÜKSEK/ORTA/DÜŞÜK] düzeyde olduğunu göstermektedir.

---

### 4. TARTIŞMA

Elde edilen bulgular [ALAN] literatürüyle karşılaştırıldığında [LİTERATÜR_KARŞILAŞTIRMA]. [LİDER_ALTERNATİF]'in üstünlüğü [AÇIKLAMA] bağlamında değerlendirilebilir.

[AĞIRLIK_YÖNTEMİ]'nin seçimi [GEREKÇE]. Benzer problemlerde [ALTERNATİF_YÖNTEMLER] ile karşılaştırmalı analiz yapılması önerilmektedir.

Çalışmanın kısıtları: (i) karar matrisi [VERİ_DÖNEMİ] dönemine ait verilerle sınırlıdır; (ii) uzman yargısı [SUBJEKTIF_ELEMAN] içermektedir; (iii) kriter seti [KRİTER_SINIRLAMASI].

---

### 5. SONUÇ

Bu çalışmada [ALAN] alanındaki [PROBLEM] için [AĞIRLIK_YÖNTEMİ] ve [SIRALAMA_YÖNTEMİ] entegrasyonuyla çok kriterli bir değerlendirme çerçevesi sunulmuştur. [LİDER_ALTERNATİF] alternatifi tüm sağlamlık testlerinde üst sırada yer almış ve karar vericiler için güvenilir bir referans noktası oluşturmuştur. Gelecek araştırmalarda [GELECEK_ÇALIŞMA] önerilmektedir.

---

### KAYNAKLAR

*(Seçilen yönteme göre dinamik olarak eklenir)*

- Saaty, T.L. (1980). *The Analytic Hierarchy Process*. McGraw-Hill, New York.
- Rezaei, J. (2015). Best-worst multi-criteria decision-making method. *Omega*, 53, 49–57.
- Keršuliene, V., Zavadskas, E.K., & Turskis, Z. (2010). Selection of rational dispute resolution method. *Journal of Business Economics and Management*, 11(2), 85–102.
- Fontela, E., & Gabus, A. (1976). *The DEMATEL Observer*. Battelle Geneva Research Center.
- Edwards, W. (1971). Social utilities. *Engineering Economist Summer Symposium Series*, 6, 119–129.
- Zeleny, M. (1982). *Multiple Criteria Decision Making*. McGraw-Hill, New York.
- Diakoulaki, D., Mavrotas, G., & Papayannakis, L. (1995). Determining objective weights in multiple criteria problems. *Computers & Operations Research*, 22(7), 763–770.
- Keshavarz-Ghorabaee, M., et al. (2021). Determination of objective weights using MEREC. *Symmetry*, 13(4), 525.
- Ecer, F., & Pamucar, D. (2022). A novel LOPCOW-DOBI MCDM methodology. *Expert Systems with Applications*, 197, 116567.
- Zavadskas, E.K., & Podvezko, V. (2016). Integrated determination of objective criteria weights. *International Journal of Information Technology & Decision Making*, 15(2), 267–283.
- Hwang, C.L., & Yoon, K. (1981). *Multiple Attribute Decision Making*. Springer, Berlin.
- Opricovic, S. (1998). *Multicriteria Optimization of Civil Engineering Systems*. Faculty of Civil Engineering, Belgrade.
- Keshavarz-Ghorabaee, M., et al. (2016). A new combinative distance-based assessment (CODAS) method. *Informatica*, 27(4), 795–831.
- Zavadskas, E.K., & Kaklauskas, A. (1996). *Determination of an efficient contractor*. Vilnius Gediminas Technical University Press.
- Brauers, W.K.M., & Zavadskas, E.K. (2006). The MOORA method. *Technological and Economic Development of Economy*, 12(2), 85–96.
- Brauers, W.K.M., & Zavadskas, E.K. (2010). Project management by MULTIMOORA. *Technological and Economic Development of Economy*, 16(2), 162–188.
- Pamučar, D., & Ćirović, G. (2015). The selection of transport and handling resources in logistics. *Expert Systems with Applications*, 42(6), 2907–2916.
- Stević, Ž., et al. (2020). Sustainable supplier selection in healthcare industries using MARCOS. *Computers & Industrial Engineering*, 140, 106231.
- Yazdani, M., et al. (2019). A combined compromise solution (CoCoSo) method. *Management Decision*, 57(9), 2501–2519.
- Brans, J.P., & Vincke, Ph. (1985). A preference ranking organisation method. *Management Science*, 31(6), 647–656.
- Deng, J.L. (1989). Introduction to grey system. *Journal of Grey System*, 1(1), 1–24.
- Dezert, J., et al. (2020). The SPOTIS rank reversal free method. *Proceedings of FUSION 2020*.
- Sotoudeh-Anvari, A. (2023). RAWEC method. *Expert Systems with Applications*, 218, 119556.
- Žižović, M., et al. (2020). Objective methods for determining criteria weight coefficients. *Mathematics*, 8(6), 1–19.
- Yakowitz, D.S., et al. (1993). Multicriteria analysis with applications to sustainable development. *European Journal of Operational Research*, 71(2), 177–185.
- Dimitrijević, S., et al. (2022). AROMAN. *Symmetry*, 14(11), 2355.
- Liu, P., & Zhu, B. (2021). DNMA method. *Expert Systems with Applications*, 164, 113973.
- Chen, C.T. (2000). Extensions of the TOPSIS for group decision-making under fuzzy environment. *Fuzzy Sets and Systems*, 114(1), 1–9.
