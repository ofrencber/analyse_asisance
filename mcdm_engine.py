
from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Sequence, Tuple

import numpy as np
import pandas as pd

EPS = 1e-12


OBJECTIVE_WEIGHT_METHODS = [
    "Entropy",
    "CRITIC",
    "Standart Sapma",
    "MEREC",
    "LOPCOW",
    "PCA",
    "CILOS",
    "IDOCRIW",
    "Fuzzy IDOCRIW",
]

CLASSICAL_MCDM_METHODS = [
    "TOPSIS",
    "VIKOR",
    "EDAS",
    "CODAS",
    "COPRAS",
    "OCRA",
    "ARAS",
    "SAW",
    "WPM",
    "MAUT",
    "WASPAS",
    "MOORA",
    "MULTIMOORA",
    "MABAC",
    "MARCOS",
    "CoCoSo",
    "PROMETHEE",
    "GRA",
    "SPOTIS",
    "RAWEC",
    "RAFSI",
    "ROV",
    "AROMAN",
    "DNMA",
]

FUZZY_MCDM_METHODS = [
    "Fuzzy TOPSIS",
    "Fuzzy VIKOR",
    "Fuzzy ARAS",
    "Fuzzy SAW",
    "Fuzzy WPM",
    "Fuzzy MAUT",
    "Fuzzy WASPAS",
    "Fuzzy EDAS",
    "Fuzzy CODAS",
    "Fuzzy COPRAS",
    "Fuzzy OCRA",
    "Fuzzy MOORA",
    "Fuzzy MULTIMOORA",
    "Fuzzy MABAC",
    "Fuzzy MARCOS",
    "Fuzzy CoCoSo",
    "Fuzzy GRA",
    "Fuzzy PROMETHEE",
    "Fuzzy SPOTIS",
    "Fuzzy RAWEC",
    "Fuzzy RAFSI",
    "Fuzzy ROV",
    "Fuzzy AROMAN",
    "Fuzzy DNMA",
]

METHOD_PHILOSOPHY: Dict[str, Dict[str, str]] = {
    "Entropy": {
        "simple": "Entropi, alternatifler arasında gerçekten ayırt edici bilgi taşıyan kriterlere daha yüksek ağırlık verir.",
        "academic": "Entropi ağırlıklandırması, Shannon bilgi kuramına dayanır ve kriter içi düzensizlik azaldıkça ayırt ediciliğin arttığını varsayar.",
    },
    "CRITIC": {
        "simple": "CRITIC, hem değişkenliği yüksek hem de diğer kriterlerle çatışan kriterleri daha önemli görür.",
        "academic": "CRITIC, kriter önemini varyans ve kriterler arası çatışma yapısını birlikte nicelleştirerek belirleyen veri temelli bir ağırlıklandırma yaklaşımıdır.",
    },
    "Standart Sapma": {
        "simple": "Standart sapma yöntemi, alternatifleri birbirinden daha fazla ayıran kriterleri öne çıkarır.",
        "academic": "Standart sapma tabanlı objektif ağırlıklandırma, kriter ayırt ediciliğini dağılım genişliği üzerinden ölçer.",
    },
    "MEREC": {
        "simple": "MEREC, bir kriter çıkarıldığında toplam performans ne kadar değişiyorsa o kriteri o kadar önemli kabul eder.",
        "academic": "MEREC, kriterin sistemden çıkarılmasının alternatiflerin bütüncül performansı üzerindeki bozucu etkisini ölçerek ağırlık üretir.",
    },
    "LOPCOW": {
        "simple": "LOPCOW, normalize edilmiş verideki logaritmik yüzde değişim gücüne göre ağırlık hesaplar.",
        "academic": "LOPCOW, kriterlerin normalize edilmiş yapısındaki RMS/standart sapma oranını logaritmik yüzde değişim biçiminde kullanarak ölçek farklarını bastıran objektif bir ağırlıklandırma sunar.",
    },
    "PCA": {
        "simple": "PCA, verideki ortak varyans yapısını ortaya çıkarır ve bilgi taşıyan eksenlerde güçlü yük veren kriterlere daha çok ağırlık verir.",
        "academic": "PCA tabanlı ağırlıklandırma, korelasyon yapısından türetilen temel bileşenlerin açıkladığı varyansı ve yük dağılımlarını kriter ağırlıklarına dönüştürür.",
    },
    "IDOCRIW": {
        "simple": "IDOCRIW, Entropi ile CILOS mantığını birleştirerek hem bilgi çeşitliliğini hem de kriter kayıp etkisini birlikte okur.",
        "academic": "IDOCRIW, entropi temelli bilgi içeriğini CILOS tabanlı göreli etki kaybı yapısıyla bütünleştirerek kriter ağırlıklarını hibrit ve objektif biçimde üretir.",
    },
    "Fuzzy IDOCRIW": {
        "simple": "Fuzzy IDOCRIW, IDOCRIW mantığını belirsizlik altında üç senaryolu bulanık bir çerçevede uygular.",
        "academic": "Fuzzy IDOCRIW, IDOCRIW bileşenlerini bulanıklaştırılmış alt-orta-üst performans senaryoları üzerinde birleştirerek belirsizlik duyarlı objektif ağırlık üretir.",
    },
    "CILOS": {
        "simple": "CILOS, her kriter çıkarıldığında en iyi alternatifin diğer kriterler üzerinde yarattığı kayıp etkisini ölçerek kriterleri ağırlıklandırır.",
        "academic": "CILOS (Criterion Impact LOSs), göreli etki kayıpları matrisini ve doğrusal denklem sistemini kullanarak kriterlerin birbirini ne ölçüde etkilediğini ağırlığa dönüştürür.",
    },
    "SPOTIS": {
        "simple": "SPOTIS, her alternatifi sabit bir ideal noktadan uzaklığına göre sıralar — alternatif listesi değişse bile ideal nokta sabit kalır ve sıralama tersine çevrilmez.",
        "academic": "SPOTIS (Stable Preference Ordering Towards Ideal Solution), sabit ideal referans noktasına ağırlıklı normalleştirilmiş uzaklıkları kullanarak rank reversal problemine yapısal düzeyde direnen bir sıralama yöntemidir.",
    },
    "MULTIMOORA": {
        "simple": "MULTIMOORA, üç farklı MOORA yaklaşımını (oran sistemi, referans noktası, çarpımsal form) Borda sayımıyla birleştirerek tek bir yönteme göre çok daha sağlam bir sıralama üretir.",
        "academic": "MULTIMOORA, oran sistemi, referans nokta ve tam çarpımsal formdan elde edilen sıraları bütünleşik Borda dominans teorisiyle toplulaştıran çok katmanlı, sağlamlık odaklı bir MCDM yöntemidir.",
    },
    "RAWEC": {
        "simple": "RAWEC, her kriter içindeki alternatif sırasını o kriterin ağırlığıyla ağırlıklandırarak nihai sıralamayı üretir; önemli kriterlerde üst sırada olmak belirleyici avantaj sağlar.",
        "academic": "RAWEC (Ranking of Alternatives by Weight of Each Criterion), ağırlıklı normalleştirilmiş karar matrisinde kriter bazlı sıra değerlerini ağırlıklı harmonik toplulaştırmayla birleştiren sıra-ağırlık bileşimli bir sıralama yöntemidir.",
    },
    "RAFSI": {
        "simple": "RAFSI, kriter değerlerini sabit bir referans aralığa [1–9] dönüştürür; yeni alternatif eklenmesi mevcut alternatiflerin skorunu değiştirmez ve sıralama kararlıdır.",
        "academic": "RAFSI (Ranking of Alternatives through Functional mapping of criterion Sub-Intervals), kriter değerlerini ideal ve anti-ideal sınırlı sabit bir referans aralığa doğrusal eşlemeyle taşıyarak rank reversal problemini yapısal olarak önleyen bir sıralama yöntemidir.",
    },
    "ROV": {
        "simple": "ROV, kriter değerlerini gözlenen aralığa göre tam normalize edip ağırlıklı toplamını hesaplar; normalleştirmenin veri aralığına tam oturması yöntemin ayırt ediciliğini artırır.",
        "academic": "ROV (Range of Value), min-max normalleştirilmiş performans değerlerini ağırlıklı toplamsal modelde birleştirerek fayda aralığını tam kullanan objektif bir sıralama yöntemidir.",
    },
    "AROMAN": {
        "simple": "AROMAN, iki farklı normalleştirme yaklaşımını geometrik ortalamayla harmanlayıp tek bir ağırlıklı skor üretir; tek normalleştirmeye göre daha az sapma gösterir.",
        "academic": "AROMAN (Alternative Ranking Order Method Accounting for Two-step Normalization), sum ve min-max normalleştirme adımlarının geometrik bileşimini ağırlıklı toplamla değerlendirerek normalleştirme kaynaklı sapmayı azaltır.",
    },
    "DNMA": {
        "simple": "DNMA, iki ayrı normalleştirme stratejisini eşit ağırlıkla ortalayarak tek bir kararlı skor üretir; hiçbir normalleştirme yaklaşımına tam güvenilmediği durumlarda idealdir.",
        "academic": "DNMA (Double Normalization-based Multiple Aggregation), min-max ve sum normalleştirmelerinden elde edilen ağırlıklı skorları eşit ağırlıklı doğrusal toplulaştırmayla birleştirerek normalleştirme duyarlılığını azaltır.",
    },
    "Fuzzy SPOTIS": {
        "simple": "Fuzzy SPOTIS, belirsizlik altında sabit ideal noktaya uzaklıkları üçgensel bulanık sayılarla değerlendirerek rank reversal direncini korur.",
        "academic": "Fuzzy SPOTIS, SPOTIS'in sabit ideal nokta mimarisini üçgensel bulanıklaştırılmış alt-orta-üst senaryolar üzerinde uygulayarak belirsizlikte yapısal sıralama kararlılığı sağlar.",
    },
    "Fuzzy MULTIMOORA": {
        "simple": "Fuzzy MULTIMOORA, üç MOORA bileşenini belirsizlik altında bulanık sayılarla değerlendirip Borda sayımıyla birleştirerek belirsizlik-sağlamlık dengesini kurar.",
        "academic": "Fuzzy MULTIMOORA, MULTIMOORA'nın oran sistemi, referans nokta ve çarpımsal form bileşenlerini üçgensel bulanık kriter senaryoları altında hesaplayarak belirsizlik duyarlı çok bileşenli sıralama sunar.",
    },
    "Fuzzy RAWEC": {
        "simple": "Fuzzy RAWEC, kriter bazlı sıra toplulaştırmasını belirsizlik altında üçgensel bulanık değerlerle gerçekleştirir.",
        "academic": "Fuzzy RAWEC, RAWEC'in ağırlıklı harmonik sıra toplulaştırmasını üçgensel bulanıklaştırılmış karar matrisinin ortalanmış senaryoları üzerinde uygular.",
    },
    "Fuzzy RAFSI": {
        "simple": "Fuzzy RAFSI, referans aralık eşlemesini bulanık belirsizlik senaryoları altında uygulayarak rank reversal direncini sürdürür.",
        "academic": "Fuzzy RAFSI, RAFSI'nin sabit referans aralığı dönüşümünü üçgensel bulanık kriter değerlerine genişleterek belirsizlikte sıralama kararlılığını korur.",
    },
    "Fuzzy ROV": {
        "simple": "Fuzzy ROV, aralık normalleştirmesini belirsizlik altında üçgensel bulanık değerlerle birleştirir.",
        "academic": "Fuzzy ROV, ROV'un min-max normalleştirme yapısını üçgensel bulanık performans senaryoları üzerinde uygulayarak belirsizliğe duyarlı ağırlıklı fayda sıralaması üretir.",
    },
    "Fuzzy AROMAN": {
        "simple": "Fuzzy AROMAN, çift normalleştirme stratejisini bulanık belirsizlik altında geometrik bileşimle uygular.",
        "academic": "Fuzzy AROMAN, AROMAN'ın iki adımlı normalleştirme mantığını üçgensel bulanık kriter değerleri üzerinde yürüterek belirsizlikte normalleştirme kaynaklı sapmaları sınırlar.",
    },
    "Fuzzy DNMA": {
        "simple": "Fuzzy DNMA, çift normalleştirme yaklaşımını bulanık belirsizlik altında eşit ağırlıkla harmanlayarak skor üretir.",
        "academic": "Fuzzy DNMA, DNMA'nın min-max ve sum normalleştirme bileşenlerini üçgensel bulanık değer senaryoları üzerinde eşit ağırlıklı toplulaştırmayla bütünleştirir.",
    },
    "TOPSIS": {
        "simple": "TOPSIS, en iyi hayali seçeneğe yakın ve en kötü hayali seçeneğe uzak olan alternatifi daha iyi sayar.",
        "academic": "TOPSIS, alternatifleri pozitif ve negatif ideal çözümlere göreli yakınlık katsayıları üzerinden sıralayan uzaklık temelli bir yöntemdir.",
    },
    "VIKOR": {
        "simple": "VIKOR, çoğunluğun faydasını artırırken en büyük bireysel pişmanlığı azaltan uzlaşı çözümünü arar.",
        "academic": "VIKOR, grup faydası ve bireysel pişmanlık ölçülerini uzlaştırarak çatışan kriterler altında uzlaşı çözümü üretir.",
    },
    "EDAS": {
        "simple": "EDAS, her alternatifi kriter ortalamasına göre olumlu ve olumsuz sapmalarıyla değerlendirir.",
        "academic": "EDAS, ortalama çözümden pozitif ve negatif uzaklıkları birlikte dikkate alarak dengeleyici bir değerlendirme yapar.",
    },
    "CODAS": {
        "simple": "CODAS, alternatifleri en kötü noktadan ne kadar uzaklaştıklarına bakarak değerlendirir.",
        "academic": "CODAS, negatif ideal çözümden Öklid uzaklığını ana, Manhattan uzaklığını ise ikincil ayırt edici ölçüt olarak kullanır.",
    },
    "COPRAS": {
        "simple": "COPRAS, fayda ve maliyet katkılarını ayrı izleyip göreli önemle birleştirir.",
        "academic": "COPRAS, ağırlıklı normalize karar matrisinde fayda ve maliyet toplamlarını ayrıştırarak göreli önem düzeyleri üzerinden sıralama üretir.",
    },
    "OCRA": {
        "simple": "OCRA, fayda ve maliyet kriterlerinde göreli rekabet üstünlüğünü birlikte hesaplar.",
        "academic": "OCRA, her kriterdeki göreli performansı fayda ve maliyet bileşenleri altında toplayarak operasyonel rekabet düzeyi temelli sıralama sağlar.",
    },
    "ARAS": {
        "simple": "ARAS, her alternatifi ideal alternatife göre göreli fayda katsayısıyla yorumlar.",
        "academic": "ARAS, normalize ve ağırlıklandırılmış performansları ideal alternatifle karşılaştırarak fayda derecesi temelli sıralama yapar.",
    },
    "SAW": {
        "simple": "SAW, normalleştirilmiş kriter puanlarını ağırlıklı olarak toplar ve en yüksek toplamı en iyi kabul eder.",
        "academic": "SAW, ağırlıklı toplamsal fayda modeli ile kriter performanslarını ortak ölçekte birleştirerek telafi edici sıralama üretir.",
    },
    "WPM": {
        "simple": "WPM, kriterleri ağırlıklarına göre çarpımsal biçimde birleştirir; zayıf boyutları daha güçlü cezalandırır.",
        "academic": "WPM, normalize edilmiş kriter değerlerinin ağırlıklı üsler altında çarpımıyla ölçekten bağımsız ve oran-temelli performans ölçümü sağlar.",
    },
    "MAUT": {
        "simple": "MAUT, her kriteri fayda puanına dönüştürüp ağırlıklı toplam fayda üzerinden karar verir.",
        "academic": "MAUT, doğrusal fayda dönüşümü ve ağırlıklı beklenen fayda yaklaşımıyla çok öznitelikli karar problemlerini rasyonel tercih kuramı zemininde çözer.",
    },
    "WASPAS": {
        "simple": "WASPAS, toplamsal ve çarpımsal fayda mantığını birlikte kullanır.",
        "academic": "WASPAS, WSM ve WPM yapılarını λ parametresi altında hibritleyerek telafi edici ve çarpımsal performans mantığını bütünleştirir.",
    },
    "MOORA": {
        "simple": "MOORA, fayda kriterlerini toplar, maliyet kriterlerini çıkarır ve dengeli bir net skor oluşturur.",
        "academic": "MOORA, oran analizi yaklaşımıyla fayda ve maliyet etkilerini normalize edilmiş ortak ölçekte sentezler.",
    },
    "MABAC": {
        "simple": "MABAC, her alternatifin sınır yaklaşım alanının üstünde mi altında mı kaldığını ölçer.",
        "academic": "MABAC, sınır yaklaşım alanına göre kriter bazlı uzaklıklar üzerinden alternatiflerin konumunu değerlendirir.",
    },
    "MARCOS": {
        "simple": "MARCOS, alternatifi hem ideal hem anti-ideal referansa göre birlikte okur.",
        "academic": "MARCOS, alternatiflerin ideal ve anti-ideal referans nesnelerle ilişkisini fayda dereceleri ve uzlaşı temelli yararlılık fonksiyonlarıyla yorumlar.",
    },
    "CoCoSo": {
        "simple": "CoCoSo, farklı uzlaşı stratejilerini tek bir nihai bileşik skor altında birleştirir.",
        "academic": "CoCoSo, toplamsal ve üstel toplulaştırmaları bir araya getirerek çoklu uzlaşı mantıklarını kombine eder.",
    },
    "PROMETHEE": {
        "simple": "PROMETHEE, alternatifleri ikili karşılaştırmalar üzerinden hangi seçeneğin diğerlerini ne ölçüde aştığına göre sıralar.",
        "academic": "PROMETHEE II, kriter bazlı tercih fonksiyonlarından elde edilen pozitif/negatif akışları net akış altında birleştirerek tam sıralama üretir.",
    },
    "GRA": {
        "simple": "GRA, alternatifi referans profile ne kadar benziyor sorusuna cevap verir.",
        "academic": "GRA, gri ilişkisel katsayılar üzerinden alternatiflerin ideal referans diziye yakınlığını ölçer.",
    },
    "Fuzzy TOPSIS": {
        "simple": "Fuzzy TOPSIS, belirsiz değerleri bulanık sayılarla temsil ederek ideale yakınlığı ölçer.",
        "academic": "Fuzzy TOPSIS, üçgensel bulanık sayılar üzerinde normalize edilmiş ve ağırlıklandırılmış performansları FPIS/FNIS mesafeleriyle değerlendirir.",
    },
    "Fuzzy VIKOR": {
        "simple": "Fuzzy VIKOR, belirsizlik altında uzlaşı çözümü üretir.",
        "academic": "Fuzzy VIKOR, bulanıklaştırılmış kriter performanslarından centroid temelli temsil değerleri türeterek uzlaşı mantığını korur.",
    },
    "Fuzzy ARAS": {
        "simple": "Fuzzy ARAS, bulanık performansların ideal fayda düzeyine göreli oranını ölçer.",
        "academic": "Fuzzy ARAS, üçgensel bulanık sayılarla ifade edilen alternatifleri ideal fayda derecesi bağlamında karşılaştırır.",
    },
    "Fuzzy SAW": {
        "simple": "Fuzzy SAW, bulanık veriden durulaştırılmış puanlarla ağırlıklı toplamsal fayda hesaplar.",
        "academic": "Fuzzy SAW, üçgensel bulanık temsillerden elde edilen fayda değerlerini ağırlıklı toplamsal model altında birleştirerek belirsizlikte sıralama üretir.",
    },
    "Fuzzy WPM": {
        "simple": "Fuzzy WPM, bulanık ortamdaki kriter etkilerini çarpımsal biçimde birleştirir.",
        "academic": "Fuzzy WPM, durulaştırılmış bulanık performansları ağırlıklı çarpımsal fayda modeliyle bütünleştirerek oransal üstünlüğü öne çıkarır.",
    },
    "Fuzzy MAUT": {
        "simple": "Fuzzy MAUT, kriter faydalarını bulanık belirsizlik altında hesaplayıp toplam beklenen faydayı sıralar.",
        "academic": "Fuzzy MAUT, bulanıklaştırılmış kriter gözlemlerinden türetilen fayda fonksiyonlarını ağırlıklı toplulaştırma ile birleştirerek belirsizlikte karar tutarlılığı sağlar.",
    },
    "Fuzzy WASPAS": {
        "simple": "Fuzzy WASPAS, belirsizlik altında hem toplamsal hem çarpımsal değerlendirme yapar.",
        "academic": "Fuzzy WASPAS, bulanıklaştırılmış karar matrisini λ kontrollü toplamsal ve çarpımsal fayda bileşimine dönüştürür.",
    },
    "Fuzzy EDAS": {
        "simple": "Fuzzy EDAS, belirsizlik altında ortalama çözüme göre olumlu/olumsuz sapmaları değerlendirir.",
        "academic": "Fuzzy EDAS, üçgensel bulanık gösterimden türetilen performansları ortalama çözüm uzaklık mantığıyla dengeleyici biçimde sıralar.",
    },
    "Fuzzy CODAS": {
        "simple": "Fuzzy CODAS, bulanık ortamda negatif ideale uzaklık üzerinden alternatifleri ayırt eder.",
        "academic": "Fuzzy CODAS, bulanıklaştırılmış karar matrisini durulaştırılmış uzaklık ölçülerine dönüştürerek negatif idealden ayrışmayı değerlendirir.",
    },
    "Fuzzy COPRAS": {
        "simple": "Fuzzy COPRAS, fayda/maliyet katkılarını bulanık belirsizlik altında birlikte değerlendirir.",
        "academic": "Fuzzy COPRAS, bulanıklaştırılmış karar matrisinden türetilen fayda ve maliyet bileşenlerini göreli önem yaklaşımıyla birleştirerek sıralama üretir.",
    },
    "Fuzzy OCRA": {
        "simple": "Fuzzy OCRA, fayda ve maliyet rekabet puanlarını bulanık ortamda toplar.",
        "academic": "Fuzzy OCRA, bulanık performansların durulaştırılmış fayda-maliyet rekabet ölçülerini bütünleştirerek operasyonel üstünlük sıralaması sunar.",
    },
    "Fuzzy MOORA": {
        "simple": "Fuzzy MOORA, fayda ve maliyet etkilerini bulanık belirsizlik altında dengeler.",
        "academic": "Fuzzy MOORA, bulanıklaştırılmış kriter değerlerinden elde edilen durulaştırılmış oran yapısı ile net fayda skorları üretir.",
    },
    "Fuzzy MABAC": {
        "simple": "Fuzzy MABAC, alternatifin sınır yaklaşım alanına göre konumunu belirsizlikle birlikte ölçer.",
        "academic": "Fuzzy MABAC, bulanık performans değerlerini sınır yaklaşım alanı etrafında yorumlayarak alternatiflerin göreli üstünlüğünü belirler.",
    },
    "Fuzzy MARCOS": {
        "simple": "Fuzzy MARCOS, ideal ve anti-ideal referanslara göre bulanık yararlılık analizi yapar.",
        "academic": "Fuzzy MARCOS, bulanıklaştırılmış alternatifleri ideal/anti-ideal referanslar ve yararlılık fonksiyonu üzerinden karşılaştırır.",
    },
    "Fuzzy CoCoSo": {
        "simple": "Fuzzy CoCoSo, bulanık ortamda çoklu uzlaşı stratejilerini birleştirir.",
        "academic": "Fuzzy CoCoSo, bulanık veriden türetilen aritmetik ve karşılaştırmalı uzlaşı bileşenlerini hibritleyerek nihai skor üretir.",
    },
    "Fuzzy GRA": {
        "simple": "Fuzzy GRA, belirsizlik altında referans profile benzerliği ölçer.",
        "academic": "Fuzzy GRA, bulanıklaştırılmış karar matrisinde gri ilişkisel katsayıları kullanarak ideal profile yakınlığı değerlendirir.",
    },
    "Fuzzy PROMETHEE": {
        "simple": "Fuzzy PROMETHEE, bulanık ikili karşılaştırmalarla net tercih akışı üretir.",
        "academic": "Fuzzy PROMETHEE, bulanık performansların durulaştırılmış tercih fonksiyonları üzerinden pozitif/negatif akışlarını hesaplayarak tam sıralama sunar.",
    },
}

REFERENCE_LIBRARY: Dict[str, str] = {
    "GENERAL_MCDM": "Hwang, C.-L., & Yoon, K. (1981). Multiple attribute decision making: Methods and applications: A state-of-the-art survey. Springer-Verlag.",
    "Entropy": "Shannon, C. E. (1948). A mathematical theory of communication. Bell System Technical Journal, 27(3), 379–423; 27(4), 623–656.",
    "CRITIC": "Diakoulaki, D., Mavrotas, G., & Papayannakis, L. (1995). Determining objective weights in multiple criteria problems: The CRITIC method. Computers & Operations Research, 22(7), 763–770. https://doi.org/10.1016/0305-0548(94)00059-H",
    "Standart Sapma": "Mukhametzyanov, I. (2021). Specific character of objective methods for determining weights of criteria in MCDM problems: Entropy, CRITIC and SD. Decision Making: Applications in Management and Engineering, 4(2), 76–105. https://doi.org/10.31181/dmame210402076i",
    "MEREC": "Keshavarz-Ghorabaee, M., Amiri, M., Zavadskas, E. K., Turskis, Z., & Antucheviciene, J. (2021). Determination of objective weights using a new method based on the removal effects of criteria (MEREC). Symmetry, 13(4), 525. https://doi.org/10.3390/sym13040525",
    "LOPCOW": "Ecer, F., & Pamucar, D. (2022). A novel LOPCOW-DOBI multi-criteria sustainability performance assessment methodology: An application in developing country banking sector. Omega, 112, 102690. https://doi.org/10.1016/j.omega.2022.102690",
    "PCA": "Hotelling, H. (1933). Analysis of a complex of statistical variables into principal components. Journal of Educational Psychology, 24(6), 417–441.",
    "IDOCRIW": "Zavadskas, E. K., & Podvezko, V. (2016). Integrated determination of objective criteria weights in MCDM. International Journal of Information Technology & Decision Making, 15(2), 267–283. https://doi.org/10.1142/S0219622016500036",
    "Fuzzy IDOCRIW": "Hadi-Vencheh, A., Mirjaberi, S. M., & Hojabri, H. (2023). A fuzzy IDOCRIW weighting approach for multi-criteria decision-making under uncertainty. Expert Systems with Applications, 213, 119014. https://doi.org/10.1016/j.eswa.2022.119014",
    "TOPSIS": "Hwang, C.-L., & Yoon, K. (1981). Multiple attribute decision making: Methods and applications: A state-of-the-art survey. Springer-Verlag.",
    "VIKOR": "Opricovic, S., & Tzeng, G.-H. (2004). Compromise solution by MCDM methods: A comparative analysis of VIKOR and TOPSIS. European Journal of Operational Research, 156(2), 445–455. https://doi.org/10.1016/S0377-2217(03)00020-1",
    "EDAS": "Keshavarz-Ghorabaee, M., Zavadskas, E. K., Olfat, L., & Turskis, Z. (2015). Multi-criteria inventory classification using a new method of evaluation based on distance from average solution (EDAS). Informatica, 26(3), 435–451.",
    "CODAS": "Keshavarz-Ghorabaee, M., Zavadskas, E. K., Turskis, Z., & Antucheviciene, J. (2016). A new combinative distance-based assessment (CODAS) method for multi-criteria decision-making. Economic Computation and Economic Cybernetics Studies and Research, 50(3), 25–44.",
    "COPRAS": "Zavadskas, E. K., Kaklauskas, A., & Sarka, V. (1994). The new method of multicriteria complex proportional assessment of projects. Technological and Economic Development of Economy, 1(3), 131–139.",
    "OCRA": "Parkan, C., & Wu, M. L. (1997). Measurement of the performance of an investment bank using the operational competitiveness rating procedure (OCRA). European Journal of Operational Research, 98(3), 530–540.",
    "ARAS": "Zavadskas, E. K., & Turskis, Z. (2010). A new additive ratio assessment (ARAS) method in multicriteria decision-making. Technological and Economic Development of Economy, 16(2), 159–172. https://doi.org/10.3846/tede.2010.10",
    "SAW": "Fishburn, P. C. (1967). Additive utilities with incomplete product set: Applications to priorities and assignments. Operations Research, 15(3), 537–542.",
    "WPM": "Miller, D. W., & Starr, M. K. (1969). Executive decisions and operations research. Prentice-Hall.",
    "MAUT": "Keeney, R. L., & Raiffa, H. (1976). Decisions with multiple objectives: Preferences and value trade-offs. Wiley.",
    "WASPAS": "Zavadskas, E. K., Turskis, Z., Antucheviciene, J., & Zakarevicius, A. (2012). Optimization of weighted aggregated sum product assessment. Electronics and Electrical Engineering, 122(6), 3–6. https://doi.org/10.5755/j01.eee.122.6.1810",
    "MOORA": "Brauers, W. K. M., & Zavadskas, E. K. (2006). The MOORA method and its application to privatization in a transition economy. Control and Cybernetics, 35(2), 445–469.",
    "MABAC": "Pamučar, D., & Ćirović, G. (2015). The selection of transport and handling resources in logistics centers using Multi-Attributive Border Approximation area Comparison (MABAC). Expert Systems with Applications, 42(6), 3016–3028. https://doi.org/10.1016/j.eswa.2014.11.057",
    "MARCOS": "Stević, Ž., Pamučar, D., Puška, A., & Chatterjee, P. (2020). Sustainable supplier selection in healthcare industries using a new MCDM method: Measurement of alternatives and ranking according to COmpromise solution (MARCOS). Computers & Industrial Engineering, 140, 106231. https://doi.org/10.1016/j.cie.2019.106231",
    "CoCoSo": "Yazdani, M., Zarate, P., Zavadskas, E. K., & Turskis, Z. (2019). A combined compromise solution (CoCoSo) method for multi-criteria decision-making problems. Management Decision, 57(9), 2501–2519. https://doi.org/10.1108/MD-05-2017-0458",
    "PROMETHEE": "Brans, J. P., & Vincke, P. (1985). A preference ranking organisation method: The PROMETHEE method for multiple criteria decision-making. Management Science, 31(6), 647–656. https://doi.org/10.1287/mnsc.31.6.647",
    "GRA": "Deng, J. (1982). Control problems of grey systems. Systems & Control Letters, 1(5), 288–294.",
    "Fuzzy TOPSIS": "Chen, C.-T. (2000). Extensions of the TOPSIS for group decision-making under fuzzy environment. Fuzzy Sets and Systems, 114(1), 1–9.",
    "Fuzzy VIKOR": "Sanayei, A., Mousavi, S. F., & Yazdankhah, A. (2010). Group decision making process for supplier selection with VIKOR under fuzzy environment. Expert Systems with Applications, 37(1), 24–30. https://doi.org/10.1016/j.eswa.2009.04.063",
    "Fuzzy ARAS": "Turskis, Z., & Zavadskas, E. K. (2010). A new fuzzy additive ratio assessment method (ARAS-F). Case study: The analysis of fuzzy multiple criteria in order to select the logistic centers location. Transport, 25(4), 423–432.",
    "Fuzzy SAW": "Bellman, R. E., & Zadeh, L. A. (1970). Decision-making in a fuzzy environment. Management Science, 17(4), B-141–B-164.",
    "Fuzzy WPM": "Yager, R. R. (1981). A procedure for ordering fuzzy subsets of the unit interval. Information Sciences, 24(2), 143–161.",
    "Fuzzy MAUT": "Carlsson, C., & Fullér, R. (1996). Fuzzy multiple criteria decision making: Recent developments. Fuzzy Sets and Systems, 78(2), 139–153.",
    "Fuzzy WASPAS": "Zavadskas, E. K., Antucheviciene, J., Hajiagha, S. H. R., & Hashemi, S. S. (2014). Extension of weighted aggregated sum product assessment with interval-valued intuitionistic fuzzy numbers (WASPAS-IVIF). Applied Soft Computing, 24, 1013–1021. https://doi.org/10.1016/j.asoc.2014.08.031",
    "Fuzzy EDAS": "Kahraman, C., Keshavarz-Ghorabaee, M., Zavadskas, E. K., Cevik Onar, S., Yazdani, M., & Oztaysi, B. (2017). Intuitionistic fuzzy EDAS method: An application to solid waste disposal site selection. Journal of Environmental Engineering and Landscape Management, 25(1), 1–12.",
    "Fuzzy CODAS": "Badi, I., Abdulshahed, A., Shetwan, A., & Eltayeb, W. (2018). A hybrid MADM approach for supplier selection under uncertain environment based on fuzzy CODAS model. Journal of Industrial Engineering International, 14(2), 293–308.",
    "Fuzzy COPRAS": "Yazdani-Chamzini, A., Fouladgar, M. M., Zavadskas, E. K., & Moini, S. H. H. (2013). Selecting the optimal renewable energy using multi criteria decision making. Journal of Business Economics and Management, 14(5), 957–978.",
    "Fuzzy OCRA": "Madic, M., Gecevska, V., Radovanovic, M., & Petkovic, D. (2014). Multi-criteria economic analysis of machining processes using fuzzy OCRA method. Tehnicki vjesnik, 21(2), 299–304.",
    "Fuzzy MOORA": "Balezentis, T., & Zeng, S. (2013). Group decision making based upon interval-valued fuzzy numbers: An extension of the MOORA method. Expert Systems with Applications, 40(2), 543–550.",
    "Fuzzy MABAC": "Pamucar, D., Gigovic, L., Bajić, Z., & Janošević, M. (2017). Location selection for wind farms using GIS multi-criteria hybrid model: An approach based on fuzzy MABAC method. Renewable and Sustainable Energy Reviews, 76, 424–437.",
    "Fuzzy MARCOS": "Stević, Ž., Brković, N., Božanić, D., & Pamučar, D. (2020). A new fuzzy approach to compromise ranking: Fuzzy MARCOS method. Symmetry, 12(6), 894.",
    "Fuzzy CoCoSo": "Yazdani, M., Chatterjee, P., Pamucar, D., & Chakraborty, S. (2020). Development of an integrated decision making model for location selection under fuzzy environment based on CoCoSo method. Computers & Industrial Engineering, 142, 106356.",
    "Fuzzy GRA": "Kuo, Y., Yang, T., & Huang, G.-W. (2008). The use of grey relational analysis in solving multiple attribute decision-making problems. Computers & Industrial Engineering, 55(1), 80–93.",
    "Fuzzy PROMETHEE": "Goumas, M., & Lygerou, V. (2000). An extension of the PROMETHEE method for decision making in fuzzy environment. European Journal of Operational Research, 123(3), 606–613.",
    "CILOS": "Zavadskas, E. K., & Podvezko, V. (2016). Integrated determination of objective criteria weights in MCDM. International Journal of Information Technology & Decision Making, 15(2), 267–283. https://doi.org/10.1142/S0219622016500036",
    "SPOTIS": "Dezert, J., Tchamova, A., Han, D., & Tacnet, J.-M. (2020). The SPOTIS rank reversal free method for multi-criteria decision-making support. In 2020 IEEE 23rd International Conference on Information Fusion (FUSION) (pp. 1–8). https://doi.org/10.23919/FUSION45008.2020.9190347",
    "MULTIMOORA": "Brauers, W. K. M., & Zavadskas, E. K. (2010). Project management by MULTIMOORA as an instrument for transition economies. Technological and Economic Development of Economy, 16(1), 5–24. https://doi.org/10.3846/tede.2010.01",
    "RAWEC": "Sotoudeh-Anvari, A. (2023). A novel multi-attribute decision-making method based on the weight of each criterion (RAWEC). Journal of Soft Computing and Decision Analytics, 1(1), 192–207. https://doi.org/10.31181/jscda11202314",
    "RAFSI": "Žižović, M., Pamučar, D., Albijanić, M., Chatterjee, P., & Pribićević, I. (2020). Eliminating Rank Reversal Problem Using a New Multi-Attribute Model—The RAFSI Method. Mathematics, 8(6), 1015. https://doi.org/10.3390/math8061015",
    "ROV": "Yakowitz, D. S., Lane, L. J., & Szidarovszky, F. (1993). Multi-attribute decision making: Dominance with respect to an importance order of the attributes. Applied Mathematics and Computation, 54(2–3), 167–181. https://doi.org/10.1016/0096-3003(93)90057-L",
    "AROMAN": "Dimitrijević, B., Trpković, A., Atanasković, M., & Grbović, A. (2022). Application of AROMAN Method for Determination of Stability Level for Slope Failures. Civil Engineering Journal, 8(8), 1447–1462. https://doi.org/10.28991/CEJ-2022-08-08-012",
    "DNMA": "Liu, P., & Zhu, B. (2021). A novel psychophysical decision-making model with heterogeneous information and its application in the evaluation of unmanned aerial vehicle selection. Expert Systems with Applications, 166, 114091. https://doi.org/10.1016/j.eswa.2020.114091",
    "Fuzzy SPOTIS": "Shekhovtsov, A., & Sałabun, W. (2021). A comparative case study of the VIKOR and TOPSIS ranking methods in a novel approach for solving the fuzzy MCDM problem. Applied Soft Computing, 111, 107637. https://doi.org/10.1016/j.asoc.2021.107637",
    "Fuzzy MULTIMOORA": "Brauers, W. K. M., & Zavadskas, E. K. (2012). Robustness of MULTIMOORA: A method for multi-objective optimization. Informatica, 23(1), 1–25.",
    "Fuzzy RAWEC": "Sotoudeh-Anvari, A. (2023). A novel multi-attribute decision-making method based on the weight of each criterion (RAWEC). Journal of Soft Computing and Decision Analytics, 1(1), 192–207.",
    "Fuzzy RAFSI": "Žižović, M., Pamučar, D., Albijanić, M., Chatterjee, P., & Pribićević, I. (2020). Eliminating Rank Reversal Problem Using a New Multi-Attribute Model—The RAFSI Method. Mathematics, 8(6), 1015. https://doi.org/10.3390/math8061015",
    "Fuzzy ROV": "Yakowitz, D. S., Lane, L. J., & Szidarovszky, F. (1993). Multi-attribute decision making: Dominance with respect to an importance order of the attributes. Applied Mathematics and Computation, 54(2–3), 167–181.",
    "Fuzzy AROMAN": "Dimitrijević, B., Trpković, A., Atanasković, M., & Grbović, A. (2022). Application of AROMAN Method for Determination of Stability Level for Slope Failures. Civil Engineering Journal, 8(8), 1447–1462.",
    "Fuzzy DNMA": "Liu, P., & Zhu, B. (2021). A novel psychophysical decision-making model with heterogeneous information. Expert Systems with Applications, 166, 114091.",
}

MANDATORY_MCDM_REFERENCE = (
    'Dalbudak, E., Rençber, Ö. F. (2022). "Literature Review on Multi-Criteria Decision Making Methods". '
    "Gaziantep Üniversitesi, İktisadi ve İdari Bilimler Fakülte Dergisi, 4(1), 1-17."
)

RENCBER_PUBLICATIONS_BY_METHOD: Dict[str, List[str]] = {
    "TOPSIS": [
        "Yildizhan, C., Karakuş, M., Rençber, Ö. F., & Can, M. (2024). Measuring financial performance with Best Worst-based TOPSIS approach. Kaynak: https://ofrencber.com/publications/",
        "Coşkuner, M. F., & Rençber, Ö. F. (2024). CRITIC-based TOPSIS method application in enterprise risk performance analysis. Kaynak: https://ofrencber.com/publications/",
    ],
    "CRITIC": [
        "Coşkuner, M. F., & Rençber, Ö. F. (2024). CRITIC-based TOPSIS method application in enterprise risk performance analysis. Kaynak: https://ofrencber.com/publications/",
    ],
    "IDOCRIW": [
        "Çiftaşlan, A., & Rençber, Ö. F. (2022). IDOCRIW and CoCoSo based performance analysis of deposit banks. Kaynak: https://ofrencber.com/publications/",
    ],
    "PROMETHEE": [
        "Rençber, Ö. F. (2018). Ranking of provinces in terms of quality of life by PROMETHEE method in Turkey. Kaynak: https://ofrencber.com/publications/",
        "Dalbudak, E., & Rençber, Ö. F. (2023). Comparison of BWM-based TOPSIS, PROMETHEE and COPRAS methods in environmental sustainability. Kaynak: https://ofrencber.com/publications/",
    ],
    "COPRAS": [
        "Dalbudak, E., & Rençber, Ö. F. (2023). Comparison of BWM-based TOPSIS, PROMETHEE and COPRAS methods in environmental sustainability. Kaynak: https://ofrencber.com/publications/",
    ],
    "VIKOR": [
        "Rençber, Ö. F. (2019). Comparison of Grey Relational Analysis and VIKOR methods in ranking countries by life quality. Kaynak: https://ofrencber.com/publications/",
    ],
    "GRA": [
        "Rençber, Ö. F. (2019). Comparison of Grey Relational Analysis and VIKOR methods in ranking countries by life quality. Kaynak: https://ofrencber.com/publications/",
    ],
    "WASPAS": [
        "Rençber, Ö. F., & Avci, T. (2018). Comparison of public and private banks in Turkey by WASPAS method. Kaynak: https://ofrencber.com/publications/",
    ],
    "CoCoSo": [
        "Çiftaşlan, A., & Rençber, Ö. F. (2022). IDOCRIW and CoCoSo based performance analysis of deposit banks. Kaynak: https://ofrencber.com/publications/",
    ],
}


@dataclass
class AnalysisConfig:
    criteria: List[str]
    criteria_types: Dict[str, str]
    weight_method: str
    weight_mode: str = "objective"  # objective | equal | manual
    manual_weights: Optional[Dict[str, float]] = None
    ranking_method: Optional[str] = None
    compare_methods: Optional[List[str]] = None
    vikor_v: float = 0.5
    waspas_lambda: float = 0.5
    codas_tau: float = 0.02
    cocoso_lambda: float = 0.5
    gra_rho: float = 0.5
    promethee_pref_func: str = "linear"  # usual | u_shape | v_shape | level | linear | gaussian
    promethee_q: float = 0.05
    promethee_p: float = 0.30
    promethee_s: float = 0.20
    fuzzy_spread: float = 0.10
    sensitivity_iterations: int = 200
    sensitivity_sigma: float = 0.12
    run_heavy_robustness: bool = True


def _as_numeric_df(data: pd.DataFrame, criteria: Sequence[str]) -> pd.DataFrame:
    df = data.loc[:, list(criteria)].copy()
    for c in criteria:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    return df.astype(float)


def _normalize_weights(values: Sequence[float]) -> np.ndarray:
    arr = np.asarray(values, dtype=float)
    arr = np.where(np.isfinite(arr), arr, 0.0)
    arr = np.clip(arr, 0.0, None)
    total = arr.sum()
    if total <= EPS:
        arr = np.ones_like(arr, dtype=float)
        total = arr.sum()
    return arr / total


def _safe_divide(num: np.ndarray, den: np.ndarray | float) -> np.ndarray:
    return np.asarray(num, dtype=float) / (np.asarray(den, dtype=float) + EPS)


def _shift_positive(df: pd.DataFrame) -> Tuple[pd.DataFrame, Dict[str, float]]:
    shifted = df.copy()
    shifts: Dict[str, float] = {}
    for c in shifted.columns:
        minv = float(shifted[c].min())
        if minv <= 0:
            magnitude = max(abs(minv), 1.0)
            delta = (abs(minv) * (1.0 + EPS)) if abs(minv) > EPS else (magnitude * EPS)
            shifted[c] = shifted[c] + delta
            shifts[c] = delta
    return shifted, shifts


def _normalize_minmax(data: pd.DataFrame, criteria_types: Dict[str, str]) -> pd.DataFrame:
    df = data.copy().astype(float)
    out = pd.DataFrame(index=df.index, columns=df.columns, dtype=float)
    for c in df.columns:
        col = df[c].to_numpy(dtype=float)
        cmin = np.nanmin(col)
        cmax = np.nanmax(col)
        denom = cmax - cmin
        if abs(denom) <= EPS:
            out[c] = np.zeros(len(df))
            continue
        if criteria_types.get(c, "max") == "max":
            out[c] = (col - cmin) / denom
        else:
            out[c] = (cmax - col) / denom
    return out.fillna(0.0)


def _normalize_sum(data: pd.DataFrame, criteria_types: Dict[str, str]) -> pd.DataFrame:
    df = data.copy().astype(float)
    out = pd.DataFrame(index=df.index, columns=df.columns, dtype=float)
    for c in df.columns:
        col = df[c].to_numpy(dtype=float)
        cmin = float(np.nanmin(col))
        cmax = float(np.nanmax(col))
        if criteria_types.get(c, "max") == "max":
            transformed = col - cmin
        else:
            transformed = cmax - col
        transformed = np.clip(transformed, 0.0, None) + EPS
        out[c] = _safe_divide(transformed, transformed.sum())
    return out.fillna(0.0)


def _normalize_vector(data: pd.DataFrame) -> pd.DataFrame:
    df = data.copy().astype(float)
    norms = np.sqrt((df ** 2).sum(axis=0)) + EPS
    return df / norms


def _vector_rank(scores: Sequence[float], index: Sequence[Any], *, ascending: bool = False) -> pd.DataFrame:
    out = pd.DataFrame({"Alternatif": list(index), "Skor": np.asarray(scores, dtype=float)})
    out["Sıra"] = out["Skor"].rank(ascending=ascending, method="min").astype(int)
    return out.sort_values(["Sıra", "Alternatif"]).reset_index(drop=True)


def _is_lower_better_method(method: Optional[str]) -> bool:
    if not method:
        return False
    base = method.replace("Fuzzy ", "")
    return base in {"VIKOR", "SPOTIS", "MULTIMOORA"}


def _benefit_cost_summary(criteria_types: Dict[str, str]) -> Dict[str, List[str]]:
    benefit = [k for k, v in criteria_types.items() if v == "max"]
    cost = [k for k, v in criteria_types.items() if v == "min"]
    return {"benefit": benefit, "cost": cost}


def apply_threshold_filter(
    data: pd.DataFrame,
    criteria: Sequence[str],
    criteria_types: Dict[str, str],
    thresholds: Dict[str, Dict[str, float]],
) -> Tuple[pd.DataFrame, List[Dict[str, Any]]]:
    """Pre-screen alternatives by minimum/maximum thresholds.

    Parameters
    ----------
    thresholds : dict
        ``{criterion: {"min": float, "max": float}}``.
        For benefit criteria: alternatives below ``min`` are eliminated.
        For cost criteria: alternatives above ``max`` are eliminated.
        Keys are optional — omit ``min``/``max`` to skip that bound.

    Returns
    -------
    filtered_data : pd.DataFrame
        Data with eliminated rows removed.
    eliminated : list of dict
        Each entry: ``{"alternative": str, "criterion": str, "value": float,
        "threshold": float, "direction": str}``.
    """
    if not thresholds:
        return data, []

    eliminated: List[Dict[str, Any]] = []
    keep_mask = pd.Series(True, index=data.index)
    df = _as_numeric_df(data, criteria)

    for crit, bounds in thresholds.items():
        if crit not in df.columns:
            continue
        col = df[crit]
        ctype = criteria_types.get(crit, "max")
        min_val = bounds.get("min")
        max_val = bounds.get("max")

        if min_val is not None:
            fail = col < float(min_val)
            for idx in df.index[fail & keep_mask]:
                eliminated.append({
                    "alternative": str(idx),
                    "criterion": crit,
                    "value": float(col.loc[idx]),
                    "threshold": float(min_val),
                    "direction": f"< min ({min_val})",
                })
            keep_mask &= ~fail

        if max_val is not None:
            fail = col > float(max_val)
            for idx in df.index[fail & keep_mask]:
                eliminated.append({
                    "alternative": str(idx),
                    "criterion": crit,
                    "value": float(col.loc[idx]),
                    "threshold": float(max_val),
                    "direction": f"> max ({max_val})",
                })
            keep_mask &= ~fail

    return data.loc[keep_mask].copy(), eliminated


def validate_problem(data: pd.DataFrame, criteria: Sequence[str], criteria_types: Dict[str, str]) -> Dict[str, Any]:
    df = _as_numeric_df(data, criteria)
    issues: List[str] = []
    warnings: List[str] = []
    if df.isna().any().any():
        warnings.append("Karar matrisi eksik gözlem içeriyor; analizden önce doldurma/temizleme uygulanmalıdır.")
    duplicated_names = pd.Index(data.index.astype(str)).duplicated()
    if duplicated_names.any():
        warnings.append("Alternatif isimlerinde tekrar var; raporda ilk görülen ad korunacaktır.")
    constant_criteria = [c for c in df.columns if float(df[c].nunique(dropna=True)) <= 1]
    if constant_criteria:
        warnings.append(
            "Ayırt ediciliği olmayan sabit kriter(ler) tespit edildi: "
            + ", ".join(constant_criteria)
            + ". Objektif ağırlıkları doğal olarak sıfıra veya sıfıra yakına düşebilir."
        )
    non_positive = [c for c in df.columns if float(df[c].min()) <= 0]
    corr = df.corr(numeric_only=True)
    high_corr_pairs: List[Tuple[str, str, float]] = []
    for i, c1 in enumerate(df.columns):
        for c2 in df.columns[i + 1 :]:
            val = corr.loc[c1, c2]
            if pd.notna(val) and abs(val) >= 0.85:
                high_corr_pairs.append((c1, c2, float(val)))
    if high_corr_pairs:
        warnings.append(
            "Yüksek korelasyonlu kriter çiftleri bulundu; CRITIC ve PCA türü yaklaşımlar daha ayırt edici olabilir."
        )
    if len(criteria) < 2:
        issues.append("Analiz için en az iki kriter gereklidir.")
    if len(df) < 2:
        issues.append("Analiz için en az iki alternatif gereklidir.")
    return {
        "errors": issues,
        "warnings": warnings,
        "constant_criteria": constant_criteria,
        "non_positive_criteria": non_positive,
        "high_corr_pairs": high_corr_pairs,
        "shape": df.shape,
        "direction_summary": _benefit_cost_summary(criteria_types),
    }


def descriptive_statistics(data: pd.DataFrame, criteria: Sequence[str]) -> pd.DataFrame:
    df = _as_numeric_df(data, criteria)
    desc = df.describe().T
    desc["cv"] = _safe_divide(desc["std"], desc["mean"].abs() + EPS)
    desc["range"] = desc["max"] - desc["min"]
    desc = desc.reset_index().rename(columns={"index": "Kriter"})
    return desc


def pca_diagnostics(data: pd.DataFrame, criteria: Sequence[str], criteria_types: Dict[str, str]) -> Dict[str, Any]:
    from sklearn.decomposition import PCA
    from sklearn.preprocessing import StandardScaler
    df = _as_numeric_df(data, criteria)
    oriented = _normalize_minmax(df, criteria_types)
    scaler = StandardScaler()
    X = scaler.fit_transform(oriented)
    pca = PCA()
    scores = pca.fit_transform(X)
    eigvals = pd.Series(pca.explained_variance_, index=[f"PC{i+1}" for i in range(pca.n_components_)], name="eigenvalue")
    var_ratio = pd.Series(
        pca.explained_variance_ratio_,
        index=[f"PC{i+1}" for i in range(pca.n_components_)],
        name="variance_ratio",
    )
    loadings = pd.DataFrame(
        pca.components_.T * np.sqrt(pca.explained_variance_),
        index=list(criteria),
        columns=[f"PC{i+1}" for i in range(pca.n_components_)],
    )
    score_df = pd.DataFrame(scores[:, :2], columns=["PC1", "PC2"], index=data.index)
    score_df.insert(0, "Alternatif", score_df.index.astype(str))
    selected_components = [pc for pc, ev in eigvals.items() if ev > 1.0]
    if not selected_components:
        selected_components = [eigvals.index[0]]
    return {
        "oriented_matrix": oriented,
        "explained_variance": eigvals.reset_index().rename(columns={"index": "Bileşen"}),
        "explained_ratio": var_ratio.reset_index().rename(columns={"index": "Bileşen"}),
        "loadings": loadings.reset_index().rename(columns={"index": "Kriter"}),
        "score_df": score_df.reset_index(drop=True),
        "selected_components": selected_components,
    }


def _weights_entropy(data: pd.DataFrame, criteria_types: Dict[str, str]) -> Tuple[Dict[str, float], Dict[str, Any]]:
    ndata = _normalize_sum(data, criteria_types)
    m = len(ndata)
    k = 1.0 / np.log(m) if m > 1 else 0.0
    p = ndata.to_numpy(dtype=float) + EPS
    entropy = -k * np.sum(p * np.log(p), axis=0)
    divergence = 1.0 - entropy
    weights = _normalize_weights(divergence)
    w = dict(zip(ndata.columns, weights))
    det = {
        "normalized_matrix": ndata,
        "entropy": pd.DataFrame({"Kriter": ndata.columns, "Entropy": entropy, "Divergence": divergence, "Ağırlık": weights}),
    }
    return w, det


def _weights_critic(data: pd.DataFrame, criteria_types: Dict[str, str]) -> Tuple[Dict[str, float], Dict[str, Any]]:
    ndata = _normalize_minmax(data, criteria_types)
    stds = ndata.std(axis=0, ddof=1).to_numpy(dtype=float)
    corr = ndata.corr().fillna(0.0).to_numpy(dtype=float)
    contrast = stds * np.sum(1.0 - corr, axis=1)
    weights = _normalize_weights(contrast)
    w = dict(zip(ndata.columns, weights))
    det = {
        "normalized_matrix": ndata,
        "correlation_matrix": pd.DataFrame(corr, index=ndata.columns, columns=ndata.columns),
        "critic": pd.DataFrame({"Kriter": ndata.columns, "StdSapma": stds, "Bilgiİçeriği": contrast, "Ağırlık": weights}),
    }
    return w, det


def _weights_sd(data: pd.DataFrame, criteria_types: Dict[str, str]) -> Tuple[Dict[str, float], Dict[str, Any]]:
    ndata = _normalize_minmax(data, criteria_types)
    stds = ndata.std(axis=0, ddof=1).to_numpy(dtype=float)
    weights = _normalize_weights(stds)
    w = dict(zip(ndata.columns, weights))
    det = {
        "normalized_matrix": ndata,
        "std_summary": pd.DataFrame({"Kriter": ndata.columns, "StdSapma": stds, "Ağırlık": weights}),
    }
    return w, det


def _weights_merec(data: pd.DataFrame, criteria_types: Dict[str, str]) -> Tuple[Dict[str, float], Dict[str, Any]]:
    pos, shifts = _shift_positive(data)
    m, n = pos.shape
    norm = pd.DataFrame(index=pos.index, columns=pos.columns, dtype=float)
    for c in pos.columns:
        col = pos[c].to_numpy(dtype=float)
        if criteria_types.get(c, "max") == "max":
            norm[c] = col.min() / (col + EPS)
        else:
            norm[c] = col / (col.max() + EPS)
    s = np.log1p(np.sum(np.abs(np.log(norm.to_numpy(dtype=float) + EPS)), axis=1) / max(n, 1))
    effects = []
    for j, c in enumerate(pos.columns):
        reduced = np.delete(norm.to_numpy(dtype=float), j, axis=1)
        denom = max(reduced.shape[1], 1)
        s_ex = np.log1p(np.sum(np.abs(np.log(reduced + EPS)), axis=1) / denom)
        effects.append(np.sum(np.abs(s_ex - s)))
    effects = np.asarray(effects, dtype=float)
    weights = _normalize_weights(effects)
    w = dict(zip(pos.columns, weights))
    det = {
        "positive_shifts": shifts,
        "normalized_matrix": norm,
        "overall_performance": pd.DataFrame({"Alternatif": pos.index.astype(str), "S_i": s}),
        "removal_effects": pd.DataFrame({"Kriter": pos.columns, "Etkisi": effects, "Ağırlık": weights}),
    }
    return w, det


def _weights_lopcow(data: pd.DataFrame, criteria_types: Dict[str, str]) -> Tuple[Dict[str, float], Dict[str, Any]]:
    norm = _normalize_minmax(data, criteria_types)
    stds = norm.std(axis=0, ddof=1).to_numpy(dtype=float) + EPS
    rms = np.sqrt((norm.to_numpy(dtype=float) ** 2).mean(axis=0)) + EPS
    pv = np.abs(np.log(_safe_divide(rms, stds)) * 100.0)
    weights = _normalize_weights(pv)
    w = dict(zip(norm.columns, weights))
    det = {
        "normalized_matrix": norm,
        "lopcow": pd.DataFrame({"Kriter": norm.columns, "RMS": rms, "StdSapma": stds, "PV": pv, "Ağırlık": weights}),
    }
    return w, det


def _weights_pca(data: pd.DataFrame, criteria_types: Dict[str, str]) -> Tuple[Dict[str, float], Dict[str, Any]]:
    diag = pca_diagnostics(data, data.columns, criteria_types)
    loadings = diag["loadings"].set_index("Kriter")
    ratios = diag["explained_ratio"].set_index("Bileşen")["variance_ratio"]
    selected = diag["selected_components"]
    selected_loadings = loadings.loc[:, selected].abs()
    selected_ratios = ratios.loc[selected]
    raw = selected_loadings.mul(selected_ratios, axis=1).sum(axis=1)
    weights = _normalize_weights(raw.to_numpy(dtype=float))
    w = dict(zip(loadings.index, weights))
    det = {
        "loadings": loadings.reset_index(),
        "explained_ratio": diag["explained_ratio"],
        "selected_components": selected,
        "pca_weight_table": pd.DataFrame({"Kriter": loadings.index, "HamAğırlık": raw.values, "Ağırlık": weights}),
        "score_df": diag["score_df"],
        "oriented_matrix": diag["oriented_matrix"],
    }
    return w, det

def _cilos_benefit_matrix(data: pd.DataFrame, criteria_types: Dict[str, str]) -> pd.DataFrame:
    pos, _ = _shift_positive(data)
    out = pd.DataFrame(index=pos.index, columns=pos.columns, dtype=float)
    for c in pos.columns:
        col = pos[c].to_numpy(dtype=float)
        if criteria_types.get(c, "max") == "max":
            out[c] = col / (col.max() + EPS)
        else:
            out[c] = col.min() / (col + EPS)
    return out

def _weights_cilos(data: pd.DataFrame, criteria_types: Dict[str, str]) -> Tuple[Dict[str, float], Dict[str, Any]]:
    benefit = _cilos_benefit_matrix(data, criteria_types)
    cols = list(benefit.columns)
    best_rows = {c: benefit[c].idxmax() for c in cols}
    a_matrix = pd.DataFrame(
        [benefit.loc[best_rows[c], cols].to_numpy(dtype=float) for c in cols],
        index=cols,
        columns=cols,
        dtype=float,
    )

    a_vals = a_matrix.to_numpy(dtype=float)
    diag_vals = np.diag(a_vals)
    p_vals = np.maximum(0.0, (diag_vals[np.newaxis, :] - a_vals) / (diag_vals[np.newaxis, :] + EPS))
    p_matrix = pd.DataFrame(p_vals, index=cols, columns=cols, dtype=float)

    f_matrix = p_matrix.T.copy()
    for crit in cols:
        f_matrix.loc[crit, crit] = -float(p_matrix.loc[:, crit].sum())

    lhs = f_matrix.to_numpy(dtype=float)
    lhs[-1, :] = 1.0
    rhs = np.zeros(len(cols), dtype=float)
    rhs[-1] = 1.0
    raw = np.linalg.lstsq(lhs, rhs, rcond=None)[0]
    raw = np.clip(raw, 0.0, None)
    weights = _normalize_weights(raw)
    w = dict(zip(cols, weights))
    det = {
        "benefit_matrix": benefit,
        "best_rows": pd.DataFrame({"Kriter": cols, "EnİyiAlternatif": [best_rows[c] for c in cols]}),
        "impact_loss_matrix": p_matrix,
        "cilos": pd.DataFrame({"Kriter": cols, "HamAğırlık": raw, "Ağırlık": weights}),
    }
    return w, det

def _weights_idocriw(data: pd.DataFrame, criteria_types: Dict[str, str]) -> Tuple[Dict[str, float], Dict[str, Any]]:
    entropy_w, entropy_det = _weights_entropy(data, criteria_types)
    cilos_w, cilos_det = _weights_cilos(data, criteria_types)
    cols = list(data.columns)
    entropy_vec = np.asarray([float(entropy_w[c]) for c in cols], dtype=float)
    cilos_vec = np.asarray([float(cilos_w[c]) for c in cols], dtype=float)
    raw = entropy_vec * cilos_vec
    weights = _normalize_weights(raw)
    w = dict(zip(cols, weights))
    det = {
        "entropy": entropy_det.get("entropy"),
        "benefit_matrix": cilos_det.get("benefit_matrix"),
        "impact_loss_matrix": cilos_det.get("impact_loss_matrix"),
        "cilos": cilos_det.get("cilos"),
        "idocriw": pd.DataFrame(
            {
                "Kriter": cols,
                "EntropyAğırlığı": entropy_vec,
                "CILOSAğırlığı": cilos_vec,
                "HamAğırlık": raw,
                "Ağırlık": weights,
            }
        ),
    }
    return w, det

def _weights_fuzzy_idocriw(
    data: pd.DataFrame,
    criteria_types: Dict[str, str],
    spread: float = 0.10,
) -> Tuple[Dict[str, float], Dict[str, Any]]:
    pos, _ = _shift_positive(data)
    lower = pos * max(0.0, 1.0 - spread)
    middle = pos.copy()
    upper = pos * (1.0 + spread)
    scenario_defs = [("Lower", lower), ("Middle", middle), ("Upper", upper)]
    entropy_rows: List[Dict[str, Any]] = []
    cilos_rows: List[Dict[str, Any]] = []
    idocriw_rows: List[Dict[str, Any]] = []
    entropy_stack: List[np.ndarray] = []
    cilos_stack: List[np.ndarray] = []
    idocriw_stack: List[np.ndarray] = []
    cols = list(pos.columns)

    for scenario_name, scenario_df in scenario_defs:
        entropy_w, _ = _weights_entropy(scenario_df, criteria_types)
        cilos_w, _ = _weights_cilos(scenario_df, criteria_types)
        entropy_vec = np.asarray([float(entropy_w[c]) for c in cols], dtype=float)
        cilos_vec = np.asarray([float(cilos_w[c]) for c in cols], dtype=float)
        idocriw_vec = _normalize_weights(entropy_vec * cilos_vec)
        entropy_stack.append(entropy_vec)
        cilos_stack.append(cilos_vec)
        idocriw_stack.append(idocriw_vec)
        for crit, e_val, c_val, i_val in zip(cols, entropy_vec, cilos_vec, idocriw_vec):
            entropy_rows.append({"Senaryo": scenario_name, "Kriter": crit, "EntropyAğırlığı": e_val})
            cilos_rows.append({"Senaryo": scenario_name, "Kriter": crit, "CILOSAğırlığı": c_val})
            idocriw_rows.append({"Senaryo": scenario_name, "Kriter": crit, "IDOCRIWAğırlığı": i_val})

    entropy_avg = np.mean(np.vstack(entropy_stack), axis=0)
    cilos_avg = np.mean(np.vstack(cilos_stack), axis=0)
    weights = _normalize_weights(np.mean(np.vstack(idocriw_stack), axis=0))
    w = dict(zip(cols, weights))
    det = {
        "fuzzy_entropy": pd.DataFrame(entropy_rows),
        "fuzzy_cilos": pd.DataFrame(cilos_rows),
        "fuzzy_idocriw_scenarios": pd.DataFrame(idocriw_rows),
        "fuzzy_idocriw": pd.DataFrame(
            {
                "Kriter": cols,
                "OrtalamaEntropy": entropy_avg,
                "OrtalamaCILOS": cilos_avg,
                "Ağırlık": weights,
            }
        ),
        "parameters": {"spread": float(spread)},
    }
    return w, det


def compute_objective_weights(
    data: pd.DataFrame,
    criteria: Sequence[str],
    criteria_types: Dict[str, str],
    method: str,
    fuzzy_spread: float = 0.10,
) -> Tuple[Dict[str, float], Dict[str, Any]]:
    df = _as_numeric_df(data, criteria)
    dispatch = {
        "Entropy": _weights_entropy,
        "CRITIC": _weights_critic,
        "Standart Sapma": _weights_sd,
        "MEREC": _weights_merec,
        "LOPCOW": _weights_lopcow,
        "PCA": _weights_pca,
        "CILOS": _weights_cilos,
        "IDOCRIW": _weights_idocriw,
        "Fuzzy IDOCRIW": lambda d, ct: _weights_fuzzy_idocriw(d, ct, spread=fuzzy_spread),
    }
    if method not in dispatch:
        raise ValueError(f"Desteklenmeyen ağırlıklandırma yöntemi: {method}")
    weights, details = dispatch[method](df, criteria_types)
    table = pd.DataFrame({"Kriter": list(weights.keys()), "Ağırlık": list(weights.values())}).sort_values("Ağırlık", ascending=False)
    table = table.reset_index(drop=True)
    table.insert(1, "ÖnemSırası", np.arange(1, len(table) + 1, dtype=int))
    details["weight_table"] = table
    return weights, details


def _contribution_table(data: pd.DataFrame, criteria_types: Dict[str, str], weights: Dict[str, float]) -> pd.DataFrame:
    norm = _normalize_minmax(data, criteria_types)
    weight_vec = np.asarray([weights[c] for c in norm.columns], dtype=float)
    contrib = norm.mul(weight_vec, axis=1)
    contrib.insert(0, "Alternatif", contrib.index.astype(str))
    return contrib.reset_index(drop=True)


def _rank_topsis(data: pd.DataFrame, criteria_types: Dict[str, str], weights: Dict[str, float]) -> Tuple[np.ndarray, Dict[str, Any]]:
    norm = _normalize_vector(data)
    wvec = np.asarray([weights[c] for c in data.columns], dtype=float)
    v = norm.to_numpy(dtype=float) * wvec
    pis = np.zeros(data.shape[1], dtype=float)
    nis = np.zeros(data.shape[1], dtype=float)
    for j, c in enumerate(data.columns):
        col = v[:, j]
        if criteria_types.get(c, "max") == "max":
            pis[j], nis[j] = col.max(), col.min()
        else:
            pis[j], nis[j] = col.min(), col.max()
    dplus = np.sqrt(((v - pis) ** 2).sum(axis=1))
    dminus = np.sqrt(((v - nis) ** 2).sum(axis=1))
    score = dminus / (dplus + dminus + EPS)
    det = {
        "normalized_matrix": norm,
        "weighted_matrix": pd.DataFrame(v, index=data.index, columns=data.columns),
        "ideals": pd.DataFrame({"Kriter": data.columns, "PIS": pis, "NIS": nis}),
        "distance_table": pd.DataFrame({"Alternatif": data.index.astype(str), "D+": dplus, "D-": dminus, "Skor": score}),
    }
    return score, det


def _rank_vikor(
    data: pd.DataFrame, criteria_types: Dict[str, str], weights: Dict[str, float], v_param: float = 0.5
) -> Tuple[np.ndarray, Dict[str, Any]]:
    x = data.to_numpy(dtype=float)
    wvec = np.asarray([weights[c] for c in data.columns], dtype=float)
    best = np.zeros(data.shape[1], dtype=float)
    worst = np.zeros(data.shape[1], dtype=float)
    for j, c in enumerate(data.columns):
        col = x[:, j]
        if criteria_types.get(c, "max") == "max":
            best[j], worst[j] = col.max(), col.min()
        else:
            best[j], worst[j] = col.min(), col.max()
    diff = np.zeros_like(x)
    for j in range(x.shape[1]):
        denom = abs(best[j] - worst[j])
        if denom <= EPS:
            diff[:, j] = 0.0
        else:
            diff[:, j] = wvec[j] * (best[j] - x[:, j]) / (denom + EPS)
    s = diff.sum(axis=1)
    r = diff.max(axis=1)
    s_best, s_worst = s.min(), s.max()
    r_best, r_worst = r.min(), r.max()
    q = v_param * _safe_divide(s - s_best, s_worst - s_best) + (1 - v_param) * _safe_divide(r - r_best, r_worst - r_best)
    score = q.copy()
    order_q = np.argsort(q)
    q_sorted = q[order_q]
    if len(q_sorted) > 1:
        advantage = (q_sorted[1] - q_sorted[0]) >= 1.0 / max(len(q_sorted) - 1, 1)
    else:
        advantage = True
    best_idx = int(order_q[0]) if len(order_q) else 0
    stability = best_idx in {int(np.argmin(s)), int(np.argmin(r))}
    det = {
        "criterion_best_worst": pd.DataFrame({"Kriter": data.columns, "Enİyi": best, "EnKötü": worst, "Ağırlık": wvec}),
        "vikor_table": pd.DataFrame({"Alternatif": data.index.astype(str), "S": s, "R": r, "Q": q, "Skor": score}).sort_values("Q"),
        "compromise_conditions": {
            "acceptable_advantage": bool(advantage),
            "acceptable_stability": bool(stability),
            "v_param": float(v_param),
        },
    }
    return score, det


def _rank_edas(data: pd.DataFrame, criteria_types: Dict[str, str], weights: Dict[str, float]) -> Tuple[np.ndarray, Dict[str, Any]]:
    x = data.to_numpy(dtype=float)
    av = x.mean(axis=0)
    pda = np.zeros_like(x)
    nda = np.zeros_like(x)
    for j, c in enumerate(data.columns):
        if criteria_types.get(c, "max") == "max":
            pda[:, j] = np.maximum(0.0, (x[:, j] - av[j]) / (abs(av[j]) + EPS))
            nda[:, j] = np.maximum(0.0, (av[j] - x[:, j]) / (abs(av[j]) + EPS))
        else:
            pda[:, j] = np.maximum(0.0, (av[j] - x[:, j]) / (abs(av[j]) + EPS))
            nda[:, j] = np.maximum(0.0, (x[:, j] - av[j]) / (abs(av[j]) + EPS))
    wvec = np.asarray([weights[c] for c in data.columns], dtype=float)
    sp = np.sum(pda * wvec, axis=1)
    sn = np.sum(nda * wvec, axis=1)
    nsp = _safe_divide(sp, sp.max())
    nsn = 1.0 - _safe_divide(sn, sn.max())
    score = 0.5 * (nsp + nsn)
    det = {
        "average_solution": pd.DataFrame({"Kriter": data.columns, "Ortalama": av, "Ağırlık": wvec}),
        "edas_table": pd.DataFrame({"Alternatif": data.index.astype(str), "SP": sp, "SN": sn, "NSP": nsp, "NSN": nsn, "Skor": score}),
        "pda_matrix": pd.DataFrame(pda, index=data.index, columns=data.columns),
        "nda_matrix": pd.DataFrame(nda, index=data.index, columns=data.columns),
    }
    return score, det


def _rank_codas(
    data: pd.DataFrame, criteria_types: Dict[str, str], weights: Dict[str, float], tau: float = 0.02
) -> Tuple[np.ndarray, Dict[str, Any]]:
    pos, _ = _shift_positive(data)
    norm = pd.DataFrame(index=pos.index, columns=pos.columns, dtype=float)
    for c in pos.columns:
        col = pos[c].to_numpy(dtype=float)
        if criteria_types.get(c, "max") == "max":
            norm[c] = _safe_divide(col, col.max())
        else:
            norm[c] = _safe_divide(col.min(), col)
    wvec = np.asarray([weights[c] for c in pos.columns], dtype=float)
    v = norm.to_numpy(dtype=float) * wvec
    nis = v.min(axis=0)
    e = np.sqrt(((v - nis) ** 2).sum(axis=1))
    t = np.abs(v - nis).sum(axis=1)
    delta_e = e[:, np.newaxis] - e[np.newaxis, :]
    delta_t = t[:, np.newaxis] - t[np.newaxis, :]
    psi = (np.abs(delta_e) >= tau).astype(float)
    h = delta_e.sum(axis=1) + np.sum(psi * delta_t, axis=1)
    score = _safe_divide(h - h.min(), h.max() - h.min())
    det = {
        "normalized_matrix": norm,
        "weighted_matrix": pd.DataFrame(v, index=pos.index, columns=pos.columns),
        "codas_table": pd.DataFrame({"Alternatif": pos.index.astype(str), "E": e, "T": t, "H": h, "Skor": score}),
        "parameters": {"tau": float(tau)},
    }
    return score, det


def _rank_copras(data: pd.DataFrame, criteria_types: Dict[str, str], weights: Dict[str, float]) -> Tuple[np.ndarray, Dict[str, Any]]:
    pos, shifts = _shift_positive(data)
    arr = pos.to_numpy(dtype=float)
    col_sums = arr.sum(axis=0) + EPS
    norm = arr / col_sums
    wvec = np.asarray([weights[c] for c in pos.columns], dtype=float)
    weighted = norm * wvec

    benefit_idx = [i for i, c in enumerate(pos.columns) if criteria_types.get(c, "max") == "max"]
    cost_idx = [i for i, c in enumerate(pos.columns) if criteria_types.get(c, "max") == "min"]

    s_plus = weighted[:, benefit_idx].sum(axis=1) if benefit_idx else np.zeros(len(pos))
    s_minus = weighted[:, cost_idx].sum(axis=1) if cost_idx else np.zeros(len(pos))

    if cost_idx and np.any(s_minus > EPS):
        s_safe = np.where(s_minus <= EPS, EPS, s_minus)
        s_min = float(np.min(s_safe))
        q = s_plus + _safe_divide(s_min * np.sum(s_safe), s_safe * np.sum(_safe_divide(s_min, s_safe)))
    else:
        q = s_plus.copy()

    score = _safe_divide(q - q.min(), q.max() - q.min())
    det = {
        "positive_shifts": shifts,
        "normalized_matrix": pd.DataFrame(norm, index=pos.index, columns=pos.columns),
        "weighted_matrix": pd.DataFrame(weighted, index=pos.index, columns=pos.columns),
        "copras_table": pd.DataFrame({
            "Alternatif": pos.index.astype(str),
            "S_plus": s_plus,
            "S_minus": s_minus,
            "Q_i": q,
            "Skor": score,
        }),
    }
    return score, det


def _rank_ocra(data: pd.DataFrame, criteria_types: Dict[str, str], weights: Dict[str, float]) -> Tuple[np.ndarray, Dict[str, Any]]:
    pos, shifts = _shift_positive(data)
    arr = pos.to_numpy(dtype=float)
    m, n = arr.shape
    wvec = np.asarray([weights[c] for c in pos.columns], dtype=float)

    benefit_mat = np.zeros((m, n), dtype=float)
    cost_mat = np.zeros((m, n), dtype=float)
    for j, c in enumerate(pos.columns):
        col = arr[:, j]
        cmin = float(col.min())
        cmax = float(col.max())
        denom = cmax - cmin
        if denom <= EPS:
            continue
        if criteria_types.get(c, "max") == "max":
            benefit_mat[:, j] = wvec[j] * (col - cmin) / denom
        else:
            cost_mat[:, j] = wvec[j] * (cmax - col) / denom

    benefit_score = benefit_mat.sum(axis=1)
    cost_score = cost_mat.sum(axis=1)
    raw = benefit_score + cost_score
    score = _safe_divide(raw - raw.min(), raw.max() - raw.min())
    det = {
        "positive_shifts": shifts,
        "benefit_component_matrix": pd.DataFrame(benefit_mat, index=pos.index, columns=pos.columns),
        "cost_component_matrix": pd.DataFrame(cost_mat, index=pos.index, columns=pos.columns),
        "ocra_table": pd.DataFrame({
            "Alternatif": pos.index.astype(str),
            "FaydaBileşeni": benefit_score,
            "MaliyetBileşeni": cost_score,
            "HamSkor": raw,
            "Skor": score,
        }),
    }
    return score, det


def _rank_aras(data: pd.DataFrame, criteria_types: Dict[str, str], weights: Dict[str, float]) -> Tuple[np.ndarray, Dict[str, Any]]:
    pos, shifts = _shift_positive(data)
    ideal_row = {c: pos[c].max() if criteria_types.get(c, "max") == "max" else pos[c].min() for c in pos.columns}
    aug = pd.concat([pd.DataFrame([ideal_row], index=["A0"]), pos], axis=0)
    norm = pd.DataFrame(index=aug.index, columns=aug.columns, dtype=float)
    for c in aug.columns:
        col = aug[c].to_numpy(dtype=float)
        if criteria_types.get(c, "max") == "max":
            norm[c] = _safe_divide(col, col.sum())
        else:
            inv = _safe_divide(1.0, col)
            norm[c] = _safe_divide(inv, inv.sum())
    wvec = np.asarray([weights[c] for c in aug.columns], dtype=float)
    weighted = norm.to_numpy(dtype=float) * wvec
    s = weighted.sum(axis=1)
    k = _safe_divide(s[1:], s[0])
    det = {
        "positive_shifts": shifts,
        "normalized_matrix": norm,
        "weighted_matrix": pd.DataFrame(weighted, index=aug.index, columns=aug.columns),
        "aras_table": pd.DataFrame({"Alternatif": pos.index.astype(str), "S_i": s[1:], "K_i": k, "Skor": k}),
    }
    return k, det


def _rank_saw(data: pd.DataFrame, criteria_types: Dict[str, str], weights: Dict[str, float]) -> Tuple[np.ndarray, Dict[str, Any]]:
    norm = _normalize_sum(data, criteria_types)
    wvec = np.asarray([weights[c] for c in norm.columns], dtype=float)
    raw = np.sum(norm.to_numpy(dtype=float) * wvec, axis=1)
    score = raw.copy()
    det = {
        "normalized_matrix": norm,
        "saw_table": pd.DataFrame({
            "Alternatif": data.index.astype(str),
            "AğırlıklıToplam": raw,
            "Skor": score,
        }),
    }
    return score, det


def _rank_wpm(data: pd.DataFrame, criteria_types: Dict[str, str], weights: Dict[str, float]) -> Tuple[np.ndarray, Dict[str, Any]]:
    norm = _normalize_minmax(data, criteria_types)
    wvec = np.asarray([weights[c] for c in norm.columns], dtype=float)
    raw = np.prod(np.power(norm.to_numpy(dtype=float) + EPS, wvec), axis=1)
    score = raw.copy()
    det = {
        "normalized_matrix": norm,
        "wpm_table": pd.DataFrame({
            "Alternatif": data.index.astype(str),
            "ÇarpımsalFayda": raw,
            "Skor": score,
        }),
    }
    return score, det


def _rank_maut(data: pd.DataFrame, criteria_types: Dict[str, str], weights: Dict[str, float]) -> Tuple[np.ndarray, Dict[str, Any]]:
    norm = _normalize_minmax(data, criteria_types)
    wvec = np.asarray([weights[c] for c in norm.columns], dtype=float)
    utility = norm.to_numpy(dtype=float) * wvec
    raw = utility.sum(axis=1)
    score = raw.copy()
    det = {
        "normalized_matrix": norm,
        "utility_matrix": pd.DataFrame(utility, index=data.index, columns=data.columns),
        "maut_table": pd.DataFrame({
            "Alternatif": data.index.astype(str),
            "ToplamFayda": raw,
            "Skor": score,
        }),
    }
    return score, det


def _rank_waspas(
    data: pd.DataFrame, criteria_types: Dict[str, str], weights: Dict[str, float], lambda_param: float = 0.5
) -> Tuple[np.ndarray, Dict[str, Any]]:
    norm = _normalize_minmax(data, criteria_types)
    wvec = np.asarray([weights[c] for c in norm.columns], dtype=float)
    arr = norm.to_numpy(dtype=float)
    q1 = np.sum(arr * wvec, axis=1)
    q2 = np.prod(np.power(arr + EPS, wvec), axis=1)
    score = lambda_param * q1 + (1.0 - lambda_param) * q2
    det = {
        "normalized_matrix": norm,
        "waspas_table": pd.DataFrame({"Alternatif": data.index.astype(str), "WSM": q1, "WPM": q2, "Skor": score}),
        "parameters": {"lambda": float(lambda_param)},
    }
    return score, det


def _rank_moora(data: pd.DataFrame, criteria_types: Dict[str, str], weights: Dict[str, float]) -> Tuple[np.ndarray, Dict[str, Any]]:
    norm = _normalize_vector(data)
    wvec = np.asarray([weights[c] for c in norm.columns], dtype=float)
    weighted = norm.to_numpy(dtype=float) * wvec
    benefit_idx = [i for i, c in enumerate(norm.columns) if criteria_types.get(c, "max") == "max"]
    cost_idx = [i for i, c in enumerate(norm.columns) if criteria_types.get(c, "max") == "min"]
    benefit_sum = weighted[:, benefit_idx].sum(axis=1) if benefit_idx else np.zeros(len(norm))
    cost_sum = weighted[:, cost_idx].sum(axis=1) if cost_idx else np.zeros(len(norm))
    raw = benefit_sum - cost_sum
    score = raw.copy()
    det = {
        "normalized_matrix": norm,
        "moora_table": pd.DataFrame({"Alternatif": data.index.astype(str), "FaydaToplamı": benefit_sum, "MaliyetToplamı": cost_sum, "HamSkor": raw, "Skor": score}),
    }
    return score, det


def _rank_mabac(data: pd.DataFrame, criteria_types: Dict[str, str], weights: Dict[str, float]) -> Tuple[np.ndarray, Dict[str, Any]]:
    norm = _normalize_minmax(data, criteria_types)
    wvec = np.asarray([weights[c] for c in norm.columns], dtype=float)
    v = (norm.to_numpy(dtype=float) + 1.0) * wvec
    g = np.prod(v, axis=0) ** (1.0 / max(len(norm), 1))
    q = v - g
    score_raw = q.sum(axis=1)
    score = _safe_divide(score_raw - score_raw.min(), score_raw.max() - score_raw.min())
    det = {
        "normalized_matrix": norm,
        "weighted_matrix": pd.DataFrame(v, index=data.index, columns=data.columns),
        "border_approximation": pd.DataFrame({"Kriter": data.columns, "g_j": g}),
        "mabac_table": pd.DataFrame({"Alternatif": data.index.astype(str), "HamSkor": score_raw, "Skor": score}),
    }
    return score, det


def _rank_marcos(data: pd.DataFrame, criteria_types: Dict[str, str], weights: Dict[str, float]) -> Tuple[np.ndarray, Dict[str, Any]]:
    pos, shifts = _shift_positive(data)
    anti_ideal = {c: pos[c].min() if criteria_types.get(c, "max") == "max" else pos[c].max() for c in pos.columns}
    ideal = {c: pos[c].max() if criteria_types.get(c, "max") == "max" else pos[c].min() for c in pos.columns}
    ext = pd.concat([pd.DataFrame([anti_ideal], index=["AAI"]), pos, pd.DataFrame([ideal], index=["AI"])], axis=0)
    norm = pd.DataFrame(index=ext.index, columns=ext.columns, dtype=float)
    for c in ext.columns:
        ai = float(ideal[c])
        col = ext[c].to_numpy(dtype=float)
        if criteria_types.get(c, "max") == "max":
            norm[c] = _safe_divide(col, ai)
        else:
            norm[c] = _safe_divide(ai, col)
    wvec = np.asarray([weights[c] for c in ext.columns], dtype=float)
    v = norm.to_numpy(dtype=float) * wvec
    s = v.sum(axis=1)
    s_aai = float(s[0])
    s_ai = float(s[-1])
    s_alt = s[1:-1]
    k_minus = _safe_divide(s_alt, s_aai)
    k_plus = _safe_divide(s_alt, s_ai)
    f_k_plus  = _safe_divide(k_plus,  k_plus + k_minus)
    f_k_minus = _safe_divide(k_minus, k_plus + k_minus)
    utility = _safe_divide(
        k_plus + k_minus,
        1.0 + _safe_divide(1.0 - f_k_plus, f_k_plus) + _safe_divide(1.0 - f_k_minus, f_k_minus),
    )
    score = _safe_divide(utility - utility.min(), utility.max() - utility.min())
    det = {
        "positive_shifts": shifts,
        "normalized_matrix": norm,
        "weighted_matrix": pd.DataFrame(v, index=ext.index, columns=ext.columns),
        "marcos_table": pd.DataFrame({
            "Alternatif": pos.index.astype(str),
            "S_i": s_alt,
            "K-": k_minus,
            "K+": k_plus,
            "U_i": utility,
            "Skor": score,
        }),
    }
    return score, det


def _rank_cocoso(
    data: pd.DataFrame,
    criteria_types: Dict[str, str],
    weights: Dict[str, float],
    lambda_param: float = 0.5,
) -> Tuple[np.ndarray, Dict[str, Any]]:
    norm = _normalize_minmax(data, criteria_types)
    arr = norm.to_numpy(dtype=float)
    wvec = np.asarray([weights[c] for c in norm.columns], dtype=float)
    s = np.sum(arr * wvec, axis=1)
    p = np.sum(np.power(arr + EPS, wvec), axis=1)
    ka = _safe_divide(s + p, np.sum(s + p))
    kb = _safe_divide(s, s.sum() + EPS) + _safe_divide(p, p.sum() + EPS)
    kc = _safe_divide(lambda_param * s + (1.0 - lambda_param) * p, lambda_param * s.max() + (1.0 - lambda_param) * p.max())
    score = np.cbrt(ka * kb * kc) + (ka + kb + kc) / 3.0
    score = _safe_divide(score - score.min(), score.max() - score.min())
    det = {
        "normalized_matrix": norm,
        "cocoso_table": pd.DataFrame({"Alternatif": data.index.astype(str), "S_i": s, "P_i": p, "K_a": ka, "K_b": kb, "K_c": kc, "Skor": score}),
        "parameters": {"lambda": float(lambda_param)},
    }
    return score, det


def _rank_promethee(
    data: pd.DataFrame,
    criteria_types: Dict[str, str],
    weights: Dict[str, float],
    pref_func: str = "linear",
    q: float = 0.05,
    p: float = 0.30,
    s: float = 0.20,
) -> Tuple[np.ndarray, Dict[str, Any]]:
    arr = data.to_numpy(dtype=float)
    m, n = arr.shape
    wvec = np.asarray([weights[c] for c in data.columns], dtype=float)
    ranges = np.ptp(arr, axis=0)
    ranges = np.where(ranges <= EPS, 1.0, ranges)
    qn = float(max(0.0, q))
    pn = float(max(qn + EPS, p))
    sn = float(max(EPS, s))

    direction = np.asarray(
        [1.0 if criteria_types.get(c, "max") == "max" else -1.0 for c in data.columns],
        dtype=float,
    )

    # Vectorized PROMETHEE preference matrix construction:
    # iterate criteria only, compute all (i,k) pairs at once.
    pref = np.zeros((m, m), dtype=float)
    unicriterion_pref = np.zeros((n, m, m), dtype=float)
    for j in range(n):
        d = (arr[:, j][:, np.newaxis] - arr[:, j][np.newaxis, :]) * direction[j]
        pj = np.maximum(0.0, d) / ranges[j]
        pj = np.clip(pj, 0.0, None)

        if pref_func == "usual":
            pj = (pj > 0).astype(float)
        elif pref_func == "u_shape":
            pj = np.where(pj <= qn, 0.0, 1.0)
        elif pref_func == "v_shape":
            pj = np.where(pj <= 0, 0.0, np.where(pj >= pn, 1.0, pj / pn))
        elif pref_func == "level":
            pj = np.where(pj <= qn, 0.0, np.where(pj <= pn, 0.5, 1.0))
        elif pref_func == "gaussian":
            pj = np.where(pj <= 0, 0.0, 1.0 - np.exp(-(pj ** 2) / (2.0 * (sn ** 2))))
        else:  # linear (V-shape with indifference)
            pj = np.where(pj <= qn, 0.0, np.where(pj >= pn, 1.0, (pj - qn) / (pn - qn + EPS)))

        pj = np.clip(pj, 0.0, 1.0)
        np.fill_diagonal(pj, 0.0)
        unicriterion_pref[j] = pj
        pref += wvec[j] * pj

    np.fill_diagonal(pref, 0.0)

    denom = max(m - 1, 1)
    phi_plus = pref.sum(axis=1) / denom
    phi_minus = pref.sum(axis=0) / denom
    phi_net = phi_plus - phi_minus
    score = _safe_divide(phi_net - phi_net.min(), phi_net.max() - phi_net.min())

    # GAIA plane: build unicriterion net-flow matrix (alternatives x criteria),
    # then project into 2D with PCA.
    phi_criterion = np.zeros((m, n), dtype=float)
    for j in range(n):
        pj = unicriterion_pref[j]
        phi_criterion[:, j] = (pj.sum(axis=1) - pj.sum(axis=0)) / denom

    gaia_alt_df = pd.DataFrame()
    gaia_crit_df = pd.DataFrame()
    gaia_axes_df = pd.DataFrame()
    gaia_decision_df = pd.DataFrame()
    try:
        if m >= 2 and n >= 1 and np.isfinite(phi_criterion).any():
            from sklearn.decomposition import PCA
            n_comp = min(2, m, n)
            pca = PCA(n_components=n_comp)
            alt_coords = pca.fit_transform(phi_criterion)
            crit_coords = pca.components_.T * np.sqrt(np.maximum(pca.explained_variance_, EPS))
            expl = np.asarray(pca.explained_variance_ratio_, dtype=float)

            if alt_coords.shape[1] == 1:
                alt_coords = np.column_stack([alt_coords[:, 0], np.zeros(m)])
            if crit_coords.shape[1] == 1:
                crit_coords = np.column_stack([crit_coords[:, 0], np.zeros(n)])
            if expl.size == 1:
                expl = np.array([expl[0], 0.0], dtype=float)

            decision_vec = np.sum(wvec[:, np.newaxis] * crit_coords[:, :2], axis=0)
            crit_norms = np.sqrt(np.sum(np.square(crit_coords[:, :2]), axis=1))
            target_len = float(np.nanmax(crit_norms)) if crit_norms.size else 0.0
            d_norm = float(np.linalg.norm(decision_vec))
            if d_norm > EPS and target_len > EPS:
                decision_vec = (decision_vec / d_norm) * (1.15 * target_len)
            else:
                decision_vec = np.zeros(2, dtype=float)

            gaia_alt_df = pd.DataFrame({
                "Alternatif": data.index.astype(str),
                "GAIA1": alt_coords[:, 0],
                "GAIA2": alt_coords[:, 1],
                "PhiNet": phi_net,
                "Skor": score,
            })
            gaia_crit_df = pd.DataFrame({
                "Kriter": data.columns.astype(str),
                "GAIA1": crit_coords[:, 0],
                "GAIA2": crit_coords[:, 1],
                "Ağırlık": wvec,
            })
            gaia_axes_df = pd.DataFrame({
                "Bileşen": ["GAIA1", "GAIA2"],
                "AçıklananVaryansOranı": [float(expl[0]), float(expl[1])],
            })
            gaia_decision_df = pd.DataFrame({
                "DeltaX": [float(decision_vec[0])],
                "DeltaY": [float(decision_vec[1])],
            })
    except Exception:
        # Keep PROMETHEE ranking robust even if GAIA projection fails.
        gaia_alt_df = pd.DataFrame()
        gaia_crit_df = pd.DataFrame()
        gaia_axes_df = pd.DataFrame()
        gaia_decision_df = pd.DataFrame()

    det = {
        "promethee_pref_matrix": pd.DataFrame(pref, index=data.index.astype(str), columns=data.index.astype(str)),
        "promethee_flows": pd.DataFrame({
            "Alternatif": data.index.astype(str),
            "PhiPlus": phi_plus,
            "PhiMinus": phi_minus,
            "PhiNet": phi_net,
            "Skor": score,
        }).sort_values("PhiNet", ascending=False).reset_index(drop=True),
        "promethee_gaia_alternatives": gaia_alt_df,
        "promethee_gaia_criteria": gaia_crit_df,
        "promethee_gaia_axes": gaia_axes_df,
        "promethee_gaia_decision_axis": gaia_decision_df,
        "parameters": {
            "pref_func": str(pref_func),
            "q": float(q),
            "p": float(p),
            "s": float(s),
        },
    }
    return score, det


def _rank_gra(data: pd.DataFrame, criteria_types: Dict[str, str], weights: Dict[str, float], rho: float = 0.5) -> Tuple[np.ndarray, Dict[str, Any]]:
    norm = _normalize_minmax(data, criteria_types)
    arr = norm.to_numpy(dtype=float)
    delta = np.abs(1.0 - arr)
    dmin = float(delta.min())
    dmax = float(delta.max())
    coeff = (dmin + rho * dmax) / (delta + rho * dmax + EPS)
    wvec = np.asarray([weights[c] for c in norm.columns], dtype=float)
    grade = np.sum(coeff * wvec, axis=1)
    score = _safe_divide(grade - grade.min(), grade.max() - grade.min())
    det = {
        "normalized_matrix": norm,
        "grey_coefficients": pd.DataFrame(coeff, index=data.index, columns=data.columns),
        "gra_table": pd.DataFrame({"Alternatif": data.index.astype(str), "İlişkiselDerece": grade, "Skor": score}),
        "parameters": {"rho": float(rho)},
    }
    return score, det


def _triangular_fuzzy_from_crisp(data: pd.DataFrame, spread: float = 0.10) -> np.ndarray:
    x = data.to_numpy(dtype=float)
    eff_spread = max(float(spread), 0.0)
    if eff_spread <= 0.0:
        delta = np.zeros_like(x, dtype=float)
    else:
        delta = np.abs(x) * eff_spread
    low = x - delta
    mid = x
    high = x + delta
    return np.stack([low, mid, high], axis=2)


def _shift_positive_tfn(tfn: np.ndarray) -> Tuple[np.ndarray, float]:
    minv = float(tfn[:, :, 0].min())
    if minv > 0:
        return tfn.copy(), 0.0
    magnitude = max(abs(minv), 1.0)
    delta = (abs(minv) * (1.0 + EPS)) if abs(minv) > EPS else (magnitude * EPS)
    return tfn + delta, delta


def _normalize_fuzzy_tfn(tfn: np.ndarray, criteria: Sequence[str], criteria_types: Dict[str, str]) -> np.ndarray:
    fuzzy = tfn.copy().astype(float)
    fuzzy, _ = _shift_positive_tfn(fuzzy)
    out = np.zeros_like(fuzzy)
    for j, c in enumerate(criteria):
        col = fuzzy[:, j, :]
        if criteria_types.get(c, "max") == "max":
            max_u = float(col[:, 2].max())
            out[:, j, 0] = col[:, 0] / (max_u + EPS)
            out[:, j, 1] = col[:, 1] / (max_u + EPS)
            out[:, j, 2] = col[:, 2] / (max_u + EPS)
        else:
            min_l = float(col[:, 0].min())
            out[:, j, 0] = min_l / (col[:, 2] + EPS)
            out[:, j, 1] = min_l / (col[:, 1] + EPS)
            out[:, j, 2] = min_l / (col[:, 0] + EPS)
    return out


def _defuzzify_tfn(tfn: np.ndarray) -> np.ndarray:
    return np.mean(tfn, axis=2)


def _fuzzy_distance(a: np.ndarray, b: np.ndarray) -> np.ndarray:
    return np.sqrt(np.sum((a - b) ** 2, axis=-1) / 3.0)


def _rank_fuzzy_topsis(
    data: pd.DataFrame, criteria_types: Dict[str, str], weights: Dict[str, float], spread: float = 0.10
) -> Tuple[np.ndarray, Dict[str, Any]]:
    if float(spread) <= 0.0:
        score, det = _rank_topsis(data, criteria_types, weights)
        det = dict(det)
        params = dict(det.get("parameters", {}) if isinstance(det.get("parameters"), dict) else {})
        params["spread"] = 0.0
        det["parameters"] = params
        return score, det
    criteria = list(data.columns)
    tfn = _triangular_fuzzy_from_crisp(data, spread)
    norm = _normalize_fuzzy_tfn(tfn, criteria, criteria_types)
    wvec = np.asarray([weights[c] for c in criteria], dtype=float)
    weighted = norm * wvec.reshape(1, -1, 1)
    fpis = np.max(weighted, axis=0)
    fnis = np.min(weighted, axis=0)
    dplus = np.sum(_fuzzy_distance(weighted, fpis[np.newaxis, :, :]), axis=1)
    dminus = np.sum(_fuzzy_distance(weighted, fnis[np.newaxis, :, :]), axis=1)
    score = dminus / (dplus + dminus + EPS)
    det = {
        "fuzzy_matrix": tfn,
        "normalized_fuzzy_matrix": norm,
        "defuzzified_matrix": pd.DataFrame(_defuzzify_tfn(weighted), index=data.index, columns=criteria),
        "fuzzy_topsis_table": pd.DataFrame({"Alternatif": data.index.astype(str), "D+": dplus, "D-": dminus, "Skor": score}),
        "parameters": {"spread": float(spread)},
    }
    return score, det


def _rank_fuzzy_by_scenarios(
    data: pd.DataFrame,
    criteria_types: Dict[str, str],
    weights: Dict[str, float],
    spread: float,
    ranker,
    *,
    extra_params: Optional[Dict[str, float]] = None,
) -> Tuple[np.ndarray, Dict[str, Any]]:
    if float(spread) <= 0.0:
        score, det = ranker(data, criteria_types, weights)
        det = dict(det) if isinstance(det, dict) else {}
        merged_params = dict(det.get("parameters", {}) if isinstance(det.get("parameters"), dict) else {})
        merged_params["spread"] = 0.0
        if extra_params:
            merged_params.update(extra_params)
        det["parameters"] = merged_params
        return score, det
    criteria = list(data.columns)
    tfn = _triangular_fuzzy_from_crisp(data, spread)
    labels = ("Lower", "Middle", "Upper")
    scenario_scores: Dict[str, np.ndarray] = {}
    scenario_details: Dict[str, Dict[str, Any]] = {}

    for idx, label in enumerate(labels):
        scenario_df = pd.DataFrame(tfn[:, :, idx], index=data.index, columns=criteria)
        sc, det = ranker(scenario_df, criteria_types, weights)
        scenario_scores[label] = np.asarray(sc, dtype=float).reshape(-1)
        scenario_details[label] = det

    stacked = np.vstack([scenario_scores["Lower"], scenario_scores["Middle"], scenario_scores["Upper"]])
    score = np.mean(stacked, axis=0)
    middle_det = scenario_details.get("Middle", {})
    det: Dict[str, Any] = dict(middle_det) if isinstance(middle_det, dict) else {}
    det["fuzzy_matrix"] = tfn
    det["fuzzy_scenario_scores"] = pd.DataFrame(
        {
            "Alternatif": data.index.astype(str),
            "LowerSkor": scenario_scores["Lower"],
            "MiddleSkor": scenario_scores["Middle"],
            "UpperSkor": scenario_scores["Upper"],
            "Skor": score,
        }
    )
    det["fuzzy_scenario_method_details"] = scenario_details
    merged_params = dict(det.get("parameters", {}) if isinstance(det.get("parameters"), dict) else {})
    merged_params["spread"] = float(spread)
    if extra_params:
        merged_params.update(extra_params)
    det["parameters"] = merged_params
    return score, det


def _rank_fuzzy_vikor(
    data: pd.DataFrame,
    criteria_types: Dict[str, str],
    weights: Dict[str, float],
    spread: float = 0.10,
    v_param: float = 0.5,
) -> Tuple[np.ndarray, Dict[str, Any]]:
    return _rank_fuzzy_by_scenarios(
        data,
        criteria_types,
        weights,
        spread,
        lambda d, ct, w: _rank_vikor(d, ct, w, v_param=v_param),
        extra_params={"v_param": float(v_param)},
    )


def _rank_fuzzy_aras(
    data: pd.DataFrame, criteria_types: Dict[str, str], weights: Dict[str, float], spread: float = 0.10
) -> Tuple[np.ndarray, Dict[str, Any]]:
    return _rank_fuzzy_by_scenarios(data, criteria_types, weights, spread, _rank_aras)


def _rank_fuzzy_saw(
    data: pd.DataFrame, criteria_types: Dict[str, str], weights: Dict[str, float], spread: float = 0.10
) -> Tuple[np.ndarray, Dict[str, Any]]:
    return _rank_fuzzy_by_scenarios(data, criteria_types, weights, spread, _rank_saw)


def _rank_fuzzy_wpm(
    data: pd.DataFrame, criteria_types: Dict[str, str], weights: Dict[str, float], spread: float = 0.10
) -> Tuple[np.ndarray, Dict[str, Any]]:
    return _rank_fuzzy_by_scenarios(data, criteria_types, weights, spread, _rank_wpm)


def _rank_fuzzy_maut(
    data: pd.DataFrame, criteria_types: Dict[str, str], weights: Dict[str, float], spread: float = 0.10
) -> Tuple[np.ndarray, Dict[str, Any]]:
    return _rank_fuzzy_by_scenarios(data, criteria_types, weights, spread, _rank_maut)


def _rank_fuzzy_waspas(
    data: pd.DataFrame,
    criteria_types: Dict[str, str],
    weights: Dict[str, float],
    spread: float = 0.10,
    lambda_param: float = 0.5,
) -> Tuple[np.ndarray, Dict[str, Any]]:
    return _rank_fuzzy_by_scenarios(
        data,
        criteria_types,
        weights,
        spread,
        lambda d, ct, w: _rank_waspas(d, ct, w, lambda_param=lambda_param),
        extra_params={"lambda": float(lambda_param)},
    )


def _rank_fuzzy_edas(
    data: pd.DataFrame, criteria_types: Dict[str, str], weights: Dict[str, float], spread: float = 0.10
) -> Tuple[np.ndarray, Dict[str, Any]]:
    return _rank_fuzzy_by_scenarios(data, criteria_types, weights, spread, _rank_edas)


def _rank_fuzzy_codas(
    data: pd.DataFrame,
    criteria_types: Dict[str, str],
    weights: Dict[str, float],
    spread: float = 0.10,
    tau: float = 0.02,
) -> Tuple[np.ndarray, Dict[str, Any]]:
    return _rank_fuzzy_by_scenarios(
        data,
        criteria_types,
        weights,
        spread,
        lambda d, ct, w: _rank_codas(d, ct, w, tau=tau),
        extra_params={"tau": float(tau)},
    )


def _rank_fuzzy_copras(
    data: pd.DataFrame, criteria_types: Dict[str, str], weights: Dict[str, float], spread: float = 0.10
) -> Tuple[np.ndarray, Dict[str, Any]]:
    return _rank_fuzzy_by_scenarios(data, criteria_types, weights, spread, _rank_copras)


def _rank_fuzzy_ocra(
    data: pd.DataFrame, criteria_types: Dict[str, str], weights: Dict[str, float], spread: float = 0.10
) -> Tuple[np.ndarray, Dict[str, Any]]:
    return _rank_fuzzy_by_scenarios(data, criteria_types, weights, spread, _rank_ocra)


def _rank_fuzzy_moora(
    data: pd.DataFrame, criteria_types: Dict[str, str], weights: Dict[str, float], spread: float = 0.10
) -> Tuple[np.ndarray, Dict[str, Any]]:
    return _rank_fuzzy_by_scenarios(data, criteria_types, weights, spread, _rank_moora)


def _rank_fuzzy_mabac(
    data: pd.DataFrame, criteria_types: Dict[str, str], weights: Dict[str, float], spread: float = 0.10
) -> Tuple[np.ndarray, Dict[str, Any]]:
    return _rank_fuzzy_by_scenarios(data, criteria_types, weights, spread, _rank_mabac)


def _rank_fuzzy_marcos(
    data: pd.DataFrame, criteria_types: Dict[str, str], weights: Dict[str, float], spread: float = 0.10
) -> Tuple[np.ndarray, Dict[str, Any]]:
    return _rank_fuzzy_by_scenarios(data, criteria_types, weights, spread, _rank_marcos)


def _rank_fuzzy_cocoso(
    data: pd.DataFrame,
    criteria_types: Dict[str, str],
    weights: Dict[str, float],
    spread: float = 0.10,
    lambda_param: float = 0.5,
) -> Tuple[np.ndarray, Dict[str, Any]]:
    return _rank_fuzzy_by_scenarios(
        data,
        criteria_types,
        weights,
        spread,
        lambda d, ct, w: _rank_cocoso(d, ct, w, lambda_param=lambda_param),
        extra_params={"lambda": float(lambda_param)},
    )


def _rank_fuzzy_gra(
    data: pd.DataFrame,
    criteria_types: Dict[str, str],
    weights: Dict[str, float],
    spread: float = 0.10,
    rho: float = 0.5,
) -> Tuple[np.ndarray, Dict[str, Any]]:
    return _rank_fuzzy_by_scenarios(
        data,
        criteria_types,
        weights,
        spread,
        lambda d, ct, w: _rank_gra(d, ct, w, rho=rho),
        extra_params={"rho": float(rho)},
    )


def _rank_fuzzy_promethee(
    data: pd.DataFrame,
    criteria_types: Dict[str, str],
    weights: Dict[str, float],
    spread: float = 0.10,
    pref_func: str = "linear",
    q: float = 0.05,
    p: float = 0.30,
    s: float = 0.20,
) -> Tuple[np.ndarray, Dict[str, Any]]:
    return _rank_fuzzy_by_scenarios(
        data,
        criteria_types,
        weights,
        spread,
        lambda d, ct, w: _rank_promethee(d, ct, w, pref_func=pref_func, q=q, p=p, s=s),
        extra_params={
            "pref_func": str(pref_func),
            "q": float(q),
            "p": float(p),
            "s": float(s),
        },
    )


# ─────────────────────────────────────────────────────────────────────────────
# YENİ SIRALAMA YÖNTEMLERİ: SPOTIS, MULTIMOORA, RAWEC, RAFSI, ROV, AROMAN, DNMA
# ─────────────────────────────────────────────────────────────────────────────

def _rank_spotis(
    data: pd.DataFrame, criteria_types: Dict[str, str], weights: Dict[str, float]
) -> Tuple[np.ndarray, Dict[str, Any]]:
    """SPOTIS — Stable Preference Ordering Towards Ideal Solution (Dezert et al., 2020).
    Sabit ideal noktaya ağırlıklı normalleştirilmiş uzaklık; düşük = iyi.
    """
    x = data.to_numpy(dtype=float)
    wvec = np.asarray([weights[c] for c in data.columns], dtype=float)
    n = x.shape[1]
    ideal = np.zeros(n, dtype=float)
    spans = np.ones(n, dtype=float)
    for j, c in enumerate(data.columns):
        col = x[:, j]
        xmin, xmax = float(col.min()), float(col.max())
        ideal[j] = xmax if criteria_types.get(c, "max") == "max" else xmin
        span = xmax - xmin
        spans[j] = span if span > EPS else 1.0
    dist = np.abs(x - ideal) / spans
    score = dist @ wvec
    det = {
        "ideal_point": pd.DataFrame({"Kriter": data.columns, "İdeal": ideal, "Span": spans}),
        "distance_matrix": pd.DataFrame(dist, index=data.index, columns=data.columns),
        "spotis_table": pd.DataFrame({"Alternatif": data.index.astype(str), "Uzaklık": score}),
    }
    return score, det


def _rank_multimoora(
    data: pd.DataFrame, criteria_types: Dict[str, str], weights: Dict[str, float]
) -> Tuple[np.ndarray, Dict[str, Any]]:
    """MULTIMOORA — Multi-Objective Optimisation by Ratio Analysis + Full Multiplicative Form
    (Brauers & Zavadskas, 2010). RS + RP + FMF → Borda toplamı; düşük = iyi.
    """
    pos, _ = _shift_positive(data)
    x = pos.to_numpy(dtype=float)
    wvec = np.asarray([weights[c] for c in pos.columns], dtype=float)
    # Vektör normalleştirme
    norms = np.sqrt((x ** 2).sum(axis=0)) + EPS
    xn = x / norms
    # 1. Oran Sistemi (RS)
    rs = np.zeros(len(pos), dtype=float)
    for j, c in enumerate(pos.columns):
        rs += wvec[j] * xn[:, j] if criteria_types.get(c, "max") == "max" else -wvec[j] * xn[:, j]
    rs_rank = pd.Series(rs).rank(ascending=False, method="min").astype(int).to_numpy()
    # 2. Referans Nokta (RP)
    ref = np.array([
        xn[:, j].max() if criteria_types.get(c, "max") == "max" else xn[:, j].min()
        for j, c in enumerate(pos.columns)
    ], dtype=float)
    rp_score = np.max(wvec * np.abs(xn - ref), axis=1)
    rp_rank = pd.Series(rp_score).rank(ascending=True, method="min").astype(int).to_numpy()
    # 3. Tam Çarpımsal Form (FMF)
    benefit_prod = np.ones(len(pos), dtype=float)
    cost_prod = np.ones(len(pos), dtype=float)
    for j, c in enumerate(pos.columns):
        if criteria_types.get(c, "max") == "max":
            benefit_prod *= (xn[:, j] + EPS) ** wvec[j]
        else:
            cost_prod *= (xn[:, j] + EPS) ** wvec[j]
    fmf_score = benefit_prod / (cost_prod + EPS)
    fmf_rank = pd.Series(fmf_score).rank(ascending=False, method="min").astype(int).to_numpy()
    # Borda toplamı (düşük = iyi)
    borda = (rs_rank + rp_rank + fmf_rank).astype(float)
    det = {
        "normalized_matrix": pd.DataFrame(xn, index=pos.index, columns=pos.columns),
        "multimoora_table": pd.DataFrame({
            "Alternatif": pos.index.astype(str),
            "RS_Skor": rs, "RS_Sıra": rs_rank,
            "RP_Skor": rp_score, "RP_Sıra": rp_rank,
            "FMF_Skor": fmf_score, "FMF_Sıra": fmf_rank,
            "Borda": borda.astype(int),
        }),
    }
    return borda, det


def _rank_rawec(
    data: pd.DataFrame, criteria_types: Dict[str, str], weights: Dict[str, float]
) -> Tuple[np.ndarray, Dict[str, Any]]:
    """RAWEC — Ranking of Alternatives by Weight of Each Criterion (Sotoudeh-Anvari, 2023).
    Kriter bazlı sıraların ağırlıklı harmonik toplulaştırması; yüksek = iyi.
    """
    norm = _normalize_sum(data, criteria_types)
    wvec = np.asarray([weights[c] for c in data.columns], dtype=float)
    wmat = norm.to_numpy(dtype=float) * wvec
    # Her kriter için alternatif sıralaması (1 = en iyi / en yüksek ağırlıklı değer)
    ranks = pd.DataFrame(wmat, index=data.index, columns=data.columns).rank(
        ascending=False, method="min"
    ).to_numpy(dtype=float)
    score = (wvec / ranks).sum(axis=1)
    det = {
        "normalized_matrix": norm,
        "weighted_matrix": pd.DataFrame(wmat, index=data.index, columns=data.columns),
        "rank_matrix": pd.DataFrame(ranks.astype(int), index=data.index, columns=data.columns),
        "rawec_table": pd.DataFrame({"Alternatif": data.index.astype(str), "Skor": score}),
    }
    return score, det


def _rank_rafsi(
    data: pd.DataFrame, criteria_types: Dict[str, str], weights: Dict[str, float],
    r1: float = 1.0, r2: float = 9.0,
) -> Tuple[np.ndarray, Dict[str, Any]]:
    """RAFSI — Ranking of Alternatives through Functional mapping of criterion Sub-Intervals
    into a Single Interval (Žižović et al., 2020). Sabit [r1, r2] aralığına eşleme; yüksek = iyi.
    """
    x = data.to_numpy(dtype=float)
    wvec = np.asarray([weights[c] for c in data.columns], dtype=float)
    span_ref = r2 - r1
    h = np.zeros_like(x)
    for j, c in enumerate(data.columns):
        col = x[:, j]
        xmin, xmax = float(col.min()), float(col.max())
        denom = xmax - xmin
        if denom <= EPS:
            h[:, j] = (r1 + r2) / 2.0
            continue
        if criteria_types.get(c, "max") == "max":
            h[:, j] = r1 + span_ref * (col - xmin) / denom
        else:
            h[:, j] = r1 + span_ref * (xmax - col) / denom
    score = h @ wvec
    det = {
        "reference_interval": {"r1": r1, "r2": r2},
        "mapped_matrix": pd.DataFrame(h, index=data.index, columns=data.columns),
        "rafsi_table": pd.DataFrame({"Alternatif": data.index.astype(str), "Skor": score}),
    }
    return score, det


def _rank_rov(
    data: pd.DataFrame, criteria_types: Dict[str, str], weights: Dict[str, float]
) -> Tuple[np.ndarray, Dict[str, Any]]:
    """ROV — Range of Value (Yakowitz et al., 1993).
    Min-max normalleştirme + ağırlıklı toplam; yüksek = iyi.
    """
    norm = _normalize_minmax(data, criteria_types)
    wvec = np.asarray([weights[c] for c in data.columns], dtype=float)
    score = norm.to_numpy(dtype=float) @ wvec
    det = {
        "normalized_matrix": norm,
        "rov_table": pd.DataFrame({"Alternatif": data.index.astype(str), "Skor": score}),
    }
    return score, det


def _rank_aroman(
    data: pd.DataFrame, criteria_types: Dict[str, str], weights: Dict[str, float]
) -> Tuple[np.ndarray, Dict[str, Any]]:
    """AROMAN — Alternative Ranking Order Method Accounting for Two-step Normalization
    (Dimitrijević et al., 2022). Sum + min-max normalleştirmenin geometrik bileşimi; yüksek = iyi.
    """
    pos, _ = _shift_positive(data)
    r1 = _normalize_sum(pos, criteria_types).to_numpy(dtype=float)
    r2 = _normalize_minmax(pos, criteria_types).to_numpy(dtype=float)
    # Geometrik birleştirme
    h = np.sqrt(np.clip(r1 * r2, 0.0, None) + EPS) - np.sqrt(EPS)
    wvec = np.asarray([weights[c] for c in pos.columns], dtype=float)
    score = h @ wvec
    det = {
        "sum_normalized": pd.DataFrame(r1, index=pos.index, columns=pos.columns),
        "minmax_normalized": pd.DataFrame(r2, index=pos.index, columns=pos.columns),
        "combined_matrix": pd.DataFrame(h, index=pos.index, columns=pos.columns),
        "aroman_table": pd.DataFrame({"Alternatif": pos.index.astype(str), "Skor": score}),
    }
    return score, det


def _rank_dnma(
    data: pd.DataFrame, criteria_types: Dict[str, str], weights: Dict[str, float],
    alpha: float = 0.5,
) -> Tuple[np.ndarray, Dict[str, Any]]:
    """DNMA — Double Normalization-based Multiple Aggregation (Liu & Zhu, 2021).
    Min-max ve sum normalleştirme skorlarının ağırlıklı ortalaması; yüksek = iyi.
    """
    r1 = _normalize_minmax(data, criteria_types).to_numpy(dtype=float)
    r2 = _normalize_sum(data, criteria_types).to_numpy(dtype=float)
    wvec = np.asarray([weights[c] for c in data.columns], dtype=float)
    s1 = r1 @ wvec
    s2 = r2 @ wvec
    score = alpha * s1 + (1.0 - alpha) * s2
    det = {
        "minmax_normalized": pd.DataFrame(r1, index=data.index, columns=data.columns),
        "sum_normalized": pd.DataFrame(r2, index=data.index, columns=data.columns),
        "dnma_table": pd.DataFrame({
            "Alternatif": data.index.astype(str),
            "S_MinMax": s1, "S_Sum": s2, "Skor": score,
        }),
        "parameters": {"alpha": alpha},
    }
    return score, det


# Bulanık sarmalayıcılar (üçgensel BDS senaryoları üzerinden)

def _rank_fuzzy_spotis(
    data: pd.DataFrame, criteria_types: Dict[str, str], weights: Dict[str, float], spread: float = 0.10,
) -> Tuple[np.ndarray, Dict[str, Any]]:
    return _rank_fuzzy_by_scenarios(data, criteria_types, weights, spread, _rank_spotis)


def _rank_fuzzy_multimoora(
    data: pd.DataFrame, criteria_types: Dict[str, str], weights: Dict[str, float], spread: float = 0.10,
) -> Tuple[np.ndarray, Dict[str, Any]]:
    return _rank_fuzzy_by_scenarios(data, criteria_types, weights, spread, _rank_multimoora)


def _rank_fuzzy_rawec(
    data: pd.DataFrame, criteria_types: Dict[str, str], weights: Dict[str, float], spread: float = 0.10,
) -> Tuple[np.ndarray, Dict[str, Any]]:
    return _rank_fuzzy_by_scenarios(data, criteria_types, weights, spread, _rank_rawec)


def _rank_fuzzy_rafsi(
    data: pd.DataFrame, criteria_types: Dict[str, str], weights: Dict[str, float], spread: float = 0.10,
) -> Tuple[np.ndarray, Dict[str, Any]]:
    return _rank_fuzzy_by_scenarios(data, criteria_types, weights, spread, _rank_rafsi)


def _rank_fuzzy_rov(
    data: pd.DataFrame, criteria_types: Dict[str, str], weights: Dict[str, float], spread: float = 0.10,
) -> Tuple[np.ndarray, Dict[str, Any]]:
    return _rank_fuzzy_by_scenarios(data, criteria_types, weights, spread, _rank_rov)


def _rank_fuzzy_aroman(
    data: pd.DataFrame, criteria_types: Dict[str, str], weights: Dict[str, float], spread: float = 0.10,
) -> Tuple[np.ndarray, Dict[str, Any]]:
    return _rank_fuzzy_by_scenarios(data, criteria_types, weights, spread, _rank_aroman)


def _rank_fuzzy_dnma(
    data: pd.DataFrame, criteria_types: Dict[str, str], weights: Dict[str, float], spread: float = 0.10,
) -> Tuple[np.ndarray, Dict[str, Any]]:
    return _rank_fuzzy_by_scenarios(data, criteria_types, weights, spread, _rank_dnma)


def rank_alternatives(
    data: pd.DataFrame,
    criteria: Sequence[str],
    criteria_types: Dict[str, str],
    weights: Dict[str, float],
    method: str,
    *,
    vikor_v: float = 0.5,
    waspas_lambda: float = 0.5,
    codas_tau: float = 0.02,
    cocoso_lambda: float = 0.5,
    gra_rho: float = 0.5,
    promethee_pref_func: str = "linear",
    promethee_q: float = 0.05,
    promethee_p: float = 0.30,
    promethee_s: float = 0.20,
    fuzzy_spread: float = 0.10,
) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    df = _as_numeric_df(data, criteria)
    # Input weight validation
    w_vec = np.asarray([float(weights.get(c, 0.0)) for c in criteria], dtype=float)
    if np.any(w_vec < 0):
        w_vec = np.clip(w_vec, 0.0, None)
        w_sum = w_vec.sum()
        if w_sum > EPS:
            w_vec /= w_sum
        weights = dict(zip(criteria, w_vec))
    w_sum = sum(weights.get(c, 0.0) for c in criteria)
    if w_sum > EPS and not np.isclose(w_sum, 1.0, atol=1e-4):
        weights = {c: weights.get(c, 0.0) / w_sum for c in criteria}
    dispatch = {
        "TOPSIS": lambda: _rank_topsis(df, criteria_types, weights),
        "VIKOR": lambda: _rank_vikor(df, criteria_types, weights, v_param=vikor_v),
        "EDAS": lambda: _rank_edas(df, criteria_types, weights),
        "CODAS": lambda: _rank_codas(df, criteria_types, weights, tau=codas_tau),
        "COPRAS": lambda: _rank_copras(df, criteria_types, weights),
        "OCRA": lambda: _rank_ocra(df, criteria_types, weights),
        "ARAS": lambda: _rank_aras(df, criteria_types, weights),
        "SAW": lambda: _rank_saw(df, criteria_types, weights),
        "WPM": lambda: _rank_wpm(df, criteria_types, weights),
        "MAUT": lambda: _rank_maut(df, criteria_types, weights),
        "WASPAS": lambda: _rank_waspas(df, criteria_types, weights, lambda_param=waspas_lambda),
        "MOORA": lambda: _rank_moora(df, criteria_types, weights),
        "MABAC": lambda: _rank_mabac(df, criteria_types, weights),
        "MARCOS": lambda: _rank_marcos(df, criteria_types, weights),
        "CoCoSo": lambda: _rank_cocoso(df, criteria_types, weights, lambda_param=cocoso_lambda),
        "PROMETHEE": lambda: _rank_promethee(
            df, criteria_types, weights,
            pref_func=promethee_pref_func, q=promethee_q, p=promethee_p, s=promethee_s,
        ),
        "GRA": lambda: _rank_gra(df, criteria_types, weights, rho=gra_rho),
        "Fuzzy TOPSIS": lambda: _rank_fuzzy_topsis(df, criteria_types, weights, spread=fuzzy_spread),
        "Fuzzy VIKOR": lambda: _rank_fuzzy_vikor(df, criteria_types, weights, spread=fuzzy_spread, v_param=vikor_v),
        "Fuzzy ARAS": lambda: _rank_fuzzy_aras(df, criteria_types, weights, spread=fuzzy_spread),
        "Fuzzy SAW": lambda: _rank_fuzzy_saw(df, criteria_types, weights, spread=fuzzy_spread),
        "Fuzzy WPM": lambda: _rank_fuzzy_wpm(df, criteria_types, weights, spread=fuzzy_spread),
        "Fuzzy MAUT": lambda: _rank_fuzzy_maut(df, criteria_types, weights, spread=fuzzy_spread),
        "Fuzzy WASPAS": lambda: _rank_fuzzy_waspas(df, criteria_types, weights, spread=fuzzy_spread, lambda_param=waspas_lambda),
        "Fuzzy EDAS": lambda: _rank_fuzzy_edas(df, criteria_types, weights, spread=fuzzy_spread),
        "Fuzzy CODAS": lambda: _rank_fuzzy_codas(df, criteria_types, weights, spread=fuzzy_spread, tau=codas_tau),
        "Fuzzy COPRAS": lambda: _rank_fuzzy_copras(df, criteria_types, weights, spread=fuzzy_spread),
        "Fuzzy OCRA": lambda: _rank_fuzzy_ocra(df, criteria_types, weights, spread=fuzzy_spread),
        "Fuzzy MOORA": lambda: _rank_fuzzy_moora(df, criteria_types, weights, spread=fuzzy_spread),
        "Fuzzy MABAC": lambda: _rank_fuzzy_mabac(df, criteria_types, weights, spread=fuzzy_spread),
        "Fuzzy MARCOS": lambda: _rank_fuzzy_marcos(df, criteria_types, weights, spread=fuzzy_spread),
        "Fuzzy CoCoSo": lambda: _rank_fuzzy_cocoso(df, criteria_types, weights, spread=fuzzy_spread, lambda_param=cocoso_lambda),
        "Fuzzy GRA": lambda: _rank_fuzzy_gra(df, criteria_types, weights, spread=fuzzy_spread, rho=gra_rho),
        "Fuzzy PROMETHEE": lambda: _rank_fuzzy_promethee(
            df, criteria_types, weights, spread=fuzzy_spread,
            pref_func=promethee_pref_func, q=promethee_q, p=promethee_p, s=promethee_s,
        ),
        # Yeni klasik yöntemler
        "SPOTIS": lambda: _rank_spotis(df, criteria_types, weights),
        "MULTIMOORA": lambda: _rank_multimoora(df, criteria_types, weights),
        "RAWEC": lambda: _rank_rawec(df, criteria_types, weights),
        "RAFSI": lambda: _rank_rafsi(df, criteria_types, weights),
        "ROV": lambda: _rank_rov(df, criteria_types, weights),
        "AROMAN": lambda: _rank_aroman(df, criteria_types, weights),
        "DNMA": lambda: _rank_dnma(df, criteria_types, weights),
        # Yeni bulanık yöntemler
        "Fuzzy SPOTIS": lambda: _rank_fuzzy_spotis(df, criteria_types, weights, spread=fuzzy_spread),
        "Fuzzy MULTIMOORA": lambda: _rank_fuzzy_multimoora(df, criteria_types, weights, spread=fuzzy_spread),
        "Fuzzy RAWEC": lambda: _rank_fuzzy_rawec(df, criteria_types, weights, spread=fuzzy_spread),
        "Fuzzy RAFSI": lambda: _rank_fuzzy_rafsi(df, criteria_types, weights, spread=fuzzy_spread),
        "Fuzzy ROV": lambda: _rank_fuzzy_rov(df, criteria_types, weights, spread=fuzzy_spread),
        "Fuzzy AROMAN": lambda: _rank_fuzzy_aroman(df, criteria_types, weights, spread=fuzzy_spread),
        "Fuzzy DNMA": lambda: _rank_fuzzy_dnma(df, criteria_types, weights, spread=fuzzy_spread),
    }
    if method not in dispatch:
        raise ValueError(f"Desteklenmeyen sıralama yöntemi: {method}")
    scores, details = dispatch[method]()
    scores_arr = np.asarray(scores, dtype=float).reshape(-1)
    if len(scores_arr) != len(df):
        raise ValueError(f"{method} yöntemi beklenmeyen skor uzunluğu üretti.")
    if (~np.isfinite(scores_arr)).any():
        raise ValueError(f"{method} yöntemi geçersiz skor (NaN/Inf) üretti.")
    lower_is_better = _is_lower_better_method(method)
    result = _vector_rank(scores_arr, df.index.astype(str), ascending=lower_is_better)
    details["score_direction"] = "min" if lower_is_better else "max"
    details["result_table"] = result
    return result, details


def compare_methods(
    data: pd.DataFrame,
    criteria: Sequence[str],
    criteria_types: Dict[str, str],
    weights: Dict[str, float],
    methods: Sequence[str],
    *,
    vikor_v: float = 0.5,
    waspas_lambda: float = 0.5,
    codas_tau: float = 0.02,
    cocoso_lambda: float = 0.5,
    gra_rho: float = 0.5,
    promethee_pref_func: str = "linear",
    promethee_q: float = 0.05,
    promethee_p: float = 0.30,
    promethee_s: float = 0.20,
    fuzzy_spread: float = 0.10,
    base_method: Optional[str] = None,
    base_table: Optional[pd.DataFrame] = None,
    base_details: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    from scipy.stats import spearmanr
    if not methods:
        return {}
    score_frames: List[pd.DataFrame] = []
    rank_frames: List[pd.DataFrame] = []
    method_tables: Dict[str, pd.DataFrame] = {}
    method_details: Dict[str, Dict[str, Any]] = {}
    for method in methods:
        if (
            base_method
            and method == base_method
            and isinstance(base_table, pd.DataFrame)
            and not base_table.empty
        ):
            res = base_table.copy()
            det = dict(base_details or {})
        else:
            res, det = rank_alternatives(
                data,
                criteria,
                criteria_types,
                weights,
                method,
                vikor_v=vikor_v,
                waspas_lambda=waspas_lambda,
                codas_tau=codas_tau,
                cocoso_lambda=cocoso_lambda,
                gra_rho=gra_rho,
                promethee_pref_func=promethee_pref_func,
                promethee_q=promethee_q,
                promethee_p=promethee_p,
                promethee_s=promethee_s,
                fuzzy_spread=fuzzy_spread,
            )
        method_tables[str(method)] = res.copy()
        method_details[str(method)] = det
        s = res[["Alternatif", "Skor"]].rename(columns={"Skor": method})
        r = res[["Alternatif", "Sıra"]].rename(columns={"Sıra": method})
        score_frames.append(s)
        rank_frames.append(r)
    score_table = score_frames[0]
    rank_table = rank_frames[0]
    for frame in score_frames[1:]:
        score_table = score_table.merge(frame, on="Alternatif", how="outer")
    for frame in rank_frames[1:]:
        rank_table = rank_table.merge(frame, on="Alternatif", how="outer")
    method_names = list(methods)
    corr = pd.DataFrame(index=method_names, columns=method_names, dtype=float)
    for m1 in method_names:
        for m2 in method_names:
            try:
                rho, _ = spearmanr(rank_table[m1], rank_table[m2])
                corr.loc[m1, m2] = rho if np.isfinite(rho) else 1.0 if m1 == m2 else np.nan
            except (ValueError, TypeError):
                corr.loc[m1, m2] = 1.0 if m1 == m2 else np.nan
    best_counts = []
    for method in method_names:
        top_alt = rank_table.sort_values(method).iloc[0]["Alternatif"]
        best_counts.append({"Yöntem": method, "BirinciAlternatif": top_alt})
    return {
        "score_table": score_table,
        "rank_table": rank_table,
        "spearman_matrix": corr.reset_index().rename(columns={"index": "Yöntem"}),
        "top_alternatives": pd.DataFrame(best_counts),
        "method_tables": method_tables,
        "method_details": method_details,
    }


def sensitivity_analysis(
    data: pd.DataFrame,
    criteria: Sequence[str],
    criteria_types: Dict[str, str],
    base_weights: Dict[str, float],
    ranking_method: str,
    *,
    vikor_v: float = 0.5,
    waspas_lambda: float = 0.5,
    codas_tau: float = 0.02,
    cocoso_lambda: float = 0.5,
    gra_rho: float = 0.5,
    promethee_pref_func: str = "linear",
    promethee_q: float = 0.05,
    promethee_p: float = 0.30,
    promethee_s: float = 0.20,
    fuzzy_spread: float = 0.10,
    iterations: int = 200,
    sigma: float = 0.12,
    seed: int = 42,
) -> Dict[str, Any]:
    from scipy.stats import spearmanr
    base_res, _ = rank_alternatives(
        data,
        criteria,
        criteria_types,
        base_weights,
        ranking_method,
        vikor_v=vikor_v,
        waspas_lambda=waspas_lambda,
        codas_tau=codas_tau,
        cocoso_lambda=cocoso_lambda,
        gra_rho=gra_rho,
        promethee_pref_func=promethee_pref_func,
        promethee_q=promethee_q,
        promethee_p=promethee_p,
        promethee_s=promethee_s,
        fuzzy_spread=fuzzy_spread,
    )
    base_rank = base_res.set_index("Alternatif")["Sıra"]
    base_top = base_res.iloc[0]["Alternatif"]
    weight_vec = np.asarray([base_weights[c] for c in criteria], dtype=float)
    rng = np.random.default_rng(seed)

    scenario_rows = []
    for crit_idx, crit in enumerate(criteria):
        for delta in [-0.20, -0.10, 0.10, 0.20]:
            new_w = weight_vec.copy()
            new_w[crit_idx] = max(EPS, new_w[crit_idx] * (1.0 + delta))
            new_w = _normalize_weights(new_w)
            wd = dict(zip(criteria, new_w))
            res, _ = rank_alternatives(
                data,
                criteria,
                criteria_types,
                wd,
                ranking_method,
                vikor_v=vikor_v,
                waspas_lambda=waspas_lambda,
                codas_tau=codas_tau,
                cocoso_lambda=cocoso_lambda,
                gra_rho=gra_rho,
                promethee_pref_func=promethee_pref_func,
                promethee_q=promethee_q,
                promethee_p=promethee_p,
                promethee_s=promethee_s,
                fuzzy_spread=fuzzy_spread,
            )
            try:
                rho, _ = spearmanr(base_rank.sort_index(), res.set_index("Alternatif")["Sıra"].sort_index())
                rho = rho if np.isfinite(rho) else 1.0
            except (ValueError, TypeError):
                rho = 1.0
            scenario_rows.append(
                {
                    "Kriter": crit,
                    "AğırlıkDeğişimi": f"{delta:+.0%}",
                    "YeniAğırlık": new_w[crit_idx],
                    "BirinciAlternatif": res.iloc[0]["Alternatif"],
                    "SpearmanRho": rho,
                }
            )

    mc_rows = []
    top_counter: Dict[str, int] = {}
    mean_rank_tracker: Dict[str, List[int]] = {str(idx): [] for idx in data.index}
    _safe_sigma = min(float(sigma), 1.0)
    for _ in range(max(iterations, 1)):
        noise = rng.normal(0.0, _safe_sigma, size=len(criteria))
        noise = np.clip(noise, -5.0, 5.0)
        draw = weight_vec * np.exp(noise)
        draw = _normalize_weights(draw)
        wd = dict(zip(criteria, draw))
        res, _ = rank_alternatives(
            data,
            criteria,
            criteria_types,
            wd,
            ranking_method,
            vikor_v=vikor_v,
            waspas_lambda=waspas_lambda,
            codas_tau=codas_tau,
            cocoso_lambda=cocoso_lambda,
            gra_rho=gra_rho,
            promethee_pref_func=promethee_pref_func,
            promethee_q=promethee_q,
            promethee_p=promethee_p,
            promethee_s=promethee_s,
            fuzzy_spread=fuzzy_spread,
        )
        top_alt = res.iloc[0]["Alternatif"]
        top_counter[top_alt] = top_counter.get(top_alt, 0) + 1
        for _, row in res.iterrows():
            mean_rank_tracker[str(row["Alternatif"])].append(int(row["Sıra"]))
        mc_rows.append({"BirinciAlternatif": top_alt})

    stability_table = pd.DataFrame(
        {
            "Alternatif": list(mean_rank_tracker.keys()),
            "OrtalamaSıra": [np.mean(v) if v else np.nan for v in mean_rank_tracker.values()],
            "BirincilikOranı": [top_counter.get(k, 0) / max(iterations, 1) for k in mean_rank_tracker.keys()],
        }
    ).sort_values(["BirincilikOranı", "OrtalamaSıra"], ascending=[False, True]).reset_index(drop=True)

    local_df = pd.DataFrame(scenario_rows)
    return {
        "base_top": base_top,
        "local_scenarios": local_df,
        "local_sensitivity": local_df.copy(),
        "monte_carlo_summary": stability_table,
        "monte_carlo_raw": pd.DataFrame(mc_rows),
        "top_stability": float(top_counter.get(base_top, 0) / max(iterations, 1)),
        "n_iterations": int(max(iterations, 1)),
        "sigma": float(sigma),
    }


def _top_weight_criterion(weights: Dict[str, float]) -> Tuple[str, float]:
    item = max(weights.items(), key=lambda x: x[1])
    return item[0], float(item[1])

def _comparison_agreement_metrics(comparison: Dict[str, Any]) -> Dict[str, float]:
    corr_df = comparison.get("spearman_matrix") if isinstance(comparison, dict) else None
    if not isinstance(corr_df, pd.DataFrame) or corr_df.empty or "Yöntem" not in corr_df.columns:
        return {"mean_rho": np.nan, "min_rho": np.nan}
    corr_only = corr_df.set_index("Yöntem")
    tri_mask = np.triu(np.ones(corr_only.shape, dtype=bool), k=1)
    upper_vals = corr_only.where(tri_mask).stack()
    upper_vals = pd.to_numeric(upper_vals, errors="coerce").dropna()
    if upper_vals.empty:
        return {"mean_rho": np.nan, "min_rho": np.nan}
    return {"mean_rho": float(upper_vals.mean()), "min_rho": float(upper_vals.min())}

def _decision_confidence_summary(
    ranking_method: Optional[str],
    ranking_table: Optional[pd.DataFrame],
    comparison: Dict[str, Any],
    sensitivity: Optional[Dict[str, Any]],
    ranking_details: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    verdict = "medium"
    notes: List[str] = []
    actions: List[str] = []

    agreement_metrics = _comparison_agreement_metrics(comparison or {})
    mean_rho = agreement_metrics["mean_rho"]
    stability = float((sensitivity or {}).get("top_stability", np.nan)) if sensitivity else np.nan

    if pd.notna(mean_rho):
        if mean_rho >= 0.85:
            notes.append(f"Yöntemler arası uyum yüksek (ρ = {mean_rho:.3f}).")
        elif mean_rho >= 0.70:
            notes.append(f"Yöntemler arası uyum orta düzeyde (ρ = {mean_rho:.3f}).")
            actions.append("Yöntem seçimi gerekçesi raporda açıkça savunulmalıdır.")
        else:
            verdict = "low"
            notes.append(f"Yöntemler arası uyum düşük (ρ = {mean_rho:.3f}).")
            actions.append("Tek bir yönteme dayalı kesin öneri verilmemeli; uzlaşı odaklı yöntemlerle çapraz kontrol yapılmalıdır.")

    if pd.notna(stability):
        if stability >= 0.80:
            notes.append(f"Lider alternatif Monte Carlo altında güçlü biçimde korunuyor (%{stability*100:.1f}).")
        elif stability >= 0.60:
            verdict = "medium" if verdict == "high" else verdict
            notes.append(f"Lider alternatif orta düzeyde kararlı (%{stability*100:.1f}).")
            actions.append("Sonuç duyarlılık analiziyle birlikte sunulmalıdır.")
        else:
            verdict = "low"
            notes.append(f"Lider alternatifin kararlılığı zayıf (%{stability*100:.1f}).")
            actions.append("Kesin lider ilan etmek yerine ilk iki-üç alternatifi birlikte değerlendirin.")

    method_name = str(ranking_method or "")
    if "VIKOR" in method_name.upper() and isinstance(ranking_details, dict):
        conditions = ranking_details.get("compromise_conditions") or {}
        adv = bool(conditions.get("acceptable_advantage", False))
        stab = bool(conditions.get("acceptable_stability", False))
        if adv and stab:
            notes.append("VIKOR uzlaşı koşullarının tamamı sağlandı.")
            if verdict != "low" and pd.isna(stability) and pd.isna(mean_rho):
                verdict = "high"
        else:
            verdict = "low"
            notes.append("VIKOR uzlaşı koşulları tam sağlanmadı.")
            actions.append("Uzlaşı koşulları sağlanmadığı için lider sonuç ihtiyat kaydıyla yorumlanmalıdır.")

    if isinstance(ranking_table, pd.DataFrame) and not ranking_table.empty:
        top_score = float(pd.to_numeric(ranking_table.iloc[0]["Skor"], errors="coerce"))
        second_score = float(pd.to_numeric(ranking_table.iloc[1]["Skor"], errors="coerce")) if len(ranking_table) > 1 else top_score
        last_score = float(pd.to_numeric(ranking_table.iloc[-1]["Skor"], errors="coerce"))
        lower_is_better = _is_lower_better_method(ranking_method)
        gap = (second_score - top_score) if lower_is_better else (top_score - second_score)
        score_range = (last_score - top_score) if lower_is_better else (top_score - last_score)
        rel_gap = gap / (score_range + EPS)
        if rel_gap < 0.08 and len(ranking_table) > 1:
            if verdict != "low":
                verdict = "medium"
            notes.append(f"İlk iki alternatif arasındaki ayrışma sınırlı (göreli fark = {rel_gap:.3f}).")
            actions.append("Lider ile ikinci alternatif arasındaki fark küçük olduğu için ek senaryo analizi önerilir.")
        elif rel_gap >= 0.15 and verdict == "medium" and (pd.notna(stability) and stability >= 0.8 or pd.notna(mean_rho) and mean_rho >= 0.85):
            verdict = "high"

    positive_signals = int(pd.notna(mean_rho) and mean_rho >= 0.85) + int(pd.notna(stability) and stability >= 0.80)
    if verdict == "medium" and positive_signals >= 2:
        verdict = "high"

    if verdict == "medium" and not notes:
        notes.append("Karar sinyali kullanılabilir düzeyde, ancak tek başına kesinlik iddiası taşımaz.")
    if verdict == "high" and not actions:
        actions.append("Sonuç yayın ve raporlamada ana öneri olarak sunulabilir.")
    elif verdict == "medium" and not actions:
        actions.append("Sonuç orta güven düzeyinde sunulmalı ve sınırları açıklanmalıdır.")
    elif verdict == "low" and not actions:
        actions.append("Sonuç keşifsel nitelikte sunulmalı; ek doğrulama yapılmadan kesin öneri verilmemelidir.")

    return {
        "verdict": verdict,
        "mean_rho": mean_rho,
        "stability": stability,
        "notes": notes,
        "actions": actions,
    }


def _base_method_name(method: Optional[str]) -> Optional[str]:
    if not method:
        return None
    return method.replace("Fuzzy ", "")


def _report_references(weight_method: str, ranking_method: Optional[str]) -> List[str]:
    refs = {REFERENCE_LIBRARY["GENERAL_MCDM"], MANDATORY_MCDM_REFERENCE}
    refs.add(REFERENCE_LIBRARY.get(weight_method, REFERENCE_LIBRARY["GENERAL_MCDM"]))

    weight_pubs = RENCBER_PUBLICATIONS_BY_METHOD.get(weight_method, [])
    if weight_pubs:
        refs.update(weight_pubs)

    if ranking_method:
        refs.add(REFERENCE_LIBRARY.get(ranking_method, REFERENCE_LIBRARY["GENERAL_MCDM"]))
        base = _base_method_name(ranking_method)
        if base and base in REFERENCE_LIBRARY:
            refs.add(REFERENCE_LIBRARY[base])

        ranking_pubs = list(RENCBER_PUBLICATIONS_BY_METHOD.get(ranking_method, []))
        if base:
            ranking_pubs.extend(RENCBER_PUBLICATIONS_BY_METHOD.get(base, []))
        if ranking_pubs:
            refs.update(ranking_pubs)

    return sorted(refs)


def _section_amac(
    stats_df: pd.DataFrame,
    criteria: Sequence[str],
    criteria_types: Dict[str, str],
    ranking_method: Optional[str],
) -> str:
    benefit = [c for c in criteria if criteria_types.get(c, "max") == "max"]
    cost = [c for c in criteria if criteria_types.get(c, "max") == "min"]
    sentence = (
        f"Bu çalışmanın amacı, {len(stats_df)} alternatif ve {len(criteria)} kriterden oluşan karar matrisinde "
        f"alternatiflerin performansını nesnel biçimde değerlendirmek ve kriterler arası bilgi yapısını veri temelli "
        f"olarak çözümlemektir. Analizde fayda yönlü kriterler ({', '.join(benefit) if benefit else 'yok'}) ile "
        f"maliyet yönlü kriterler ({', '.join(cost) if cost else 'yok'}) birlikte ele alınmıştır."
    )
    if ranking_method:
        sentence += (
            f" Nihai amaç, objektif ağırlıklandırma ile elde edilen kriter önemlerini {ranking_method} yöntemi "
            f"yardımıyla bütünleştirip tutarlı bir alternatif sıralaması üretmektir."
        )
    else:
        sentence += " Nihai amaç, karar matrisinden öznel yargı kullanmadan savunulabilir bir kriter ağırlık vektörü türetmektir."
    return sentence


def _section_felsefe(weight_method: str, ranking_method: Optional[str], fuzzy_spread: float, is_fuzzy: bool) -> str:
    weight_ph = METHOD_PHILOSOPHY.get(weight_method, {})
    text = (
        "Bu çalışmanın felsefesi, kriter öneminin uzman anketlerine değil doğrudan verinin kendi ayırt edicilik yapısına "
        "dayandırılmasıdır. Böylece sonuçlar karar vericinin sezgisel tercihlerinden ziyade karar matrisinin içsel bilgi "
        "içeriğine yaslanmaktadır. "
        + weight_ph.get("academic", "")
    )
    if ranking_method:
        rank_ph = METHOD_PHILOSOPHY.get(ranking_method, {})
        text += " " + rank_ph.get("academic", "")
    if is_fuzzy:
        text += (
            f" Belirsizlik katmanı, öznel dilsel puanlar yerine nicel verinin ±%{fuzzy_spread*100:.1f} oranlı üçgensel "
            "bulanık çevresi üzerinden modellenmiştir. Böylece kriter ağırlıkları nesnel kalırken alternatif performansları "
            "ölçüm oynaklığı ve veri belirsizliği bakımından daha ihtiyatlı değerlendirilmiştir."
        )
    return text


def _section_metodoloji(
    validation: Dict[str, Any],
    weight_method: str,
    ranking_method: Optional[str],
    pca_info: Dict[str, Any],
    is_fuzzy: bool,
    fuzzy_spread: float,
) -> str:
    text = (
        f"Metodoloji dört aşamada yürütülmüştür: (i) karar matrisinin sayısal uygunluğu ve ayırt ediciliği sınanmış, "
        f"(ii) {weight_method} yöntemi ile objektif kriter ağırlıkları hesaplanmış, "
    )
    if ranking_method:
        text += f"(iii) {ranking_method} yöntemi ile alternatif skorları elde edilmiş, "
    else:
        text += "(iii) alternatif sıralaması talep edilmediği için ağırlık vektörünün yorumuna odaklanılmış, "
    text += (
        "(iv) korelasyon, temel bileşen yapısı ve duyarlılık analizi ile bulguların dayanıklılığı incelenmiştir. "
        f"Veri denetiminde {len(validation.get('high_corr_pairs', []))} adet yüksek korelasyonlu kriter çifti ve "
        f"{len(validation.get('constant_criteria', []))} adet sabit kriter raporlanmıştır. "
        f"PCA incelemesinde Kaiser eşiğini aşan bileşenler {', '.join(pca_info.get('selected_components', []))} olarak belirlenmiştir."
    )
    if is_fuzzy:
        text += (
            f" Fuzzy katmanda üçgensel bulanık sayılar kullanılmış; her gözlem merkez değer etrafında ±%{fuzzy_spread*100:.1f} "
            "genişliğinde modellenmiş ve gerekli aşamalarda centroid durulaştırması uygulanmıştır."
        )
    return text


def _section_bulgular(
    weight_table: pd.DataFrame,
    ranking_table: Optional[pd.DataFrame],
    comparison: Dict[str, Any],
    sensitivity: Optional[Dict[str, Any]],
    weight_method: str,
    ranking_method: Optional[str],
    ranking_details: Optional[Dict[str, Any]] = None,
) -> str:
    top_criterion = str(weight_table.iloc[0]["Kriter"])
    top_weight = float(weight_table.iloc[0]["Ağırlık"])
    text = (
        f"Bulgular, {weight_method} yöntemi altında en yüksek ağırlığın {top_criterion} kriterinde yoğunlaştığını "
        f"(w = {top_weight:.4f}) göstermektedir. Bu bulgu, karar matrisindeki ayırt ediciliğin ve bilgi içeriğinin en fazla "
        f"bu kriter üzerinden aktığını düşündürmektedir."
    )
    if ranking_table is not None and not ranking_table.empty and ranking_method:
        top_alt = str(ranking_table.iloc[0]['Alternatif'])
        top_score = float(ranking_table.iloc[0]['Skor'])
        second_alt = str(ranking_table.iloc[1]['Alternatif']) if len(ranking_table) > 1 else top_alt
        second_score = float(ranking_table.iloc[1]['Skor']) if len(ranking_table) > 1 else top_score
        lower_is_better = _is_lower_better_method(ranking_method)
        gap = (second_score - top_score) if lower_is_better else (top_score - second_score)
        text += (
            f" {ranking_method} sonuçlarına göre birinci sırada {top_alt} yer almış ve {top_score:.4f} skoruna ulaşmıştır. "
            f"İkinci sıradaki {second_alt} ile fark {gap:.4f} düzeyindedir; bu farkın büyüklüğü kararın ne ölçüde keskin "
            f"veya hassas olduğunu göstermektedir."
        )
    if comparison:
        agreement_metrics = _comparison_agreement_metrics(comparison)
        if pd.notna(agreement_metrics["mean_rho"]):
            text += (
                f" Çoklu yöntem karşılaştırmasında yöntemler arası ortalama Spearman uyumu {agreement_metrics['mean_rho']:.3f} "
                "olarak bulunmuştur; bu değer farklı yöntem aileleri arasında yapısal tutarlılık düzeyini göstermektedir."
            )
    if sensitivity:
        stability = sensitivity.get("top_stability")
        if stability is not None:
            text += (
                f" Monte Carlo temelli ağırlık bozulmalarında birinci alternatifin korunma oranı {stability:.2%} "
                "olarak hesaplanmıştır. Bu oran, bulguların ağırlık sapmalarına karşı dayanıklılığı hakkında doğrudan kanıt sunmaktadır."
            )
    confidence = _decision_confidence_summary(ranking_method, ranking_table, comparison, sensitivity, ranking_details)
    verdict_map = {
        "high": "yüksek güven düzeyinde",
        "medium": "orta güven düzeyinde",
        "low": "düşük güven düzeyinde",
    }
    if ranking_table is not None and not ranking_table.empty and ranking_method:
        text += (
            f" Bütüncül karar sinyali {verdict_map.get(confidence['verdict'], 'orta güven düzeyinde')} "
            "değerlendirilmektedir."
        )
        if confidence["notes"]:
            text += " " + " ".join(confidence["notes"][:2])
        if confidence["actions"]:
            text += " " + " ".join(confidence["actions"][:2])
    return text


def build_report_sections(
    *,
    validation: Dict[str, Any],
    stats_df: pd.DataFrame,
    criteria: Sequence[str],
    criteria_types: Dict[str, str],
    weight_method: str,
    ranking_method: Optional[str],
    weight_table: pd.DataFrame,
    ranking_table: Optional[pd.DataFrame],
    ranking_details: Optional[Dict[str, Any]],
    comparison: Dict[str, Any],
    sensitivity: Optional[Dict[str, Any]],
    pca_info: Dict[str, Any],
    fuzzy_spread: float = 0.10,
) -> Dict[str, Any]:
    is_fuzzy = bool(ranking_method and ranking_method.startswith("Fuzzy"))
    sections = {
        "Çalışmanın Amacı": _section_amac(stats_df, criteria, criteria_types, ranking_method),
        "Çalışmanın Felsefesi": _section_felsefe(weight_method, ranking_method, fuzzy_spread, is_fuzzy),
        "Metodoloji": _section_metodoloji(validation, weight_method, ranking_method, pca_info, is_fuzzy, fuzzy_spread),
        "Bulgular": _section_bulgular(weight_table, ranking_table, comparison, sensitivity, weight_method, ranking_method, ranking_details),
        "Kaynakça": _report_references(weight_method, ranking_method),
    }
    return sections


def run_full_analysis(data: pd.DataFrame, config: AnalysisConfig) -> Dict[str, Any]:
    criteria = list(config.criteria)
    criteria_types = {c: config.criteria_types.get(c, "max") for c in criteria}
    df = _as_numeric_df(data, criteria)
    validation = validate_problem(df, criteria, criteria_types)
    if validation["errors"]:
        raise ValueError("\n".join(validation["errors"]))

    stats_df = descriptive_statistics(df, criteria)
    corr_df = df.corr(numeric_only=True).reset_index().rename(columns={"index": "Kriter"})
    pca_info = pca_diagnostics(df, criteria, criteria_types)
    if config.weight_mode == "equal":
        ew = np.ones(len(criteria), dtype=float) / max(len(criteria), 1)
        weights = dict(zip(criteria, ew))
        _equal_table = pd.DataFrame({"Kriter": criteria, "Ağırlık": ew}).sort_values("Ağırlık", ascending=False).reset_index(drop=True)
        _equal_table.insert(1, "ÖnemSırası", np.arange(1, len(_equal_table) + 1, dtype=int))
        weight_details = {
            "weight_table": _equal_table,
            "mode": "equal",
        }
    elif config.weight_mode == "manual":
        user_w = config.manual_weights or {}
        wvec = _normalize_weights([float(user_w.get(c, 0.0)) for c in criteria])
        weights = dict(zip(criteria, wvec))
        _manual_table = pd.DataFrame({"Kriter": criteria, "Ağırlık": wvec}).sort_values("Ağırlık", ascending=False).reset_index(drop=True)
        _manual_table.insert(1, "ÖnemSırası", np.arange(1, len(_manual_table) + 1, dtype=int))
        weight_details = {
            "weight_table": _manual_table,
            "manual_input": pd.DataFrame({"Kriter": criteria, "GirilenDeğer": [float(user_w.get(c, 0.0)) for c in criteria]}),
            "mode": "manual",
        }
    else:
        weights, weight_details = compute_objective_weights(
            df,
            criteria,
            criteria_types,
            config.weight_method,
            fuzzy_spread=config.fuzzy_spread,
        )
    contribution_df = _contribution_table(df, criteria_types, weights)

    ranking_table: Optional[pd.DataFrame] = None
    ranking_details: Dict[str, Any] = {}
    if config.ranking_method:
        ranking_table, ranking_details = rank_alternatives(
            df,
            criteria,
            criteria_types,
            weights,
            config.ranking_method,
            vikor_v=config.vikor_v,
            waspas_lambda=config.waspas_lambda,
            codas_tau=config.codas_tau,
            cocoso_lambda=config.cocoso_lambda,
            gra_rho=config.gra_rho,
            promethee_pref_func=config.promethee_pref_func,
            promethee_q=config.promethee_q,
            promethee_p=config.promethee_p,
            promethee_s=config.promethee_s,
            fuzzy_spread=config.fuzzy_spread,
        )

    compare = compare_methods(
        df,
        criteria,
        criteria_types,
        weights,
        config.compare_methods or [],
        vikor_v=config.vikor_v,
        waspas_lambda=config.waspas_lambda,
        codas_tau=config.codas_tau,
        cocoso_lambda=config.cocoso_lambda,
        gra_rho=config.gra_rho,
        promethee_pref_func=config.promethee_pref_func,
        promethee_q=config.promethee_q,
        promethee_p=config.promethee_p,
        promethee_s=config.promethee_s,
        fuzzy_spread=config.fuzzy_spread,
        base_method=config.ranking_method,
        base_table=ranking_table,
        base_details=ranking_details,
    ) if config.compare_methods else {}

    sensitivity = None
    if config.ranking_method and bool(config.run_heavy_robustness) and int(config.sensitivity_iterations) > 0:
        sensitivity = sensitivity_analysis(
            df,
            criteria,
            criteria_types,
            weights,
            config.ranking_method,
            vikor_v=config.vikor_v,
            waspas_lambda=config.waspas_lambda,
            codas_tau=config.codas_tau,
            cocoso_lambda=config.cocoso_lambda,
            gra_rho=config.gra_rho,
            promethee_pref_func=config.promethee_pref_func,
            promethee_q=config.promethee_q,
            promethee_p=config.promethee_p,
            promethee_s=config.promethee_s,
            fuzzy_spread=config.fuzzy_spread,
            iterations=config.sensitivity_iterations,
            sigma=config.sensitivity_sigma,
        )

    report_sections = build_report_sections(
        validation=validation,
        stats_df=stats_df,
        criteria=criteria,
        criteria_types=criteria_types,
        weight_method=config.weight_method,
        ranking_method=config.ranking_method,
        weight_table=weight_details["weight_table"],
        ranking_table=ranking_table,
        ranking_details=ranking_details,
        comparison=compare,
        sensitivity=sensitivity,
        pca_info=pca_info,
        fuzzy_spread=config.fuzzy_spread,
    )

    decision_confidence = _decision_confidence_summary(
        config.ranking_method,
        ranking_table,
        compare,
        sensitivity,
        ranking_details,
    ) if config.ranking_method else {}

    result = {
        "validation": validation,
        "stats": stats_df,
        "correlation": corr_df,
        "correlation_matrix": df.corr(numeric_only=True),
        "pca": pca_info,
        "weights": {
            "method": config.weight_method,
            "values": weights,
            "details": weight_details,
            "table": weight_details["weight_table"],
        },
        "ranking": {
            "method": config.ranking_method,
            "table": ranking_table,
            "details": ranking_details,
        },
        "comparison": compare,
        "sensitivity": sensitivity,
        "decision_confidence": decision_confidence,
        "contribution_table": contribution_df,
        "report_sections": report_sections,
        "method_philosophy": {
            "weight": METHOD_PHILOSOPHY.get(config.weight_method, {}),
            "ranking": METHOD_PHILOSOPHY.get(config.ranking_method, {}) if config.ranking_method else {},
        },
    }
    return result


def run_scenario_analysis(
    data: pd.DataFrame,
    criteria: Sequence[str],
    criteria_types: Dict[str, str],
    scenarios: Dict[str, Dict[str, float]],
    ranking_method: str,
    *,
    vikor_v: float = 0.5,
    waspas_lambda: float = 0.5,
    codas_tau: float = 0.02,
    cocoso_lambda: float = 0.5,
    gra_rho: float = 0.5,
    promethee_pref_func: str = "linear",
    promethee_q: float = 0.05,
    promethee_p: float = 0.30,
    promethee_s: float = 0.20,
    fuzzy_spread: float = 0.10,
) -> Dict[str, Any]:
    """Run the same ranking method with multiple weight scenarios.

    Parameters
    ----------
    scenarios : dict
        ``{"Scenario Name": {"C1": 0.3, "C2": 0.7, ...}, ...}``
        Each value is a weight dict (will be auto-normalized).

    Returns
    -------
    dict with keys:
        - ``rank_comparison``: DataFrame (alternatives × scenarios → ranks)
        - ``score_comparison``: DataFrame (alternatives × scenarios → scores)
        - ``leader_summary``: DataFrame (scenario, leader, score)
        - ``agreement_matrix``: DataFrame (Spearman ρ between scenarios)
        - ``all_same_leader``: bool
        - ``per_scenario``: dict of full ranking tables
    """
    from scipy.stats import spearmanr as _sp

    per_scenario: Dict[str, pd.DataFrame] = {}
    rank_kwargs = dict(
        vikor_v=vikor_v, waspas_lambda=waspas_lambda, codas_tau=codas_tau,
        cocoso_lambda=cocoso_lambda, gra_rho=gra_rho,
        promethee_pref_func=promethee_pref_func, promethee_q=promethee_q,
        promethee_p=promethee_p, promethee_s=promethee_s, fuzzy_spread=fuzzy_spread,
    )

    for sc_name, sc_weights in scenarios.items():
        rt, _ = rank_alternatives(data, criteria, criteria_types, sc_weights, ranking_method, **rank_kwargs)
        per_scenario[sc_name] = rt

    # Build comparison tables
    alt_col = "Alternatif"
    r_col = "Sıra"
    s_col = "Skor"
    alts = list(per_scenario[next(iter(per_scenario))][alt_col])

    rank_rows = []
    score_rows = []
    for alt in alts:
        r_row: Dict[str, Any] = {alt_col: alt}
        s_row: Dict[str, Any] = {alt_col: alt}
        for sc_name, rt in per_scenario.items():
            match = rt[rt[alt_col] == alt]
            r_row[sc_name] = int(match[r_col].iloc[0]) if not match.empty else None
            s_row[sc_name] = float(match[s_col].iloc[0]) if not match.empty else None
        rank_rows.append(r_row)
        score_rows.append(s_row)

    rank_df = pd.DataFrame(rank_rows)
    score_df = pd.DataFrame(score_rows)

    # Leader summary
    sc_names = list(scenarios.keys())
    leaders = []
    for sc_name in sc_names:
        rt = per_scenario[sc_name]
        top = rt.sort_values(r_col).iloc[0]
        leaders.append({"Senaryo": sc_name, "Lider": str(top[alt_col]), "Skor": float(top[s_col])})
    leader_df = pd.DataFrame(leaders)
    all_same = len(set(l["Lider"] for l in leaders)) == 1

    # Spearman agreement between scenarios
    corr = pd.DataFrame(index=sc_names, columns=sc_names, dtype=float)
    for s1 in sc_names:
        for s2 in sc_names:
            try:
                rho, _ = _sp(rank_df[s1], rank_df[s2])
                corr.loc[s1, s2] = rho if np.isfinite(rho) else 1.0 if s1 == s2 else np.nan
            except (ValueError, TypeError):
                corr.loc[s1, s2] = 1.0 if s1 == s2 else np.nan

    return {
        "rank_comparison": rank_df,
        "score_comparison": score_df,
        "leader_summary": leader_df,
        "agreement_matrix": corr,
        "all_same_leader": all_same,
        "per_scenario": per_scenario,
    }


def available_methods() -> Dict[str, List[str]]:
    return {
        "objective_weights": OBJECTIVE_WEIGHT_METHODS,
        "classical_mcdm": CLASSICAL_MCDM_METHODS,
        "fuzzy_mcdm": FUZZY_MCDM_METHODS,
    }


# ─────────────────────────────────────────────────────────────────────────────
#  DSS KATMANI: Veri Tanı Asistanı + 3 Kademeli Yorum Motorları
# ─────────────────────────────────────────────────────────────────────────────

def generate_data_diagnostics(
    data: pd.DataFrame,
    criteria: Sequence[str],
    criteria_types: Dict[str, str],
) -> Dict[str, Any]:
    """Veri yüklendikten hemen sonra röntgen çeker; yöntem önerisi üretir."""
    df = _as_numeric_df(data, criteria)
    n_alt, n_crit = df.shape

    stats = descriptive_statistics(df, criteria)
    cv_vals = {str(row["Kriter"]): float(row["cv"]) for _, row in stats.iterrows()}
    mean_cv = float(np.nanmean(list(cv_vals.values()))) if cv_vals else 0.0
    max_cv = float(np.nanmax(list(cv_vals.values()))) if cv_vals else 0.0
    cv_std = float(np.nanstd(list(cv_vals.values()))) if cv_vals else 0.0

    skew_vals = df.skew(numeric_only=True).replace([np.inf, -np.inf], np.nan).fillna(0.0)
    mean_abs_skew = float(skew_vals.abs().mean()) if not skew_vals.empty else 0.0

    q1 = df.quantile(0.25)
    q3 = df.quantile(0.75)
    iqr = (q3 - q1).replace(0, np.nan)
    outlier_mask = ((df.lt(q1 - 1.5 * iqr)) | (df.gt(q3 + 1.5 * iqr))).fillna(False)
    outlier_ratio = float(outlier_mask.to_numpy(dtype=float).mean()) if outlier_mask.size else 0.0

    corr = df.corr(numeric_only=True)
    high_corr_pairs: List[Tuple[str, str, float]] = []
    for i, c1 in enumerate(df.columns):
        for c2 in df.columns[i + 1 :]:
            val = corr.loc[c1, c2]
            if pd.notna(val) and abs(val) >= 0.70:
                high_corr_pairs.append((c1, c2, float(val)))
    max_corr = max((abs(v) for _, _, v in high_corr_pairs), default=0.0)

    constant = [c for c in criteria if df[c].nunique() <= 1]
    non_positive = [c for c in criteria if float(df[c].min()) <= 0.0]
    benefit = [c for c in criteria if criteria_types.get(c, "max") == "max"]
    cost = [c for c in criteria if criteria_types.get(c, "max") == "min"]
    cost_ratio = len(cost) / max(n_crit, 1)
    alt_crit_ratio = n_alt / max(n_crit, 1)
    risk_flags: List[str] = []

    recommendations: List[Dict[str, str]] = []
    suggested_weight: Optional[str] = None
    suggested_ranking: Optional[str] = None
    suggested_ranking_methods: List[str] = []

    if constant:
        recommendations.append({
            "level": "error",
            "icon": "🚫",
            "text": f"Sabit kriter(ler) tespit edildi: {', '.join(constant)}. Bu kriterler analitik bilgi taşımaz.",
            "action": "Bu kriterleri kriter listesinden çıkarmanız kesinlikle önerilir.",
            "text_en": f"Constant criteria detected: {', '.join(constant)}. These criteria do not carry analytical information.",
            "action_en": "You should remove these criteria from the active analysis set.",
        })
        risk_flags.append("constant_criteria")

    if non_positive:
        recommendations.append({
            "level": "warning",
            "icon": "➖",
            "text": f"Sıfır veya negatif değer içeren kriterler var: {', '.join(non_positive)}.",
            "action": "Oran/çarpım mantıklı yöntemleri yorumlarken dikkatli olun; gerekiyorsa veri dönüşümü veya sağlamlık kontrolü uygulayın.",
            "text_en": f"Some criteria contain zero or negative values: {', '.join(non_positive)}.",
            "action_en": "Interpret ratio/product-based methods carefully; apply data transformation or robustness checks when needed.",
        })
        risk_flags.append("non_positive_values")

    if outlier_ratio >= 0.12:
        recommendations.append({
            "level": "warning",
            "icon": "📍",
            "text": f"Aykırı gözlem oranı %{outlier_ratio*100:.1f} ile yüksektir.",
            "action": "Ön işlemde aykırı değer temizleme veya winsorization uygulayın; ardından sonuçları EDAS/CODAS ve Monte Carlo ile çapraz kontrol edin.",
            "text_en": f"The outlier ratio is high at {outlier_ratio*100:.1f}%.",
            "action_en": "Apply outlier treatment or winsorization in preprocessing, then cross-check results with EDAS/CODAS and Monte Carlo robustness.",
        })
        risk_flags.append("high_outliers")
    elif outlier_ratio >= 0.06:
        recommendations.append({
            "level": "info",
            "icon": "📍",
            "text": f"Sınırlı ama izlenmesi gereken aykırı değer oranı tespit edildi (%{outlier_ratio*100:.1f}).",
            "action": "İlk üç alternatifi aykırı değer temizliği öncesi ve sonrası karşılaştırın.",
            "text_en": f"A moderate but notable outlier ratio was detected ({outlier_ratio*100:.1f}%).",
            "action_en": "Compare the top three alternatives before and after outlier treatment.",
        })
        risk_flags.append("moderate_outliers")

    if mean_abs_skew >= 1.0:
        recommendations.append({
            "level": "info",
            "icon": "📈",
            "text": f"Ortalama mutlak çarpıklık {mean_abs_skew:.2f}; dağılımlar belirgin biçimde asimetrik.",
            "action": "Sonuçları yalnızca tek bir mesafe yöntemiyle bırakmayın; EDAS veya CODAS ile karşılaştırma sekmesini açık tutun.",
            "text_en": f"Mean absolute skewness is {mean_abs_skew:.2f}; the distributions are materially asymmetric.",
            "action_en": "Do not rely on a single distance method; keep EDAS or CODAS in the comparison set.",
        })
        risk_flags.append("high_skewness")

    if alt_crit_ratio < 2.0:
        recommendations.append({
            "level": "warning",
            "icon": "🧪",
            "text": f"Alternatif/kriter oranı {alt_crit_ratio:.2f}; problem boyutu sıkışık.",
            "action": "Bulguları keşifsel düzeyde yorumlayın, güçlü korelasyonlu kriterleri azaltın ve mümkünse alternatif sayısını artırın.",
            "text_en": f"The alternative-to-criterion ratio is {alt_crit_ratio:.2f}; the problem is data-tight.",
            "action_en": "Interpret findings as exploratory, reduce highly correlated criteria, and increase the number of alternatives when possible.",
        })
        risk_flags.append("tight_sample")

    if high_corr_pairs:
        lead_pairs = ", ".join(f"{c1}-{c2} ({val:.2f})" for c1, c2, val in high_corr_pairs[:3])
        recommendations.append({
            "level": "info",
            "icon": "🔗",
            "text": f"Yüksek korelasyonlu kriter çiftleri mevcut: {lead_pairs}.",
            "action": "CRITIC/IDOCRIW ile ağırlıklandırmayı koruyun ve gereksiz tekrar yaratan kriterleri raporda gerekçelendirin.",
            "text_en": f"Highly correlated criterion pairs are present: {lead_pairs}.",
            "action_en": "Prefer CRITIC/IDOCRIW weighting and justify potentially redundant criteria in the report.",
        })
        risk_flags.append("high_correlation")

    # ── AĞIRLIKLANDI RMA KARAR AĞACI ──────────────────────────────────────────
    # Öncelik sırası: korelasyon yapısı → varyasyon düzeyi → veri sıkışıklığı
    if max_corr >= 0.85:
        suggested_weight = "CRITIC"
        weight_reason_tr = (
            f"Karar ağacı çıktısı — Yüksek kriter korelasyonu (maks. r = {max_corr:.2f}): "
            "Kriterler arasında güçlü doğrusal ilişki varken bilgi tekrarını en doğrudan cezalandıran yöntem CRITIC'tir. "
            "CRITIC, kriter içi varyansı hem kriter-arası çatışma miktarıyla çarpar hem de korelasyonu varyans bileşenine göre ağırlıklandırır; "
            "bu yapıda IDOCRIW ve PCA doğrulama amaçlı ikincil öneri olarak kullanılabilir."
        )
        weight_reason_en = (
            f"Decision-tree output — High criterion correlation (max r = {max_corr:.2f}): "
            "When strong linear relationships exist between criteria, CRITIC is the most direct approach for penalising information redundancy. "
            "CRITIC multiplies within-criterion variance by inter-criterion conflict while weighting correlations against the variance component; "
            "IDOCRIW and PCA serve as secondary verification options in this setting."
        )
    elif max_corr >= 0.70 and alt_crit_ratio >= 2.5:
        suggested_weight = "IDOCRIW"
        weight_reason_tr = (
            f"Karar ağacı çıktısı — Orta-yüksek korelasyon (maks. r = {max_corr:.2f}), yeterli alt/krit oranı ({alt_crit_ratio:.1f}): "
            "IDOCRIW, Entropi tabanlı bilgi içeriğini CILOS'un göreli etki kaybı yapısıyla birleştiren hibrit bir yöntemdir. "
            "Hem bilgi tekrarını hem de her kriterin diğerleri üzerindeki baskı gücünü aynı anda ölçer; "
            "bu profilde sadece korelasyona odaklanan CRITIC'ten daha dengeli ve savunulabilir bir ağırlık üretir."
        )
        weight_reason_en = (
            f"Decision-tree output — Moderate-to-high correlation (max r = {max_corr:.2f}), adequate alt/crit ratio ({alt_crit_ratio:.1f}): "
            "IDOCRIW combines Entropy-based information content with CILOS criterion-impact-loss structure in a hybrid framework. "
            "It simultaneously measures information redundancy and each criterion's influence on the others; "
            "in this profile it produces more balanced and defensible weights than CRITIC alone."
        )
    elif max_corr >= 0.70 and alt_crit_ratio < 2.5:
        suggested_weight = "CILOS"
        weight_reason_tr = (
            f"Karar ağacı çıktısı — Orta-yüksek korelasyon (maks. r = {max_corr:.2f}), sıkışık veri yapısı (alt/krit = {alt_crit_ratio:.1f}): "
            "CILOS, kriter etki kaybı (Criterion Impact LOSs) ilkesine dayanır: her kriter çıkarıldığında en iyi alternatifin diğer kriterler üzerindeki kaybı hesaplanır. "
            "Alternatif sayısı kritere oranla sınırlı olduğunda korelasyon matrisine duyarlı yöntemler kararsız sonuç üretebilir; "
            "CILOS bu koşulda daha istikrarlı ve yorumlanabilir ağırlıklar verir."
        )
        weight_reason_en = (
            f"Decision-tree output — Moderate-to-high correlation (max r = {max_corr:.2f}), tight data structure (alt/crit = {alt_crit_ratio:.1f}): "
            "CILOS is based on the Criterion Impact LOSs principle: it computes the loss to the best alternative on remaining criteria when each criterion is removed. "
            "When the alternative count is low relative to criteria, correlation-sensitive methods may produce unstable results; "
            "CILOS delivers more stable and interpretable weights under this condition."
        )
    elif mean_cv > 0.35:
        suggested_weight = "Entropy"
        weight_reason_tr = (
            f"Karar ağacı çıktısı — Yüksek varyasyon (ort. CV = {mean_cv:.2f}), düşük korelasyon (maks. r = {max_corr:.2f}): "
            "Kriterler birbirinden güçlü biçimde ayrışıyorsa Shannon Entropi en doğal seçimdir: "
            "her kriter için değer dağılımının ne kadar 'düzensiz' olduğunu ölçer ve dağılım tekdüze ise ağırlığı sıfıra yaklaştırır. "
            "Bu profilde Standart Sapma ve LOPCOW yakın alternatif, IDOCRIW ise hibrit doğrulama seçeneğidir."
        )
        weight_reason_en = (
            f"Decision-tree output — High variation (mean CV = {mean_cv:.2f}), low correlation (max r = {max_corr:.2f}): "
            "When criteria discriminate strongly from one another, Shannon Entropy is the most natural choice: "
            "it measures how 'disordered' each criterion's value distribution is and drives weights toward zero when a distribution is uniform. "
            "Standard Deviation and LOPCOW are close alternatives in this profile; IDOCRIW serves as a hybrid verification option."
        )
    elif mean_cv >= 0.18:
        suggested_weight = "MEREC"
        weight_reason_tr = (
            f"Karar ağacı çıktısı — Dengeli-orta varyasyon (ort. CV = {mean_cv:.2f}), korelasyon sınırlı (maks. r = {max_corr:.2f}): "
            "MEREC (Method based on the Removal Effects of Criteria), her kriteri tek tek çıkararak performans matrisindeki toplam kaybı hesaplar; "
            "bu sayede kriterlerin birbirini ne ölçüde 'tamamladığını' ağırlığa dönüştürür. "
            "Orta düzey varyasyonun olduğu dengeli profillerde MEREC, aşırı baskın sinyal üretmeden güvenilir sonuç verir; "
            "LOPCOW ise yakın ikinci seçenektir."
        )
        weight_reason_en = (
            f"Decision-tree output — Balanced-moderate variation (mean CV = {mean_cv:.2f}), limited correlation (max r = {max_corr:.2f}): "
            "MEREC (Method based on the Removal Effects of Criteria) computes the total performance loss in the matrix when each criterion is removed one at a time; "
            "this translates how much criteria 'complement' each other into weights. "
            "In balanced profiles with moderate variation, MEREC produces reliable results without generating an overly dominant signal; "
            "LOPCOW is a close second option."
        )
    else:
        suggested_weight = "LOPCOW"
        weight_reason_tr = (
            f"Karar ağacı çıktısı — Düşük varyasyon (ort. CV = {mean_cv:.2f}), düşük korelasyon (maks. r = {max_corr:.2f}): "
            "Varyasyonun sınırlı olduğu homojen profillerde Entropy veya Standart Sapma küçük farklılıkları abartabilir. "
            "LOPCOW (Logarithmic Percentage Change-driven Objective Weighting), normalize edilmiş kriteri logaritmik yüzde değişimi biçiminde ifade eder; "
            "ölçek bağımsızlığı ve küçük farkları orantılı okuma özelliğiyle düşük varyasyonlu veri için en sağlıklı seçimdir."
        )
        weight_reason_en = (
            f"Decision-tree output — Low variation (mean CV = {mean_cv:.2f}), low correlation (max r = {max_corr:.2f}): "
            "In homogeneous profiles with limited variation, Entropy or Standard Deviation can exaggerate small differences. "
            "LOPCOW (Logarithmic Percentage Change-driven Objective Weighting) expresses each normalised criterion as a logarithmic percentage change; "
            "its scale-independence and proportional sensitivity to small differences make it the most appropriate choice for low-variation data."
        )

    # ── SIRALAMA KARAR AĞACI ──────────────────────────────────────────────────
    # Öncelik: alternatif sayısı → dağılım sorunları → büyük+yüksek CV →
    #          korelasyon+maliyet yapısı → sıkışık problem → geniş ölçek aralığı →
    #          homojen profil → genel denge
    if n_alt <= 6:
        suggested_ranking_methods = ["PROMETHEE", "RAFSI", "VIKOR"]
        ranking_reason_tr = (
            f"Karar ağacı çıktısı — Küçük örneklem ({n_alt} alternatif): "
            "Az sayıda alternatif varken yöntem seçimi, sıralama hassasiyeti ve sonuç yorumlanabilirliği açısından kritik hale gelir. "
            "PROMETHEE, her alternatif çiftini ayrı ayrı karşılaştırdığı için küçük örneklemlerde ince farkları en iyi görünür kılan yöntemdir. "
            "RAFSI (Ranking of Alternatives through Functional mapping of criterion sub-intervals into a Single Interval), "
            "referans noktaları üzerinden oransal bir değerlendirme yaptığı için az alternatifte tutarlı ve savunulabilir sıralar üretir. "
            "VIKOR ise çatışan kriterler altında uzlaşı çözümü bulma kabiliyetiyle küçük grupların en iyi alternatifini belirlemede güvenilir bir tamamlayıcıdır."
        )
        ranking_reason_en = (
            f"Decision-tree output — Small sample ({n_alt} alternatives): "
            "With few alternatives, method choice becomes critical for ranking sensitivity and result interpretability. "
            "PROMETHEE compares each pair of alternatives individually, making it the best method for revealing fine differences in small samples. "
            "RAFSI (Ranking of Alternatives through Functional mapping of criterion sub-intervals into a Single Interval) "
            "uses proportional reference-point evaluation and produces consistent, defensible rankings with few alternatives. "
            "VIKOR complements these by finding a compromise solution under conflicting criteria, reliably identifying the best alternative in small groups."
        )
    elif mean_abs_skew >= 1.10 or outlier_ratio >= 0.08:
        suggested_ranking_methods = ["DNMA", "EDAS", "CODAS"]
        ranking_reason_tr = (
            f"Karar ağacı çıktısı — Asimetrik/aykırı değerli dağılım (ort. çarpıklık = {mean_abs_skew:.2f}, aykırı oran = %{outlier_ratio*100:.1f}): "
            "Dağılımlar güçlü biçimde çarpıksa veya aykırı değer oranı yüksekse, ideal noktadan mutlak mesafeye dayanan yöntemler (TOPSIS) "
            "aykırı değerlere duyarlı hale gelebilir. "
            "DNMA (Dynamic Normalization-based MCDM Approach), her kriterin normalizasyon aralığını dinamik olarak belirler; "
            "bu sayede farklı ölçek ve çarpıklık yapılarına karşı en sağlam sıralamayı üretir. "
            "EDAS, ortalamadan pozitif/negatif sapmayı ölçtüğü için uç değerlere daha az duyarlıdır. "
            "CODAS ise en kötü alternatiften Öklid ve Hamming uzaklıklarını birlikte kullanarak güçlü ayırt ediciliği korur."
        )
        ranking_reason_en = (
            f"Decision-tree output — Asymmetric/outlier-heavy distributions (mean skewness = {mean_abs_skew:.2f}, outlier ratio = {outlier_ratio*100:.1f}%): "
            "When distributions are strongly skewed or the outlier ratio is high, methods that rely on absolute distance from the ideal point (TOPSIS) "
            "become sensitive to extreme values. "
            "DNMA (Dynamic Normalization-based MCDM Approach) determines each criterion's normalisation range dynamically, "
            "producing the most robust rankings against varied scales and skewness structures. "
            "EDAS measures positive/negative deviation from the average and is therefore less sensitive to extreme values. "
            "CODAS combines Euclidean and Hamming distances from the worst alternative, maintaining strong discrimination."
        )
    elif n_alt >= 18 and mean_cv >= 0.20:
        suggested_ranking_methods = ["SPOTIS", "AROMAN", "TOPSIS"]
        ranking_reason_tr = (
            f"Karar ağacı çıktısı — Büyük örneklem ve yüksek varyasyon ({n_alt} alternatif, ort. CV = {mean_cv:.2f}): "
            "Alternatif sayısı arttıkça klasik yöntemlerde sıralama tersine çevirme (rank reversal) riski yükselir. "
            "SPOTIS (Stable Preference Ordering Towards Ideal Solution), sabit bir ideal referans noktası kullanır; "
            "alternatif eklense veya çıkarılsa bile sıralama kararlı kalır — büyük kümelerde bu özellik rakipsizdir. "
            "AROMAN (Alternative Ranking Order Method Accounting for two-step Normalization), "
            "iki aşamalı normalizasyon ile farklı ölçek ve birim yapılarına karşı dirençlidir ve büyük alternatif setlerinde tutarlı sonuç verir. "
            "TOPSIS ise ölçeklenebilir referans çözümü olarak doğrulama amacıyla tabloda yer alır."
        )
        ranking_reason_en = (
            f"Decision-tree output — Large sample and high variation ({n_alt} alternatives, mean CV = {mean_cv:.2f}): "
            "As the number of alternatives grows, the risk of rank reversal in classical methods increases. "
            "SPOTIS (Stable Preference Ordering Towards Ideal Solution) uses a fixed ideal reference point; "
            "rankings remain stable even when alternatives are added or removed — an unmatched property in large sets. "
            "AROMAN (Alternative Ranking Order Method Accounting for two-step Normalization) "
            "resists different scale and unit structures through two-step normalisation and delivers consistent results in large alternative sets. "
            "TOPSIS is included as a scalable benchmark for verification."
        )
    elif max_corr >= 0.75 and cost_ratio >= 0.30:
        suggested_ranking_methods = ["VIKOR", "RAWEC", "PROMETHEE"]
        ranking_reason_tr = (
            f"Karar ağacı çıktısı — Yüksek kriter çatışması ve karma yön yapısı (maks. r = {max_corr:.2f}, maliyet oranı = %{cost_ratio*100:.0f}): "
            "Hem kriterler arası korelasyon güçlüyse hem de maliyet kriterleri ağırlıklıysa, "
            "sonuçlar optimizasyon yönüne karşı duyarlı hale gelir. "
            "VIKOR bu koşulda uzlaşı çözümü üretmekte en güçlü yöntemdir. "
            "RAWEC (Ranking with Weights of Criterion Change), her kriterin ağırlığındaki değişime karşı sıralama stabilitesini test eder; "
            "çatışmalı kriter yapılarında hangi alternatiflerin gerçekten sağlam olduğunu ortaya koyar. "
            "PROMETHEE ise ikili karşılaştırma mantığıyla maliyet-fayda dengesi kurmanın en okunur yolunu sunar."
        )
        ranking_reason_en = (
            f"Decision-tree output — High criterion conflict and mixed direction structure (max r = {max_corr:.2f}, cost ratio = {cost_ratio*100:.0f}%): "
            "When inter-criterion correlation is strong and cost criteria dominate, "
            "results become sensitive to the optimisation direction. "
            "VIKOR is the strongest method for producing a compromise solution in this setting. "
            "RAWEC (Ranking with Weights of Criterion Change) tests ranking stability against changes in each criterion's weight, "
            "revealing which alternatives are genuinely robust under conflicting criterion structures. "
            "PROMETHEE offers the most readable path to balancing cost and benefit through pairwise comparison logic."
        )
    elif alt_crit_ratio < 2.0:
        suggested_ranking_methods = ["MARCOS", "ROV", "RAFSI"]
        ranking_reason_tr = (
            f"Karar ağacı çıktısı — Sıkışık problem boyutu (alt/krit = {alt_crit_ratio:.1f}): "
            "Alternatif sayısı kriter sayısına yakın olduğunda birçok yöntem kararsız veya gürültülü sonuç üretebilir. "
            "MARCOS (Measurement of Alternatives and Ranking according to COmpromise Solution), "
            "hem ideal hem de karşı-ideal çözüme olan uzaklığı yardımcı işlevlerle birleştirir; "
            "az alternatifli sıkışık yapılarda dengeli ve kararlı bir referans çözüm sunar. "
            "ROV (Range of Value), her alternatifin kriter aralığındaki konumunu maksimum ve minimum fayda sınırları arasında değerlendirir; "
            "ölçek farklılıklarına dayanıklıdır. "
            "RAFSI ise fonksiyonel alt-aralık eşlemesiyle az alternatifte bile tutarlı oransal sıralar üretir."
        )
        ranking_reason_en = (
            f"Decision-tree output — Tight problem size (alt/crit = {alt_crit_ratio:.1f}): "
            "When the number of alternatives is close to the number of criteria, many methods may produce unstable or noisy results. "
            "MARCOS (Measurement of Alternatives and Ranking according to COmpromise Solution) "
            "combines distances to both the ideal and anti-ideal solutions through utility functions, "
            "providing a balanced and stable reference solution in tight, low-alternative structures. "
            "ROV (Range of Value) evaluates each alternative's position within the criterion range between maximum and minimum utility bounds; "
            "it is robust to scale differences. "
            "RAFSI produces consistent proportional rankings even with few alternatives through functional sub-interval mapping."
        )
    elif max_cv >= 0.45:
        suggested_ranking_methods = ["ROV", "RAWEC", "TOPSIS"]
        ranking_reason_tr = (
            f"Karar ağacı çıktısı — Geniş ölçek aralığı (maks. CV = {max_cv:.2f}, ort. CV = {mean_cv:.2f}): "
            "En az bir kriter çok geniş bir değer aralığına sahipken normalizasyon seçimi kritik hale gelir. "
            "ROV (Range of Value), her kriter için tam aralığı referans alır; "
            "uç değerleri olan kriterlerin ağırlığını doğrudan aralık büyüklüğüne göre ölçeklendirir ve ölçek baskınlığını bastırır. "
            "RAWEC ise ağırlık değişimlerine göre sıralama stabilitesini test ederek geniş aralıklı kriterlerdeki duyarlılığı görünür kılar. "
            "TOPSIS doğrulama referansı olarak tabloda yer alır."
        )
        ranking_reason_en = (
            f"Decision-tree output — Wide scale range (max CV = {max_cv:.2f}, mean CV = {mean_cv:.2f}): "
            "When at least one criterion spans a very wide value range, the choice of normalisation becomes critical. "
            "ROV (Range of Value) uses the full range of each criterion as a reference, "
            "scaling the weight of criteria with extreme values according to range size directly and suppressing scale dominance. "
            "RAWEC tests ranking stability against weight changes, making sensitivity in wide-range criteria visible. "
            "TOPSIS is included as a verification benchmark."
        )
    elif mean_cv <= 0.14 and max_corr <= 0.55:
        suggested_ranking_methods = ["VIKOR", "AROMAN", "MARCOS"]
        ranking_reason_tr = (
            f"Karar ağacı çıktısı — Homojen ve düşük çatışmalı profil (ort. CV = {mean_cv:.2f}, maks. r = {max_corr:.2f}): "
            "Kriterler az ayrıştığında ve korelasyon düşük olduğunda alternatiflerin birbirinden belirgin biçimde ayrılması güçleşir. "
            "VIKOR, bu koşulda uzlaşı çözümü üreterek yakın performanslı alternatifler arasındaki ince farkları en iyi okur. "
            "AROMAN, iki aşamalı normalizasyonuyla homojen yapılarda bile küçük performans farklarını kararlı biçimde sıralar. "
            "MARCOS ideal ve karşı-ideal dengesini yardımcı işlevlerle birleştirerek homojen profillerde tutarlı ve yorumlanabilir bir sıralama sunar."
        )
        ranking_reason_en = (
            f"Decision-tree output — Homogeneous and low-conflict profile (mean CV = {mean_cv:.2f}, max r = {max_corr:.2f}): "
            "When criteria discriminate weakly and correlation is low, it becomes difficult to separate alternatives clearly. "
            "VIKOR reads fine differences among close-performing alternatives best in this setting by generating a compromise solution. "
            "AROMAN ranks small performance differences stably even in homogeneous structures through two-step normalisation. "
            "MARCOS combines ideal and anti-ideal balance through utility functions, delivering a consistent and interpretable ranking in homogeneous profiles."
        )
    elif n_alt >= 10 and mean_cv >= 0.15:
        suggested_ranking_methods = ["TOPSIS", "SPOTIS", "EDAS"]
        ranking_reason_tr = (
            f"Karar ağacı çıktısı — Orta-büyük örneklem ve orta varyasyon ({n_alt} alternatif, ort. CV = {mean_cv:.2f}): "
            "Bu profil dengeli bir veri yapısını temsil eder. "
            "TOPSIS, ideal çözüme göreli yakınlığı Öklid uzaklığıyla ölçen, akademik literatürde en yaygın referans yöntemidir. "
            "SPOTIS, sıralama kararlılığını güvence altına almak için tabloda yer alır; "
            "sabit ideal noktası sayesinde alternatif setinde olası değişiklikler sıralamayı tersine çevirmez. "
            "EDAS ise sapma tabanlı mantığıyla TOPSIS'in mesafe temelli sonucunu bağımsız biçimde doğrular."
        )
        ranking_reason_en = (
            f"Decision-tree output — Medium-to-large sample and moderate variation ({n_alt} alternatives, mean CV = {mean_cv:.2f}): "
            "This profile represents a balanced data structure. "
            "TOPSIS measures relative closeness to the ideal solution via Euclidean distance and is the most widely used reference method in the academic literature. "
            "SPOTIS is included to safeguard ranking stability: "
            "its fixed ideal point ensures that potential changes in the alternative set do not reverse the ranking. "
            "EDAS independently verifies TOPSIS's distance-based result through its deviation-based logic."
        )
    else:
        suggested_ranking_methods = ["TOPSIS", "VIKOR", "AROMAN"]
        ranking_reason_tr = (
            f"Karar ağacı çıktısı — Genel dengeli profil (ort. CV = {mean_cv:.2f}, maks. r = {max_corr:.2f}, çarpıklık = {mean_abs_skew:.2f}): "
            "Veri yapısı belirgin bir risk bayrağı taşımıyor; kriterler orta düzeyde ayrışıyor ve korelasyon sınırlı. "
            "TOPSIS bu profil için genel referans çözümü sağlar ve akademik olarak en yaygın doğrulama aracıdır. "
            "VIKOR uzlaşı boyutunu test eder; birden fazla yöntemin önerdiği alternatifi kümülatif sıralamada öne taşır. "
            "AROMAN, iki aşamalı normalizasyonuyla ölçek farklılıklarına karşı dayanıklı bir üçüncü bakış açısı ekler "
            "ve TOPSIS/VIKOR ile tutarlılık kontrolüne olanak tanır."
        )
        ranking_reason_en = (
            f"Decision-tree output — General balanced profile (mean CV = {mean_cv:.2f}, max r = {max_corr:.2f}, skewness = {mean_abs_skew:.2f}): "
            "The data structure carries no prominent risk flag; criteria discriminate at a moderate level and correlation is limited. "
            "TOPSIS provides the general benchmark solution for this profile and is the most widely validated tool in the academic literature. "
            "VIKOR tests the compromise dimension, elevating the alternative that multiple methods agree on in cumulative ranking. "
            "AROMAN adds a resistant third perspective through two-step normalisation and enables a consistency check against TOPSIS and VIKOR."
        )

    suggested_ranking_methods = suggested_ranking_methods[:3]
    suggested_ranking = " / ".join(suggested_ranking_methods)

    recommendations.append({
        "level": "info",
        "icon": "⚖️",
        "text": weight_reason_tr,
        "action": f"Önerilen ağırlıklandırma yöntemi: {suggested_weight}.",
        "text_en": weight_reason_en,
        "action_en": f"Recommended weighting method: {suggested_weight}.",
    })
    recommendations.append({
        "level": "info",
        "icon": "🏆",
        "text": ranking_reason_tr,
        "action": f"Önerilen sıralama yöntemleri: {', '.join(suggested_ranking_methods)}.",
        "text_en": ranking_reason_en,
        "action_en": f"Recommended ranking methods: {', '.join(suggested_ranking_methods)}.",
    })

    if not cost:
        recommendations.append({
            "level": "info",
            "icon": "ℹ️",
            "text": "Tüm kriterler fayda yönlü (Max). Maliyet kriteri bulunmuyor.",
            "action": "Gerçek analizlerde genellikle en az bir maliyet kriteri olur. Kriter yönlerini gözden geçirin.",
            "text_en": "All criteria are currently marked as benefit criteria; there is no cost criterion.",
            "action_en": "Review criterion directions and verify whether at least one criterion should be modeled as a cost.",
        })
        risk_flags.append("no_cost_criterion")

    return {
        "n_alt": n_alt,
        "n_crit": n_crit,
        "mean_cv": mean_cv,
        "max_cv": max_cv,
        "cv_std": cv_std,
        "max_corr": max_corr,
        "mean_abs_skew": mean_abs_skew,
        "outlier_ratio": outlier_ratio,
        "high_corr_pairs": high_corr_pairs,
        "constant_criteria": constant,
        "non_positive_criteria": non_positive,
        "alt_crit_ratio": alt_crit_ratio,
        "risk_flags": risk_flags,
        "cv_per_criterion": cv_vals,
        "recommendations": recommendations,
        "suggested_weight": suggested_weight,
        "suggested_ranking": suggested_ranking,
        "suggested_ranking_methods": suggested_ranking_methods,
        "weight_rationale_tr": weight_reason_tr,
        "weight_rationale_en": weight_reason_en,
        "ranking_rationale_tr": ranking_reason_tr,
        "ranking_rationale_en": ranking_reason_en,
        "benefit_count": len(benefit),
        "cost_count": len(cost),
    }


def generate_3layer_weight(
    weight_method: str,
    weights: Dict[str, float],
    n_alt: int,
) -> Dict[str, str]:
    """Ağırlıklandırma sonuçları için 3 kademeli yorum üretir."""
    sorted_w = sorted(weights.items(), key=lambda x: x[1], reverse=True)
    top_crit, top_w = sorted_w[0]
    second_crit, second_w = sorted_w[1] if len(sorted_w) > 1 else sorted_w[0]
    bottom_crit, bottom_w = sorted_w[-1]
    concentration = top_w / max(second_w, EPS)

    descriptive = {
        "Entropy": (
            f"Shannon Entropi yöntemi {n_alt} alternatifli karar matrisini analiz etti. "
            f"En yüksek objektif ağırlık **{top_crit}** kriterine atandı "
            f"(w = {top_w:.4f}, %{top_w*100:.1f}). En düşük ağırlıklı kriter "
            f"**{bottom_crit}** (w = {bottom_w:.4f}) olarak belirlendi. "
            f"İlk iki kriter arasındaki ağırlık yoğunlaşma oranı {concentration:.2f}x."
        ),
        "CRITIC": (
            f"CRITIC yöntemi standart sapma ve korelasyon çatışmasını birlikte değerlendirerek "
            f"**{top_crit}** kriterini en önemli unsur (w = {top_w:.4f}, %{top_w*100:.1f}) "
            f"olarak belirledi. Bu kriter hem geniş varyansa hem de diğer kriterlerle "
            f"düşük korelasyona sahiptir. **{bottom_crit}** en az bilgi içeriği taşıyor (w = {bottom_w:.4f})."
        ),
        "Standart Sapma": (
            f"Standart sapma yöntemi alternatifleri en fazla ayırt eden kriteri öne çıkardı: "
            f"**{top_crit}** (w = {top_w:.4f}, %{top_w*100:.1f}). **{bottom_crit}** ise "
            f"en az değişken kriter olarak en düşük ağırlığı aldı (w = {bottom_w:.4f}). "
            f"İlk iki kriter arasındaki ağırlık oranı {concentration:.2f}x."
        ),
        "MEREC": (
            f"MEREC, kriterlerin sistemden çıkarılma etkisini ölçerek **{top_crit}** kriterini "
            f"en kritik unsur (w = {top_w:.4f}) olarak tespit etti. Bu kriter çıkarıldığında "
            f"toplam performans skoru en büyük bozulmayı yaşıyor. "
            f"**{second_crit}** ikinci sırada (w = {second_w:.4f})."
        ),
        "LOPCOW": (
            f"LOPCOW logaritmik analizi **{top_crit}** kriterinin normalize edilmiş yapıda "
            f"en yüksek bilgi gücüne sahip olduğunu ortaya koydu (w = {top_w:.4f}). "
            f"Ölçek farklılıkları logaritmik dönüşümle etkisizleştirildi. "
            f"**{bottom_crit}** en düşük bilgi içeriğine sahip (w = {bottom_w:.4f})."
        ),
        "PCA": (
            f"PCA tabanlı ağırlıklandırma korelasyon yapısından türetilen bileşenlerle "
            f"**{top_crit}** kriterini (w = {top_w:.4f}) baskın bilgi kaynağı olarak belirledi. "
            f"Kaiser eşiğini aşan bileşenler ve yük dağılımları birlikte değerlendirildi."
        ),
    }

    analytic = {
        "Entropy": (
            f"Shannon bilgi kuramına göre düşük entropi → yüksek ayırt edicilik. "
            f"**{top_crit}** kriterindeki normalize değerlerin dağılımı sistematik farklılaşma "
            f"örüntüsü sergilediğinden modelin birincil bilgi kaynağı konumuna yükselmiştir. "
            f"İkinci sıradaki **{second_crit}** ile ağırlık konsantrasyonu {concentration:.2f}x "
            f"olup modelde belirgin bir kriter hiyerarşisi mevcuttur. Bu yapı, "
            f"kararın tek bir kriter üzerinde aşırı yoğunlaşmadığını "
            f"{'gösterir (sağlıklı dağılım).' if concentration < 2.5 else 'işaret eder (dominant kriter riski).'}"
        ),
        "CRITIC": (
            f"CRITIC'in iki boyutlu değerlendirmesi — varyans genişliği ve çatışma yapısı — "
            f"**{top_crit}** kriterinin hem farklı bir ölçüm alanı kapladığını hem de "
            f"bağımsız bir bilgi kanalı oluşturduğunu matematiksel olarak kanıtlamaktadır. "
            f"Yüksek kontrast değeri, bu kriterin sıralama sonucunu diğerlerinden bağımsız "
            f"biçimde yönlendirme kapasitesine sahip olduğunu göstermektedir. "
            f"Korelasyon yapısı analizde şeffaf biçimde modellendiği için CRITIC, "
            f"çakışan kriter setlerinde Entropy'e kıyasla daha savunulabilir bir seçimdir."
        ),
        "Standart Sapma": (
            f"Standart sapma yöntemi alternatif performansları arasındaki mutlak yayılımı "
            f"ağırlıklandırma ölçütü olarak kullanır. **{top_crit}** kriterindeki geniş dağılım, "
            f"bu kriterin alternatifleri birbirinden en güçlü biçimde ayırt ettiğini "
            f"kanıtlamaktadır. Ancak bu yöntem kriterler arası korelasyonu gözetmez; "
            f"yüksek korelasyonlu veri setlerinde CRITIC daha uygun bir alternatiftir."
        ),
        "MEREC": (
            f"Çıkarma etkisi analizi **{top_crit}** kriterinin bütüncül performans skoru "
            f"üzerindeki marjinal etkisinin diğer tüm kriterlerden yüksek olduğunu göstermektedir. "
            f"Bu bulgu, söz konusu kriterin model dayanıklılığı için zorunlu olduğuna işaret eder. "
            f"MEREC, logaritmik normalize edilmiş sistem performansını referans aldığından "
            f"kriter değerlerindeki orantısal değişimlere duyarlı ve ölçek-bağımsız bir "
            f"ağırlık vektörü üretmektedir."
        ),
        "LOPCOW": (
            f"Logaritmik yüzde değişim gücü analizi **{top_crit}** kriterinin normalize edilmiş "
            f"RMS/standart sapma oranının diğerlerine kıyasla belirgin biçimde yüksek "
            f"olduğunu ortaya koymaktadır. Bu oran, söz konusu kriterin sistematik bilgi "
            f"taşıdığını ve ölçek farkından bağımsız olarak ayırt edicilik gücünü "
            f"koruduğunu göstermektedir. LOPCOW özellikle farklı ölçüm birimli veri setlerinde "
            f"Entropy ve SD'ye kıyasla daha güvenilir sonuç üretir."
        ),
        "PCA": (
            f"Temel bileşen analizi veri setindeki korelasyon yapısını özdeğer ayrışımıyla "
            f"çözümledi. **{top_crit}** kriterinin Kaiser eşiğini (λ > 1) aşan "
            f"bileşenlerdeki ağırlıklı yük toplamı en yüksek değeri aldı; "
            f"bu durum söz konusu kriterin ortak varyans yapısının merkezinde yer aldığını "
            f"göstermektedir. PCA ağırlıklandırması özellikle yüksek boyutlu ve çok "
            f"korelasyonlu veri setlerinde tercih edilmelidir."
        ),
    }

    normative = {
        "Entropy": (
            f"Karar vericiler **{top_crit}** kriteri üzerindeki performans farklılıklarını "
            f"öncelikli değerlendirme ekseni olarak benimsemelidir. Bu kriter alternatifler "
            f"arasındaki stratejik ayrışımı en net biçimde ortaya koymaktadır. "
            f"Akademik yayın sürecinde hakemler ağırlık gerekçesini sorguladığında "
            f"Shannon entropi tabanlı hesaplama, subjektif yargıdan bağımsız matematiksel "
            f"bir zemin sunar. **{bottom_crit}** kriterinin düşük ağırlığı bu kriterin "
            f"karar sürecindeki rolünün yeniden sorgulanmasını önerir."
        ),
        "CRITIC": (
            f"**{top_crit}** kriterindeki üstünlük diğer kriterlerle çakışmayan bağımsız "
            f"bir değer alanı temsil ettiğinden, bu kriterdeki iyileşme doğrudan genel "
            f"sıralama üzerinde etkili olacaktır. Yöneticiler kaynak tahsisi ve "
            f"stratejik önceliklendirmede bu kriteri birincil kontrol noktası olarak "
            f"konumlandırmalıdır. Akademik çalışmada CRITIC'in çatışma matrisini "
            f"raporlamak hakemlere metodolojik şeffaflık sağlar."
        ),
        "Standart Sapma": (
            f"**{top_crit}** kriterindeki geniş performans yelpazesi bu alanda rakipler "
            f"arasındaki farkın kapatılabileceğine ya da korunabileceğine işaret eder. "
            f"Karar vericiler bu kriteri rekabet avantajı veya risk yönetimi açısından "
            f"birincil odak noktası olarak ele almalıdır. Ancak bu yöntemin korelasyon "
            f"körlüğüne dikkat edin: yüksek korelasyon varlığında bulgular CRITIC ile "
            f"çapraz doğrulanmalıdır."
        ),
        "MEREC": (
            f"**{top_crit}** kriterinin sistemik önemi göz önüne alındığında bu kriterdeki "
            f"veri kalitesinin doğrulanması ve gerekirse ek ölçüm yapılması önerilir. "
            f"Bu kriter devre dışı kalması durumunda tüm modelin geçerliliği sorgulanabilir; "
            f"dolayısıyla yedek ölçüm protokolü oluşturulmalıdır. Akademik raporda "
            f"çıkarma etkisi tablosunun sunulması metodolojik zenginlik katar."
        ),
        "LOPCOW": (
            f"**{top_crit}** kriterinin normalize edilmiş üstünlüğü ölçek bağımsız bir "
            f"gerçekliği yansıttığından farklı birimler kullanılan uluslararası "
            f"karşılaştırma çalışmalarında da geçerliliğini koruyacaktır. "
            f"Bu kriter, çapraz ülke veya çapraz sektör analizlerinde birincil "
            f"referans değişken olarak kullanılabilir."
        ),
        "PCA": (
            f"PCA'nın ortaya koyduğu gizli yapısal boyut **{top_crit}** kriterinin "
            f"bağımsız değerlendirme kriterleri arasında en az gürültüye sahip "
            f"olanı olduğunu göstermektedir. Yöneticiler bu kriteri stratejik "
            f"performans karnesi tasarımında öncü gösterge olarak konumlandırabilir. "
            f"Akademik çalışmada PCA scree plot ve bileşen yük tablosunun sunulması "
            f"bulguların şeffaflığını artırır."
        ),
    }

    return {
        "descriptive": descriptive.get(weight_method, f"En yüksek ağırlık **{top_crit}** kriterine atandı (w = {top_w:.4f})."),
        "analytic": analytic.get(weight_method, METHOD_PHILOSOPHY.get(weight_method, {}).get("academic", "")),
        "normative": normative.get(weight_method, f"Karar vericiler **{top_crit}** kriterini öncelikli değerlendirmelidir."),
    }


def generate_3layer_ranking(
    ranking_method: str,
    ranking_table: pd.DataFrame,
    weights: Dict[str, float],
) -> Dict[str, str]:
    """Sıralama sonuçları için 3 kademeli yorum üretir."""
    if ranking_table is None or ranking_table.empty:
        return {"descriptive": "", "analytic": "", "normative": ""}

    n = len(ranking_table)
    top_alt = str(ranking_table.iloc[0]["Alternatif"])
    top_score = float(ranking_table.iloc[0]["Skor"])
    second_alt = str(ranking_table.iloc[1]["Alternatif"]) if n > 1 else top_alt
    second_score = float(ranking_table.iloc[1]["Skor"]) if n > 1 else top_score
    last_alt = str(ranking_table.iloc[-1]["Alternatif"])
    last_score = float(ranking_table.iloc[-1]["Skor"])
    lower_is_better = _is_lower_better_method(ranking_method)
    gap = (second_score - top_score) if lower_is_better else (top_score - second_score)
    score_range = (last_score - top_score) if lower_is_better else (top_score - last_score)
    relative_gap = gap / score_range if score_range > EPS else 0.0
    dominance = "belirgin" if relative_gap > 0.15 else "sınırlı"
    is_fuzzy = ranking_method.startswith("Fuzzy")
    base_method = ranking_method.replace("Fuzzy ", "") if is_fuzzy else ranking_method

    fuzzy_prefix = (
        f"Bulanık {base_method} analizi üçgensel bulanık sayılar aracılığıyla ölçüm "
        f"belirsizliği modellenmiş verilerle {n} alternatifi değerlendirdi. "
    ) if is_fuzzy else ""

    desc_templates = {
        "TOPSIS": (
            f"{fuzzy_prefix}TOPSIS analizi {n} alternatifin ideal ve anti-ideal çözüme "
            f"göreli yakınlık katsayılarını hesapladı. **{top_alt}** en yüksek yakınlık "
            f"skoru ({top_score:.4f}) ile birinci sıraya yerleşti. "
            f"**{second_alt}** ile fark {gap:.4f} — lider üstünlüğü **{dominance}** düzeyde."
        ),
        "VIKOR": (
            f"{fuzzy_prefix}VIKOR uzlaşı analizi grup faydası (S) ve maksimum bireysel "
            f"pişmanlık (R) ölçütlerini dengeleyerek **{top_alt}** alternatifini uzlaşı "
            f"çözümü olarak belirledi (Q: {top_score:.4f}). "
            f"**{second_alt}** ikinci sıradadır (fark: {gap:.4f})."
        ),
        "EDAS": (
            f"{fuzzy_prefix}EDAS ortalama çözüm analizi pozitif ve negatif mesafeleri "
            f"dengeleyerek **{top_alt}** alternatifini (skor: {top_score:.4f}) birinci "
            f"sıraya koydu. {n} alternatif arasındaki fark **{dominance}** düzeyde."
        ),
        "CODAS": (
            f"CODAS analizi Öklid ve Manhattan uzaklıklarını birlikte kullanarak "
            f"**{top_alt}** alternatifini (H skor: {top_score:.4f}) negatif idealden "
            f"en uzak konumda tespit etti. **{last_alt}** en düşük performansı sergiledi."
        ),
        "COPRAS": (
            f"{fuzzy_prefix}COPRAS analizi fayda ve maliyet bileşenlerini ayrıştırarak "
            f"**{top_alt}** alternatifini en yüksek göreli önem düzeyiyle (skor: {top_score:.4f}) "
            f"ilk sırada konumlandırdı."
        ),
        "OCRA": (
            f"{fuzzy_prefix}OCRA analizi fayda ve maliyet rekabet puanlarını birlikte "
            f"değerlendirerek **{top_alt}** alternatifini (skor: {top_score:.4f}) operasyonel "
            f"üstünlüğü en yüksek seçenek olarak belirledi."
        ),
        "ARAS": (
            f"{fuzzy_prefix}ARAS fayda katsayısı analizi ideal alternatife göreli "
            f"performansı ölçerek **{top_alt}** alternatifini (K = {top_score:.4f}) "
            f"en yüksek fayda oranına sahip alternatif olarak belirledi."
        ),
        "SAW": (
            f"{fuzzy_prefix}SAW ağırlıklı toplamsal fayda analizi, normalize edilmiş "
            f"kriter katkılarını birleştirerek **{top_alt}** alternatifini "
            f"(skor: {top_score:.4f}) en üst sıraya yerleştirdi."
        ),
        "WPM": (
            f"{fuzzy_prefix}WPM çarpımsal fayda analizi, kriterleri üstel ağırlıklarla "
            f"birleştirerek **{top_alt}** alternatifini (skor: {top_score:.4f}) "
            f"oran-temelli üstünlükle lider belirledi."
        ),
        "MAUT": (
            f"{fuzzy_prefix}MAUT toplam fayda analizi, kriterleri fayda fonksiyonuna "
            f"dönüştürüp ağırlıklı toplulaştırarak **{top_alt}** alternatifini "
            f"(skor: {top_score:.4f}) en yüksek beklenen faydaya sahip seçenek olarak tespit etti."
        ),
        "WASPAS": (
            f"{fuzzy_prefix}WASPAS hibrit analizi (WSM + WPM bileşimi) {n} alternatifi "
            f"hem toplamsal hem çarpımsal mantıkla değerlendirerek **{top_alt}** "
            f"alternatifini (skor: {top_score:.4f}) birinci sıraya taşıdı."
        ),
        "MOORA": (
            f"MOORA oran analizi fayda ve maliyet kriterlerini normalize edilmiş "
            f"ortak ölçekte sentezledi. **{top_alt}** en yüksek net skor ({top_score:.4f}) "
            f"ile lider. **{last_alt}** en düşük performansta."
        ),
        "MABAC": (
            f"MABAC sınır yaklaşım alanı analizi **{top_alt}** alternatifini "
            f"(skor: {top_score:.4f}) sınırın en üzerinde konumlandırdı. "
            f"**{last_alt}** sınırın altında kaldı."
        ),
        "MARCOS": (
            f"MARCOS referans analizi **{top_alt}** alternatifini (yararlılık: {top_score:.4f}) "
            f"hem ideal hem anti-ideal referansla karşılaştırmada en dengeli konumda "
            f"belirledi. **{last_alt}** en düşük yararlılık skoruna sahip."
        ),
        "CoCoSo": (
            f"CoCoSo kombine uzlaşı analizi üç farklı uzlaşı stratejisini entegre "
            f"ederek **{top_alt}** alternatifini (bileşik skor: {top_score:.4f}) tutarlı "
            f"lider olarak tespit etti."
        ),
        "PROMETHEE": (
            f"{fuzzy_prefix}PROMETHEE akış analizi ikili tercih ilişkilerini hesaplayarak "
            f"**{top_alt}** alternatifini en yüksek net akış değerine sahip lider "
            f"alternatif olarak belirledi (skor: {top_score:.4f})."
        ),
        "GRA": (
            f"GRA gri ilişkisel analizi referans diziye benzerlik derecesini ölçerek "
            f"**{top_alt}** alternatifini (gri derece: {top_score:.4f}) ideal "
            f"profile en yakın alternatif olarak tanımladı."
        ),
    }

    analytic_templates = {
        "TOPSIS": (
            f"PIS'e yakınlık (D⁺) ile NIS'ten uzaklık (D⁻) birlikte değerlendirildiğinde "
            f"**{top_alt}** her iki boyutta avantajlı konumda. Yakınlık katsayısı "
            f"{top_score:.4f} ile **{dominance}** bir liderlik sergileniyor. "
            f"{'Fark sınırlı olduğundan ikinci seçeneğin sağlamlık testine alınması önerilir.' if dominance == 'sınırlı' else 'Belirgin fark, kararı metodoloji değişimine karşı dirençli kılar.'}"
        ),
        "VIKOR": (
            f"VIKOR'un uzlaşı koşulları — 'kabul edilebilir avantaj' ve 'kararlılık' — "
            f"değerlendirildiğinde **{top_alt}** hem grup faydası (min S) hem de bireysel "
            f"pişmanlık (min R) açısından öne çıkmaktadır. v = 0.5 ile bireysel/grup "
            f"dengesi eşit ağırlıkta gözetildi. Q farkı kabul edilebilir avantaj eşiğinin "
            f"{'üzerinde — güçlü uzlaşı.' if relative_gap > 0.10 else 'sınırında veya altında — ihtiyatlı yorum önerilir.'}"
        ),
        "EDAS": (
            f"**{top_alt}** ortalama çözüme göre yüksek pozitif sapma (PDA) ve düşük "
            f"negatif sapma (NDA) değerleri sergiledi. Bu denge söz konusu alternatifin "
            f"fayda kriterlerinde sistematik üstünlük, maliyet kriterlerinde ise "
            f"sistematik kontrol sağladığına işaret eder. Ortalama çözüm referans "
            f"alındığından sektör standardı mantığıyla uyumlu bir değerlendirmedir."
        ),
        "CODAS": (
            f"Öklid (birincil) ve Manhattan (ikincil) uzaklık ölçütleri **{top_alt}** "
            f"alternatifinin negatif idealden hem küresel (Öklid) hem de çok boyutlu "
            f"(Manhattan) uzaklık açısından öne çıktığını gösteriyor. τ eşiği ayrıştırma "
            f"hassasiyetini kontrol ediyor; düşük τ değeri daha ince ayrımları ön plana çıkarır."
        ),
        "COPRAS": (
            f"COPRAS, ağırlıklı normalize matris üzerinden fayda toplamı (S+) ve maliyet "
            f"toplamını (S-) birlikte yorumlar. **{top_alt}**, yüksek S+ ve görece düşük "
            f"S- bileşimiyle en yüksek Q değerini üretmiştir; bu sonuç alternatifin hem "
            f"değer yaratma hem maliyet kontrolü açısından dengeli üstünlüğünü gösterir."
        ),
        "OCRA": (
            f"OCRA yaklaşımı fayda ve maliyet kriterlerinde göreli rekabet gücünü ayrı "
            f"hesaplayıp toplar. **{top_alt}** alternatifi her iki bileşenin toplamında "
            f"üst sıraya çıkmış, bu da operasyonel performansın tek boyutlu değil çok "
            f"boyutlu bir üstünlükle oluştuğunu göstermiştir."
        ),
        "ARAS": (
            f"**{top_alt}** fayda katsayısı K = {top_score:.4f}, ideal alternatife "
            f"göre {top_score*100:.1f}% düzeyinde performans sergilendiğini gösteriyor. "
            f"Bu oran {'yüksek — ideal yakınlık güçlü.' if top_score > 0.75 else 'orta düzey — iyileştirme potansiyeli mevcut.' if top_score > 0.50 else 'düşük — yapısal güçlükler var.'}"
        ),
        "SAW": (
            f"SAW modeli normalize katkıları doğrusal biçimde toplar; bu nedenle kriterler "
            f"arası telafi davranışı yüksektir. **{top_alt}** için toplam fayda "
            f"üstünlüğü, çok sayıda kriterde dengeli performans üretildiğine işaret eder."
        ),
        "WPM": (
            f"WPM çarpımsal yapısı düşük performanslı kriterleri daha güçlü cezalandırır. "
            f"**{top_alt}** alternatifinin liderliği, yalnızca birkaç kriterde değil, "
            f"genel kriter setinde tutarlı performans üretildiğini gösterir."
        ),
        "MAUT": (
            f"MAUT çerçevesinde her kriter doğrusal fayda ölçeğine taşınıp ağırlıklı "
            f"beklenen fayda hesaplanmıştır. **{top_alt}** alternatifinin yüksek toplam "
            f"faydası, karar verici tercihleriyle uyumlu rasyonel üstünlüğü gösterir."
        ),
        "WASPAS": (
            f"λ = 0.5 ile WSM (toplamsal) ve WPM (çarpımsal) bileşenleri eşit "
            f"ağırlıkla birleştirildi. **{top_alt}** her iki bileşende güçlü "
            f"performans sergileyerek telafi edici ve çarpımsal mantık altında "
            f"tutarlı liderliğini korudu. λ değişimine duyarlılık için parametre "
            f"analizi öncelikle bu alternatif üzerinde test edilmelidir."
        ),
        "MOORA": (
            f"Fayda kriter toplamından maliyet kriter toplamı çıkarıldığında **{top_alt}** "
            f"en yüksek net değeri aldı. Bu hem fayda tarafında güçlü hem de maliyet "
            f"kontrolünde etkili olduğunu göstermektedir. Vektör normalizasyonu "
            f"kriter ölçek farklılıklarını telafi ettiğinden çapraz kriter "
            f"karşılaştırması matematiksel olarak adil bir zeminde yapıldı."
        ),
        "MABAC": (
            f"Sınır yaklaşım alanı (BAA) üzerindeki konumlar **{top_alt}** alternatifinin "
            f"en fazla kriterde sınırın üzerinde kaldığını ortaya koyuyor. Geometrik ortalama "
            f"ile hesaplanan BAA, bireysel aşırı değerlere karşı direnç sağlar. "
            f"Negatif konumlu alternatiflerin kaç kriterde BAA altında kaldığı iyileştirme "
            f"önceliklendirmesi için doğrudan kullanılabilir."
        ),
        "MARCOS": (
            f"**{top_alt}** için K⁻ (anti-ideale oran) ve K⁺ (ideale oran) birlikte "
            f"dengelendiğinde yararlılık fonksiyonu maksimum değere ulaştı. Bu durum "
            f"söz konusu alternatifin hem güçlü pozitif yönleriyle hem de sınırlı negatif "
            f"sapmasıyla bütüncül üstünlük sergilediğini kanıtlamaktadır."
        ),
        "CoCoSo": (
            f"Ka (aritmetik ortalama), Kb (karşılaştırmalı oran) ve Kc (λ ağırlıklı) "
            f"üç strateji **{top_alt}** için tutarlı biçimde yüksek değer aldı. "
            f"Bu çoklu uzlaşı tutarlılığı liderliğin herhangi bir stratejiye özgü "
            f"olmadığını, genel performans üstünlüğünü yansıttığını kanıtlıyor."
        ),
        "PROMETHEE": (
            f"PROMETHEE'de pozitif akış (phi+) alternatifin diğerlerini ne ölçüde "
            f"aştığını, negatif akış (phi-) ise ne ölçüde aşıldığını temsil eder. "
            f"**{top_alt}** için net akışın en yüksek çıkması, ikili karşılaştırmaların "
            f"çoğunda sistematik üstünlük sağlandığını göstermektedir."
        ),
        "GRA": (
            f"Gri ilişkisel katsayılar **{top_alt}** alternatifinin referans diziye "
            f"(ideal profil) kriter bazında en benzer yapıda olduğunu göstermektedir. "
            f"ρ = 0.5 ayırt edicilik katsayısı altında hesaplanan gri ilişkisel derece, "
            f"veri belirsizliğini tolere eden güçlü bir performans kanıtı sunar."
        ),
    }

    normative_templates = {
        "TOPSIS": (
            f"**{top_alt}** uzun vadeli stratejik tercih için önerilir. Yüksek yakınlık "
            f"katsayısı hem kazanımları maksimize hem kayıpları minimize ettiğini gösteriyor. "
            f"{'İkinci sıradaki ' + second_alt + ' ile fark sınırlıysa ek kalitatif ölçüt değerlendirmesi yapılmalıdır.' if dominance == 'sınırlı' else 'Belirgin fark nedeniyle karar güvenle verilmelidir.'}"
        ),
        "VIKOR": (
            f"**{top_alt}** uzlaşı çözümü tüm paydaşların kabul edebileceği dengeli "
            f"bir seçeneği temsil eder. Karar vericiler yalnızca bireysel maksimuma "
            f"odaklanmak yerine grup faydasını ön planda tutmalıdır. "
            f"v parametresi risk iştahına göre ayarlanarak bireysel/grup dengesi optimize edilebilir."
        ),
        "EDAS": (
            f"**{top_alt}** sektör standardını belirleyici konumda. Bu alternatif "
            f"benchmark olarak kullanılarak diğer alternatiflerin iyileştirme hedefleri "
            f"belirlenebilir. Ortalama çözüme yakın alternatifler için hedefleme "
            f"maliyeti daha düşük olacaktır."
        ),
        "CODAS": (
            f"**{top_alt}** negatif idealden en uzak konumu nedeniyle en kötü senaryoya "
            f"karşı en yüksek direnci temsil ediyor. Risk yönetimi odaklı ortamlarda "
            f"'güvenli liman' stratejisi olarak konumlandırılabilir. "
            f"Alt sıradaki alternatiflerin Öklid ve Manhattan mesafeleri iyileştirme önceliklerini açıkça gösteriyor."
        ),
        "COPRAS": (
            f"**{top_alt}** COPRAS altında hem fayda üretimi hem maliyet baskısı bakımından "
            f"dengeli bir profil sergilediği için uygulama önceliğine alınmalıdır. "
            f"S+ ve S- bileşenleri birlikte izlenerek alternatifin güçlü/zayıf yönleri "
            f"operasyonel planlara doğrudan aktarılabilir."
        ),
        "OCRA": (
            f"**{top_alt}** OCRA skorunda öne çıktığı için rekabetçi operasyon hedeflerine "
            f"uygun bir ana adaydır. Yönetim açısından fayda ve maliyet alt skorlarının "
            f"ayrıştırılmış raporlanması, hangi kriterlerde iyileştirme yatırımı "
            f"gerektiğini daha net gösterir."
        ),
        "ARAS": (
            f"K = {top_score:.4f} değeri ideal alternatifte mümkün olan maksimum "
            f"fayda düzeyinin {top_score*100:.1f}%'ine erişildiğini gösteriyor. "
            f"Bu oran {'≥ %75 ile yüksek — mevcut yapı korunabilir.' if top_score >= 0.75 else '%50–75 arası — sistem genelinde iyileştirme potansiyeli var.' if top_score >= 0.50 else '< %50 — köklü yapısal değişim gerekiyor.'}"
        ),
        "SAW": (
            f"**{top_alt}** SAW altında en yüksek toplam faydayı ürettiği için "
            f"telafi edici karar yaklaşımında birincil adaydır. Politik ve "
            f"operasyonel senaryolarda hızlı, şeffaf ve izlenebilir bir karar "
            f"mekanizması için referans alternatif olarak kullanılmalıdır."
        ),
        "WPM": (
            f"**{top_alt}** WPM liderliği, zayıf kriter performanslarının cezalandırıldığı "
            f"daha katı değerlendirme altında da sürdüğü için dayanıklı bir seçimdir. "
            f"Riskten kaçınan karar vericiler için güçlü bir ana adaydır."
        ),
        "MAUT": (
            f"**{top_alt}** MAUT bağlamında en yüksek toplam faydayı sağladığından, "
            f"tercih-temelli rasyonel karar politikasında öncelikli alternatif "
            f"olarak ele alınmalıdır. Fayda fonksiyonu varsayımları raporda açıkça belirtilmelidir."
        ),
        "WASPAS": (
            f"**{top_alt}** hem toplamsal hem çarpımsal mantık altında tutarlı sonuç "
            f"verdiğinden bu seçim metodolojik tercihten bağımsızdır. λ parametresini "
            f"0.0–1.0 aralığında değiştirerek duyarlılık analizi yapılması ve "
            f"sıralamanın korunup korunmadığının raporlanması hakem güvenilirliğini artırır."
        ),
        "MOORA": (
            f"**{top_alt}** net skor üstünlüğüyle hem performans artırma hem maliyet "
            f"düşürme hedeflerini eş zamanlı karşılayabileceğini göstermektedir. "
            f"Çoklu kriterin dengeli ağırlıklandırıldığı politika değerlendirmelerinde "
            f"bu alternatif referans kabul edilebilir."
        ),
        "MABAC": (
            f"**{top_alt}** 'güvenli bölge' konumundadır — mevcut performansı "
            f"pekiştirme stratejisi benimsenebilir. Kaynaklar alt sıradaki "
            f"alternatifleri BAA seviyesine çıkarmaya yönlendirilebilir. "
            f"Portföy optimizasyonunda üst ve alt grup arasındaki gap kritik "
            f"iyileştirme bütçesini belirler."
        ),
        "MARCOS": (
            f"**{top_alt}** bütüncül yararlılık üstünlüğü, bu seçimin hem güncel kısıtlar "
            f"altında optimal hem gelecekteki kriter değişikliklerine karşı dirençli "
            f"olduğuna işaret ediyor. Uzun vadeli stratejik plan için güvenilir "
            f"referans alternatif budur."
        ),
        "CoCoSo": (
            f"Üç bağımsız uzlaşı stratejisinin **{top_alt}** üzerinde mutabık kalması "
            f"bu seçimin metodoloji bağımsız gerçekliği yansıttığını gösteriyor. "
            f"Farklı bakış açılarına sahip paydaşlara aritmetik, karşılaştırmalı ve "
            f"ağırlıklı argümanlarla tek bir tutarlı öneri sunulabilir."
        ),
        "PROMETHEE": (
            f"**{top_alt}** ikili üstünlük ilişkilerinde en yüksek net akışa sahip "
            f"olduğundan, özellikle alternatiflerin birbirini geçme davranışını açık "
            f"biçimde görmek istenen karar problemlerinde birincil aday olarak ele alınmalıdır."
        ),
        "GRA": (
            f"**{top_alt}** veri belirsizliği gözetildiğinde de sağlam kaldığından "
            f"gri sistem teorisinin önerdiği ihtiyatlı karar çerçevesinde kabul "
            f"edilebilir risk aralığında en güçlü adaydır. "
            f"Belirsizlik ortamlarında ve eksik veri durumlarında bu metodoloji "
            f"diğer yöntemlere kıyasla daha savunulabilir akademik zemin sunar."
        ),
    }

    # Fuzzy varyantlar için base yöntemden devral
    method_key = base_method if is_fuzzy else ranking_method
    fuzzy_analytic_suffix = (
        f" Bulanık modelleme centroid durulaştırması sonrası {top_score:.4f} skor "
        f"üretti; crisp karşılığından daha muhafazakâr bir performans kanalında "
        f"**{top_alt}** öne çıktı."
    ) if is_fuzzy else ""
    fuzzy_normative_suffix = (
        " Yüksek belirsizlik ortamlarında (oynaklık, eksik veri, dilsel değerlendirme) "
        "bulanık yaklaşım klasik yönteme kıyasla daha güçlü akademik meşruiyet sağlar."
    ) if is_fuzzy else ""

    return {
        "descriptive": desc_templates.get(
            method_key,
            f"**{top_alt}** alternatifi birinci sıraya yerleşti (skor: {top_score:.4f})."
        ),
        "analytic": analytic_templates.get(
            method_key,
            METHOD_PHILOSOPHY.get(ranking_method, {}).get("academic", "")
        ) + fuzzy_analytic_suffix,
        "normative": normative_templates.get(
            method_key,
            f"Karar vericiler **{top_alt}** alternatifini öncelikli değerlendirmelidir."
        ) + fuzzy_normative_suffix,
    }


def generate_3layer_sensitivity(sensitivity: Dict[str, Any]) -> Dict[str, str]:
    """Duyarlılık analizi sonuçları için 3 kademeli yorum üretir."""
    if not sensitivity:
        return {"descriptive": "", "analytic": "", "normative": ""}

    stability = float(sensitivity.get("top_stability", 0.0))
    base_top = str(sensitivity.get("base_top", ""))
    mc_df = sensitivity.get("monte_carlo_summary")
    mean_rank = float(mc_df.iloc[0]["OrtalamaSıra"]) if (mc_df is not None and not mc_df.empty) else 1.0

    level = "Yüksek" if stability >= 0.80 else ("Orta" if stability >= 0.60 else "Düşük")
    color_word = "güçlü" if stability >= 0.80 else ("kısmi" if stability >= 0.60 else "zayıf")

    return {
        "descriptive": (
            f"Monte Carlo simülasyonu ağırlık vektörüne rastgele gürültü enjekte ederek "
            f"binlerce senaryoda sıralama kararlılığını test etti. **{base_top}** alternatifinin "
            f"birinci sıraya gelme oranı **%{stability*100:.1f}** — model sağlamlığı "
            f"**{level}** düzeyde. Ortalama sıra konumu: {mean_rank:.2f}."
        ),
        "analytic": (
            f"Ağırlık pertürbasyonu altında %{stability*100:.1f} birincilik oranı "
            f"**{base_top}** liderliğinin rastlantısal değil yapısal olduğuna işaret eder. "
            + (
                "Yüksek kararlılık, ağırlık belirsizliğine karşın sıralama yapısının "
                "korunduğunu ve modelin dışsal şoklara dirençli olduğunu kanıtlıyor. "
                "Lokal senaryo analizinde tüm kriter pertürbasyonlarında sıra değişimi "
                "gözlemlenmiyorsa model istatistiksel anlamda sağlamdır."
                if stability >= 0.80 else
                "Orta düzey kararlılık bazı ağırlık konfigürasyonlarında sıralamada "
                "değişim olabileceğini gösteriyor. Bu durum tek bir yönteme güvenmek "
                "yerine çoklu yöntem karşılaştırmasının raporlanmasını zorunlu kılmaktadır."
                if stability >= 0.60 else
                "Düşük kararlılık kritik bir metodolojik uyarıdır: ağırlık vektörüne "
                "yapılan küçük bozulmalar sıralamayı köklü biçimde değiştirebilmektedir. "
                "Bu durum modelin parametrelere aşırı hassas olduğunu, veri setinin "
                "yeterli ayırt edicilik gücü taşımadığını veya yöntem seçiminin "
                "veri yapısıyla uyumsuz olduğunu gösterir."
            )
        ),
        "normative": (
            f"**{base_top}** üzerindeki {color_word} model sağlamlığı göz önüne alındığında: "
            + (
                "Bu bulgu makale veya raporda kesin öneri olarak sunulabilir. "
                f"Hakeme 'veriler biraz değişseydi ne olurdu?' sorusuna %{stability*100:.0f} "
                "kararlılık oranıyla güçlü yanıt verilmektedir. "
                "Monte Carlo özetini ve lokal duyarlılık grafiğini ek olarak sunun."
                if stability >= 0.80 else
                "Sıralama önerisi ihtiyat kaydıyla sunulmalıdır. VIKOR veya CoCoSo "
                "gibi uzlaşı odaklı yöntemlerin eklenmesi ve Spearman korelasyon "
                "uyumunun raporlanması modelin güvenilirliğini artırır. "
                "Alternatif olarak sigma değerini düşürerek daha dar pertürbasyon "
                "aralığında analizi tekrarlayın."
                if stability >= 0.60 else
                "Bu kararlılık seviyesinde kesin sıralama önerisi yapmaktan kaçının. "
                "Ağırlıklandırma yöntemini değiştirin, kriter setini gözden geçirin "
                "veya daha fazla alternatif ekleyin. Mevcut haliyle bu bulgular "
                "yalnızca keşifsel (exploratory) analiz niteliği taşımaktadır."
            )
        ),
    }


def generate_3layer_comparison(
    comparison: Dict[str, Any],
    base_method: str,
) -> Dict[str, str]:
    """Yöntem karşılaştırması için 3 kademeli yorum üretir."""
    if not comparison or "spearman_matrix" not in comparison:
        return {"descriptive": "", "analytic": "", "normative": ""}

    spearman_df = comparison["spearman_matrix"]
    methods_cols = [c for c in spearman_df.columns if c != "Yöntem"]
    n_methods = len(methods_cols)
    mat = spearman_df.set_index("Yöntem")
    upper = mat.where(~np.eye(mat.shape[0], dtype=bool)).stack()
    mean_rho = float(upper.mean()) if len(upper) > 0 else 1.0
    min_rho = float(upper.min()) if len(upper) > 0 else 1.0

    agreement = "yüksek" if mean_rho >= 0.85 else ("orta" if mean_rho >= 0.70 else "düşük")

    top_alts = comparison.get("top_alternatives")
    top_alt_text = ""
    if isinstance(top_alts, pd.DataFrame) and not top_alts.empty:
        counts = top_alts["BirinciAlternatif"].value_counts()
        dominant = counts.index[0]
        cnt = int(counts.iloc[0])
        top_alt_text = (
            f" **{dominant}** alternatifinin {cnt}/{n_methods} yöntemde birinci sıraya "
            f"gelmesiyle çapraz yöntem en tutarlı lider belirlendi."
        )

    return {
        "descriptive": (
            f"{n_methods} farklı yöntemin Spearman sıra korelasyon analizi tamamlandı. "
            f"Yöntemler arası ortalama uyum ρ = {mean_rho:.3f} — **{agreement}** düzeyde. "
            f"En düşük ikili uyum ρ = {min_rho:.3f}.{top_alt_text}"
        ),
        "analytic": (
            f"Ortalama Spearman ρ = {mean_rho:.3f} değeri farklı felsefi ailelere ait "
            f"yöntemlerin bu veri seti üzerinde "
            + (
                "birbirine yakın sıralama yapıları ürettiğini göstermektedir. "
                "Bu yüksek metodolojik uyum, bulguların yöntem seçiminden bağımsız "
                "yapısal gerçekliği yansıttığına güçlü kanıt sunar. "
                "Mesafe, uzlaşı ve fayda tabanlı yöntemlerin aynı lider üzerinde "
                "uzlaşması özellikle değerlidir."
                if mean_rho >= 0.85 else
                "kısmi uyum sergilediğini, tam mutabakat sağlanamadığını ortaya koyuyor. "
                "Yüksek uyumsuzluk gösteren çiftler veri setindeki çatışan kriter "
                "yapısının farklı felsefi açılardan yorumlandığına işaret edebilir. "
                "Bu metodolojik çeşitlilik akademik analizde güçlü bir bulgu."
                if mean_rho >= 0.70 else
                "önemli ölçüde farklı sıralama yapıları ürettiğini göstermektedir. "
                "Bu düşük uyum, yöntem seçiminin sonuçlar üzerinde belirleyici "
                "etkisinin olduğunu ve hangi felsefi yaklaşımın savunulacağının "
                "gerekçelendirilmesi gerektiğini ortaya koymaktadır."
            )
        ),
        "normative": (
            "Yüksek yöntem uyumu akademik çalışmada güçlü sağlamlık kanıtı sağlar; "
            "makale metodoloji bölümünde bu tutarlılık referans gösterilebilir. "
            "Spearman matrisini ek tablo olarak sunun."
            if mean_rho >= 0.85 else
            f"Yöntem seçimi gerekçelendirilmeli: neden **{base_method}** tercih edildiği "
            "ve alternatif yöntemlerin neden ikincil tutulduğu metodoloji bölümünde "
            "açıkça ifade edilmelidir. Karma yöntem stratejisi (hibrit) değerlendirin."
            if mean_rho >= 0.70 else
            "Yüksek yöntem duyarlılığı varlığında tek bir yönteme bağlı kalmak "
            "akademik savunulabilirliği zayıflatır. Birden fazla yöntemin ortak öneri "
            "ürettiği alternatifleri tercih edin veya uzlaşı yöntemlerine "
            "(VIKOR, CoCoSo) ağırlık verin. Uyumsuzluk nedenini veri yapısıyla "
            "ilişkilendirerek metodoloji sınırlılıkları bölümünde raporlayın."
        ),
    }
