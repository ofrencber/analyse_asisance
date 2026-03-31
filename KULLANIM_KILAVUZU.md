# MCDM Toolbox Kullanım Kılavuzu

Bu kılavuz, uygulamada bir analizi baştan sona nasıl yürüttüğümüzü adım adım açıklar.

## 1. Uygulamayı Açma

1. Ana sayfada isterseniz YouTube tanıtım videosu ve Instagram bağlantısını kullanın.

## 2. Veri Yükleme

1. Sol panelde **Veri Girişi** bölümünü açın.
2. `CSV` veya `XLSX` dosyanızı yükleyin.
3. Test için isterseniz **Örnek Veri Kullan** butonunu tıklayın.

Not:
- Analiz için en az 2 sayısal kriter gerekir.
- Metin sütunları otomatik olarak kriter dışı kalır.

## 3. Veri Ön İşleme

1. **Eksik Veri Tamamla** seçeneğini açın (opsiyonel).
2. Yöntem seçin: `Medyan`, `Ortalama`, `Interpolasyon`, `Sıfır`.
3. Gerekirse **Aykırı Değerleri Temizle** seçin.
4. Ayarları kontrol ettikten sonra devam edin.

## 4. Adım 1 - Analiz Amacı ve Veri Yapısı

1. Analiz amacını seçin:
   - Yalnızca ağırlık
   - Ağırlık + sıralama
2. Veri yapısını seçin:
   - Yıl verisi yok / önemsiz
   - Panel veri

## 5. Panel Veri Kullanımı (Seçildiyse)

1. **Yıl sütunu**nu seçin.
2. Dönemleri işaretleyin.
3. Hızlı kontrol butonları:
   - **Tümünü Seç**
   - **Tümünü Temizle**

Önemli:
- İlk açılışta dönemler seçili gelmez.
- Panel analiz için en az 1 dönem seçilmelidir.

## 6. Adım 2 - Kriter Doğrulama

1. Her kriter için **dahil/haric** seçimini yapın.
2. Kriter yönünü doğrulayın:
   - `Fayda (Max)`
   - `Maliyet (Min)`
3. **Veri Ön İşleme Bitti** butonu ile yöntem seçimine geçin.

## 7. Adım 3 - Yöntem Seçimi

1. Ağırlıklandırma modunu seçin:
   - Objektif
   - Eşit
   - Manuel
2. Ağırlık yöntemini seçin (ör. Entropy, CRITIC, SD, vb.).
3. Sıralama yöntem(ler)ini seçin (ör. TOPSIS, VIKOR, EDAS, vb.).

Not:
- Birden fazla sıralama yöntemi seçerseniz karşılaştırma çıktıları aktif olur.
- Sonuç ekranında yöntemlere ait sıralamalar alt sekmelerde gösterilir.

## 8. Parametreler

1. Otomatik önerilen parametreleri kullanabilir veya manuel değiştirebilirsiniz.
2. Gerekli ise şu parametreler ayarlanır:
   - VIKOR `v`
   - WASPAS `lambda`
   - CODAS `tau`
   - CoCoSo `lambda`
   - GRA `rho`
   - PROMETHEE `q/p/s` ve fonksiyon tipi
   - Fuzzy spread
   - Monte Carlo iterasyon/sigma

## 9. Analizi Çalıştırma

1. **Analiz Zamanı** butonuna basın.
2. Uygulama hesaplamayı tamamlayınca sonuç sekmeleri görünür.

Panel modunda:
- Seçilen bir yıl hata verirse uygulama diğer yıllarla devam eder.
- Uygun panel sonuç üretilemezse tek dönem fallback ile analiz tamamlanır.
- Atlanan dönemler sonuç ekranında uyarı olarak gösterilir.

## 10. Sonuçları Okuma

Sekmeler:
1. **Temel İstatistik**
2. **Ağırlıklar**
3. **Sıralama**
4. **Karşılaştırma**
5. **Sağlamlık**
6. **Dayanıklılık**
7. **Çıktı**

KPI bandında:
- Lider alternatif
- Sıralama yöntemi
- Kararlılık
- Analiz süresi
- Veri boyutu

## 11. Hesaplama Detayları

1. **Hesaplama Adımlarını Göster/Gizle** ile teknik detayları açın.
2. Alt sekmelerde inceleyin:
   - Ağırlık adımları
   - Sıralama adımları
   - Duyarlılık adımları
3. `Ağırlık detay tablosu` ve `Sıralama detay tablosu` ile ara hesapları görün.

## 12. Rapor ve Çıktı Alma

1. **Excel** çıktısı alın.
2. **DOCX** akademik raporu indirin.
3. Panel veri modunda tüm yıl özetleri çıktı dosyalarına dahil edilir.

## 13. Sık Karşılaşılan Durumlar

1. "En az 2 kriter seçmelisiniz":
   - Kriter dahil kutularını kontrol edin.
2. Panel analiz başlamıyor:
   - Yıl sütunu doğru mu?
   - En az bir dönem seçili mi?
3. Bazı yıllar görünmüyor:
   - O yıl için veri satırı veya yeterli alternatif olmayabilir.
4. Sonuçlar beklenenden farklı:
   - Kriter yönleri (max/min) ve manuel ağırlıkları tekrar kontrol edin.

## 14. Önerilen Kullanım Sırası

1. Veri yükle
2. Ön işleme kararını ver
3. Amaç + veri yapısını seç
4. Kriterleri doğrula
5. Yöntemleri seç
6. Parametreleri kontrol et
7. Analizi çalıştır
8. Sonuçları sekmelerde yorumla
9. Excel ve DOCX çıktısını al

