# Hakediş Kontrol Sistemi

`hakedis_kontrol.py` — Götürü bedel sözleşme için hakediş kontrol Excel'i üretir.

## Ne yapar
- **Sayfa 1 (Hakediş Özet):** Sözleşme bedeli, tamamlanma yüzdesi, kümülatif hakediş, önceki ödemeler, KDV (%20), stopaj (-%3), net ödeme — tümü formüllü
- **Sayfa 2 (Malzeme Kalemleri):** 20 satır malzeme/ekipman girişi; miktar × birim fiyat × KDV otomatik hesaplıyor
- Sarı hücreler = kullanıcı girer / doğrulaması gerekir
- Yeşil hücreler = hesaplanan değerler
- Kırmızı = stopaj kesintisi

## Çalıştırma
```bash
python3 skills/hakedis_kontrol.py \
  --taseron "Firma Adı" \
  --sozlesme "SZL-2025-001" \
  --bedel 1000000 \
  --yuzde 0.65 \
  --kdv 0.20 \
  --stopaj 0.03 \
  --cikti /opt/asistan/ciktilar
```

## Parametreler
| Parametre | Açıklama | Varsayılan |
|-----------|----------|-----------|
| --taseron | Taşeron firma adı | "Taşeron Adı" |
| --sozlesme | Sözleşme numarası | "SZL-2025-001" |
| --bedel | Sözleşme bedeli (KDV hariç, ₺) | 1.000.000 |
| --yuzde | Tamamlanma yüzdesi (0-1) | 0.50 |
| --kdv | KDV oranı | 0.20 |
| --stopaj | Stopaj oranı | 0.03 |
| --cikti | Çıktı klasörü | /opt/asistan/ciktilar |
