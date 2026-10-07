# Birim Fiyat Havuzu

`birim_fiyat_havuzu.py` — Polisan Kimya sözleşmelerinden çıkarılan birim fiyatları tek Excel'e toplar.

## Kapsam (2026 sözleşmeleri)
| Sözleşme | Firma | Poz | Kapsam |
|----------|-------|-----|--------|
| 2026/1013 | Özgül Teknik Metal İnş. | 135 | Yevmiye, Mekanik (boru/vana/pompa CS-SS-galvaniz), İnşaat, HDPE boru, Çelik imalat |
| 2026/1015 | Dönmez Kumlama Boya | 38 | Kumlama (bazalt/cüruf/grit), boyama, iskele, kimyasal temizlik |
| 2026/1023 | NOT İnşaat | 35 | PPRC/PVC boru, tesisat armatür, beton/kazı, genel işçilik |
| 2026/1043 | Manrel Mühendislik | 38 | PSV/TRV test ve Bakım-A, alev tutucu, patlama önleme vanası |

## Çalıştırma
```bash
python3 skills/birim_fiyat_havuzu.py
```
Çıktı: `/opt/asistan/ciktilar/YYYY-MM-DD_birim_fiyat_havuzu.xlsx`

## Sayfa yapısı
- **ÖZET**: 4 sözleşme listesi ve uyarı notu
- **Özgül Teknik / Dönmez Kumlama / NOT İnşaat / Manrel**: Poz No + Kategori + Açıklama + Birim + Birim Fiyat + Not

## Önemli
- `DOĞRULANMALI` yazanlar: PDF'den sayısal değer okunamayan pozlar — sözleşme orijinalinden teyit edin.
- Tüm fiyatlar KDV hariçtir.
- Yeni sözleşme eklemek için ilgili veri listesini (`*_DATA`) genişletin ve `sayfa_olustur()` çağrısı ekleyin.
