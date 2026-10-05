# Pompa bakım maliyet tablosu

`pompa_bakim_maliyet.py`, beş kalemlik formüllü bir Excel bakım maliyet tablosu üretir. Miktar ve birim fiyat alanlarını boş bırakır; doğrulanmamış maliyet varsayımı yapmaz. Tutarları ve eksik bilgi kontrolünü Excel formülleriyle hesaplar.

Çalıştırma:

```bash
/opt/asistan/venv/bin/python /opt/asistan/skills/pompa_bakim_maliyet.py
```

Farklı dosya yolu için `--output /tam/yol/dosya.xlsx` kullanın. Varsayılan çıktı, `/opt/asistan/ciktilar` altında tarihli bir Excel dosyasıdır.
