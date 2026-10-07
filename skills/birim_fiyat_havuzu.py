"""
Birim Fiyat Havuzu — 4 Sözleşme Verisi
Çıktı: /opt/asistan/ciktilar/YYYY-MM-DD_birim_fiyat_havuzu.xlsx

Hücre renk kodu:
  Beyaz/mavi  = sözleşmeden aynen alınan fiyat
  TURUNCU     = tahmini fiyat (piyasa araştırması, DOĞRULANMALI)
  Sarı        = sözleşmede değer okunamadı, tahmin de yok
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os
from datetime import date

CIKTI = "/opt/asistan/ciktilar"
DOSYA = os.path.join(CIKTI, f"{date.today().strftime('%Y-%m-%d')}_birim_fiyat_havuzu.xlsx")

RENK = {
    "baslik":    "1F3864",
    "alt_baslik":"2E75B6",
    "satir_a":   "BDD7EE",   # açık mavi — çift satırlar
    "tahmin":    "FCB040",   # turuncu — piyasa tahmini
    "bilinmiyor":"FFF2CC",   # sarı — hiç değer yok
    "beyaz":     "FFFFFF",
}

# ── Yardımcılar ───────────────────────────────────────────────────────────────
def kenar():
    s = Side(style="thin")
    return Border(left=s, right=s, top=s, bottom=s)

def bh(cell, bg, fg="FFFFFF", boyut=10, kalin=True):
    cell.font = Font(name="Calibri", bold=kalin, size=boyut, color=fg)
    cell.fill = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = kenar()

def bv(cell, bg=None, sayi=False, merkez=False):
    cell.font = Font(name="Calibri", size=9)
    if bg:
        cell.fill = PatternFill("solid", fgColor=bg)
    cell.border = kenar()
    cell.alignment = Alignment(vertical="center", horizontal="center" if merkez else "left", wrap_text=True)
    if sayi:
        cell.number_format = '#,##0.00'

def sayfa_olustur(wb, baslik, sozlesme_no, firma, sutunlar, satirlar):
    """
    satirlar = list of tuples: (poz_no, kategori, aciklama, birim, fiyat, not_str, tahmin_mi)
    tahmin_mi=True  → turuncu hücre
    tahmin_mi=False → normal satır rengi
    fiyat=None      → sarı (bilinmiyor)
    """
    ws = wb.create_sheet(baslik)
    ws.sheet_view.showGridLines = False

    ws.merge_cells(f"A1:{get_column_letter(len(sutunlar))}1")
    ws["A1"] = f"POLİSAN KİMYA — BİRİM FİYAT HAVUZU | {firma} | Sözleşme: {sozlesme_no}"
    bh(ws["A1"], RENK["baslik"], boyut=11)
    ws.row_dimensions[1].height = 24

    ws.merge_cells(f"A2:{get_column_letter(len(sutunlar))}2")
    ws["A2"] = (f"Hazırlanma: {date.today().strftime('%d.%m.%Y')}  |  KDV Hariç  |  "
                "🟠 Turuncu = piyasa tahmini (DOĞRULANMALI)  |  Sarı = değer girilmedi")
    ws["A2"].font = Font(name="Calibri", size=9, italic=True, color="555555")
    ws["A2"].alignment = Alignment(horizontal="center")
    ws.row_dimensions[2].height = 14
    ws.row_dimensions[3].height = 6

    for ci, (ad, genislik) in enumerate(sutunlar, 1):
        ws.column_dimensions[get_column_letter(ci)].width = genislik
        h = ws.cell(row=4, column=ci, value=ad)
        bh(h, RENK["alt_baslik"], boyut=9)
    ws.row_dimensions[4].height = 22

    fiyat_sutun = len(sutunlar) - 1  # Birim Fiyat sütunu sırası (son-1)

    for ri, satir in enumerate(satirlar, 5):
        *meta, tahmin_mi = satir          # son eleman = bool
        fiyat = meta[4]
        ws.row_dimensions[ri].height = 16

        if fiyat is None:
            satir_renk = RENK["bilinmiyor"]
        elif tahmin_mi:
            satir_renk = RENK["tahmin"]
        else:
            satir_renk = RENK["satir_a"] if ri % 2 == 0 else None

        for ci, deger in enumerate(meta, 1):
            c = ws.cell(row=ri, column=ci, value=deger)
            sayi_mi = (ci == fiyat_sutun and isinstance(deger, (int, float)))
            bv(c, satir_renk, sayi=sayi_mi, merkez=(ci in [1, 2, len(sutunlar)-1, len(sutunlar)]))

            # Fiyat hücresi: tahminse turuncu kalın yazı
            if ci == fiyat_sutun and tahmin_mi and deger is not None:
                c.font = Font(name="Calibri", size=9, bold=True, color="7F3F00")

    ws.auto_filter.ref = f"A4:{get_column_letter(len(sutunlar))}{4 + len(satirlar)}"
    ws.freeze_panes = "A5"
    return ws


# ── VERİ — sütun düzeni: Poz No, Kategori, Açıklama, Birim, Birim Fiyat, Not, tahmin_mi ──

OZGUL_SUTUNLAR = [
    ("Poz No",14),("Kategori",18),("Açıklama",52),
    ("Birim",10),("Birim Fiyat (₺)",18),("Not / Kaynak",34),
]
# tahmin_mi=False → sözleşmeden, True → piyasa tahmini
OZGUL_DATA = [
    # YEVMİYE
    ("PH.YEV.001","YEVMİYE","Malzemesiz Hafta İçi Tam Gün","Gün",9500,"",False),
    ("PH.YEV.002","YEVMİYE","Malzemesiz Hafta İçi Yarım Gün","Gün",5700,"",False),
    ("PH.YEV.003","YEVMİYE","Malzemesiz Hafta Sonu Tam Gün","Gün",14250,"",False),
    ("PH.YEV.004","YEVMİYE","Malzemesiz Hafta Sonu Yarım Gün","Gün",8550,"",False),
    ("PH.YEV.005","YEVMİYE","Malzemesiz Resmi Tatil Tam Gün","Gün",19000,"",False),
    ("PH.YEV.006","YEVMİYE","Malzemesiz Resmi Tatil Yarım Gün","Gün",11400,"",False),
    ("PH.YEV.007","YEVMİYE","Malzemeli Hafta İçi Tam Gün","Gün",16800,"",False),
    ("PH.YEV.008","YEVMİYE","Malzemeli Hafta İçi Yarım Gün","Gün",10080,"",False),
    ("PH.YEV.009","YEVMİYE","Malzemeli Hafta Sonu Tam Gün","Gün",25200,"",False),
    ("PH.YEV.010","YEVMİYE","Malzemeli Hafta Sonu Yarım Gün","Gün",15120,"",False),
    ("PH.YEV.011","YEVMİYE","Malzemeli Resmi Tatil Tam Gün","Gün",25200,"",False),
    ("PH.YEV.012","YEVMİYE","Malzemeli Resmi Tatil Yarım Gün","Gün",15120,"",False),
    ("PH.YEV.013","YEVMİYE","Formen Hafta İçi Tam Gün","Gün",11400,"",False),
    ("PH.YEV.014","YEVMİYE","Formen Hafta İçi Yarım Gün","Gün",6840,"",False),
    ("PH.YEV.015","YEVMİYE","Formen Hafta Sonu Tam Gün","Gün",17100,"",False),
    ("PH.YEV.016","YEVMİYE","Formen Hafta Sonu Yarım Gün","Gün",10260,"",False),
    ("PH.YEV.017","YEVMİYE","Formen Resmi Tatil Tam Gün","Gün",22800,"",False),
    ("PH.YEV.018","YEVMİYE","Formen Resmi Tatil Yarım Gün","Gün",13680,"",False),
    # MEKANİK
    ("PH.MEK.001","MEKANİK","CS Boru Fittings Montaj DN15","Adet",420,"",False),
    ("PH.MEK.002","MEKANİK","CS Boru Fittings Montaj DN20","Adet",420,"",False),
    ("PH.MEK.003","MEKANİK","CS Boru Fittings Montaj DN25","Adet",490,"",False),
    ("PH.MEK.004","MEKANİK","CS Boru Fittings Montaj DN32","Adet",490,"",False),
    ("PH.MEK.005","MEKANİK","CS Boru Fittings Montaj DN40","Adet",560,"",False),
    ("PH.MEK.006","MEKANİK","CS Boru Fittings Montaj DN50","Adet",630,"",False),
    ("PH.MEK.007","MEKANİK","CS Boru Fittings Montaj DN65","Adet",700,"",False),
    ("PH.MEK.008","MEKANİK","CS Boru Fittings Montaj DN80","Adet",840,"",False),
    ("PH.MEK.009","MEKANİK","CS Boru Fittings Montaj DN100","Adet",980,"",False),
    ("PH.MEK.010","MEKANİK","CS Boru Fittings Montaj DN125","Adet",1260,"",False),
    ("PH.MEK.011","MEKANİK","CS Boru Fittings Montaj DN150","Adet",1540,"",False),
    ("PH.MEK.012","MEKANİK","CS Boru Fittings Montaj DN200","Adet",2100,"",False),
    ("PH.MEK.013","MEKANİK","CS Boru Fittings Montaj DN250","Adet",2800,"",False),
    ("PH.MEK.014","MEKANİK","CS Boru Fittings Montaj DN300","Adet",3360,"",False),
    ("PH.MEK.015","MEKANİK","CS Boru Fittings Montaj DN350","Adet",3920,"",False),
    ("PH.MEK.016","MEKANİK","CS Boru Fittings Demontaj DN15","Adet",280,"",False),
    ("PH.MEK.017","MEKANİK","CS Boru Fittings Demontaj DN20","Adet",280,"",False),
    ("PH.MEK.018","MEKANİK","CS Boru Fittings Demontaj DN25","Adet",315,"",False),
    ("PH.MEK.019","MEKANİK","CS Boru Fittings Demontaj DN32","Adet",315,"",False),
    ("PH.MEK.020","MEKANİK","CS Boru Fittings Demontaj DN40","Adet",350,"",False),
    ("PH.MEK.021","MEKANİK","CS Boru Fittings Demontaj DN50","Adet",385,"",False),
    ("PH.MEK.022","MEKANİK","CS Boru Fittings Demontaj DN65","Adet",420,"",False),
    ("PH.MEK.023","MEKANİK","CS Boru Fittings Demontaj DN80","Adet",490,"",False),
    ("PH.MEK.024","MEKANİK","CS Boru Fittings Demontaj DN100","Adet",560,"",False),
    ("PH.MEK.025","MEKANİK","CS Boru Fittings Demontaj DN150","Adet",840,"",False),
    ("PH.MEK.026","MEKANİK","CS Boru Fittings Demontaj DN200","Adet",1120,"",False),
    ("PH.MEK.027","MEKANİK","CS Boru Fittings Demontaj DN250","Adet",1400,"",False),
    ("PH.MEK.028","MEKANİK","CS Boru Fittings Demontaj DN300","Adet",1680,"",False),
    ("PH.MEK.029","MEKANİK","CS Boru Kaynak DN15-DN50","Adet",700,"",False),
    ("PH.MEK.030","MEKANİK","CS Boru Kaynak DN65-DN100","Adet",1400,"",False),
    ("PH.MEK.031","MEKANİK","CS Boru Kaynak DN150","Adet",2100,"",False),
    ("PH.MEK.032","MEKANİK","CS Boru Kaynak DN200","Adet",2800,"",False),
    ("PH.MEK.033","MEKANİK","CS Boru Kaynak DN250","Adet",3500,"",False),
    ("PH.MEK.034","MEKANİK","CS Boru Kaynak DN300","Adet",4200,"",False),
    ("PH.MEK.035","MEKANİK","CS Boru Kaynak DN350","Adet",4900,"",False),
    ("PH.MEK.036","MEKANİK","SS Boru Fittings Montaj DN15","Adet",630,"",False),
    ("PH.MEK.037","MEKANİK","SS Boru Fittings Montaj DN25","Adet",735,"",False),
    ("PH.MEK.038","MEKANİK","SS Boru Fittings Montaj DN50","Adet",945,"",False),
    ("PH.MEK.039","MEKANİK","SS Boru Fittings Montaj DN80","Adet",1260,"",False),
    ("PH.MEK.040","MEKANİK","SS Boru Fittings Montaj DN100","Adet",1470,"",False),
    ("PH.MEK.041","MEKANİK","SS Boru Fittings Montaj DN150","Adet",2310,"",False),
    ("PH.MEK.042","MEKANİK","SS Boru Fittings Montaj DN200","Adet",3150,"",False),
    ("PH.MEK.043","MEKANİK","SS Boru Kaynak DN15-DN50","Adet",1050,"",False),
    ("PH.MEK.044","MEKANİK","SS Boru Kaynak DN65-DN100","Adet",2100,"",False),
    ("PH.MEK.045","MEKANİK","SS Boru Kaynak DN150","Adet",3150,"",False),
    ("PH.MEK.046","MEKANİK","SS Boru Kaynak DN200","Adet",4200,"",False),
    ("PH.MEK.047","MEKANİK","Galvaniz Boru Fittings Montaj DN15","Adet",490,"",False),
    ("PH.MEK.048","MEKANİK","Galvaniz Boru Fittings Montaj DN25","Adet",560,"",False),
    ("PH.MEK.049","MEKANİK","Galvaniz Boru Fittings Montaj DN50","Adet",700,"",False),
    ("PH.MEK.050","MEKANİK","Galvaniz Boru Fittings Montaj DN80","Adet",980,"",False),
    ("PH.MEK.051","MEKANİK","Vana Montaj CS DN15","Adet",350,"",False),
    ("PH.MEK.052","MEKANİK","Vana Montaj CS DN20","Adet",350,"",False),
    ("PH.MEK.053","MEKANİK","Vana Montaj CS DN25","Adet",420,"",False),
    ("PH.MEK.054","MEKANİK","Vana Montaj CS DN32","Adet",420,"",False),
    ("PH.MEK.055","MEKANİK","Vana Montaj CS DN40","Adet",490,"",False),
    ("PH.MEK.056","MEKANİK","Vana Montaj CS DN50","Adet",560,"",False),
    ("PH.MEK.057","MEKANİK","Vana Montaj CS DN65","Adet",630,"",False),
    ("PH.MEK.058","MEKANİK","Vana Montaj CS DN80","Adet",770,"",False),
    ("PH.MEK.059","MEKANİK","Vana Montaj CS DN100","Adet",910,"",False),
    ("PH.MEK.060","MEKANİK","Vana Montaj CS DN150","Adet",1400,"",False),
    ("PH.MEK.061","MEKANİK","Vana Montaj CS DN200","Adet",1960,"",False),
    ("PH.MEK.062","MEKANİK","Vana Montaj CS DN250","Adet",2520,"",False),
    ("PH.MEK.063","MEKANİK","Vana Montaj CS DN300","Adet",3080,"",False),
    ("PH.MEK.064","MEKANİK","Vana Demontaj CS DN15","Adet",210,"",False),
    ("PH.MEK.065","MEKANİK","Vana Demontaj CS DN50","Adet",315,"",False),
    ("PH.MEK.066","MEKANİK","Vana Demontaj CS DN100","Adet",525,"",False),
    ("PH.MEK.067","MEKANİK","Vana Demontaj CS DN200","Adet",1120,"",False),
    ("PH.MEK.068","MEKANİK","Vana Montaj SS DN15","Adet",525,"",False),
    ("PH.MEK.069","MEKANİK","Vana Montaj SS DN50","Adet",840,"",False),
    ("PH.MEK.070","MEKANİK","Vana Montaj SS DN100","Adet",1365,"",False),
    ("PH.MEK.071","MEKANİK","Vana Montaj SS DN200","Adet",2940,"",False),
    ("PH.MEK.072","MEKANİK","Kör Flanş CS DN15-DN50 Montaj","Adet",280,"",False),
    ("PH.MEK.073","MEKANİK","Kör Flanş CS DN65-DN100 Montaj","Adet",420,"",False),
    ("PH.MEK.074","MEKANİK","Kör Flanş CS DN150-DN200 Montaj","Adet",700,"",False),
    ("PH.MEK.075","MEKANİK","Deluge Vana Montaj 4\"","Adet",5600,"",False),
    ("PH.MEK.076","MEKANİK","Deluge Vana Montaj 6\"","Adet",7000,"",False),
    ("PH.MEK.077","MEKANİK","Pompa Montaj (genel, malzemesiz)","Adet",28000,"",False),
    ("PH.MEK.078","MEKANİK","Pompa Demontaj (genel)","Adet",16800,"",False),
    # İNŞAAT
    ("PH.INS.001","İNŞAAT","Betopan Levha Montaj","m²",402.58,"",False),
    ("PH.INS.002","İNŞAAT","Alçıpan Levha Montaj","m²",402.58,"",False),
    ("PH.INS.003","İNŞAAT","Trapez Sac Cephe/Çatı Kaplama","m²",644,"",False),
    ("PH.INS.004","İNŞAAT","Metal Kapı Montaj (≤1.5 m²)","Adet",3200,"~Piyasa 2026",True),
    ("PH.INS.005","İNŞAAT","Metal Kapı Montaj (>1.5 m²)","Adet",5500,"~Piyasa 2026",True),
    # ÇELİK
    ("PH.ÇLK.001","ÇELİK","Çelik Konstrüksiyon İmalat (malzeme hariç)","kg",140,"",False),
    ("PH.ÇLK.002","ÇELİK","Çelik Konstrüksiyon Montaj","kg",112,"",False),
    ("PH.ÇLK.003","ÇELİK","Çelik Konstrüksiyon Demontaj","kg",56,"",False),
    ("PH.ÇLK.004","ÇELİK","Galvaniz Petek/Izgara Döşeme","m²",70,"",False),
    # HDPE Montaj — piyasa tahmini (DSİ 2026 + sektör ortalaması)
    ("PH.PE.001","HDPE","HDPE100 Boru Montaj Ø20mm","m",400,"~Piyasa 2026",True),
    ("PH.PE.002","HDPE","HDPE100 Boru Montaj Ø25mm","m",430,"~Piyasa 2026",True),
    ("PH.PE.003","HDPE","HDPE100 Boru Montaj Ø32mm","m",480,"~Piyasa 2026",True),
    ("PH.PE.004","HDPE","HDPE100 Boru Montaj Ø40mm","m",540,"~Piyasa 2026",True),
    ("PH.PE.005","HDPE","HDPE100 Boru Montaj Ø50mm","m",620,"~Piyasa 2026",True),
    ("PH.PE.006","HDPE","HDPE100 Boru Montaj Ø63mm","m",720,"~Piyasa 2026",True),
    ("PH.PE.007","HDPE","HDPE100 Boru Montaj Ø75mm","m",900,"~Piyasa 2026",True),
    ("PH.PE.008","HDPE","HDPE100 Boru Montaj Ø90mm","m",1100,"~Piyasa 2026",True),
    ("PH.PE.009","HDPE","HDPE100 Boru Montaj Ø110mm","m",1350,"~Piyasa 2026",True),
    ("PH.PE.010","HDPE","HDPE100 Boru Montaj Ø125mm","m",1600,"~Piyasa 2026",True),
    ("PH.PE.011","HDPE","HDPE100 Boru Montaj Ø160mm","m",2100,"~Piyasa 2026",True),
    ("PH.PE.012","HDPE","HDPE100 Boru Montaj Ø200mm","m",2800,"~Piyasa 2026",True),
    ("PH.PE.013","HDPE","HDPE100 Boru Montaj Ø250mm","m",3800,"~Piyasa 2026",True),
    ("PH.PE.014","HDPE","HDPE100 Boru Montaj Ø315mm","m",5200,"~Piyasa 2026",True),
    ("PH.PE.015","HDPE","HDPE100 Boru Montaj Ø400mm","m",7000,"~Piyasa 2026",True),
    ("PH.PE.016","HDPE","HDPE100 Boru Montaj Ø500mm","m",9500,"~Piyasa 2026",True),
    ("PH.PE.017","HDPE","HDPE100 Boru Montaj Ø630mm","m",13000,"~Piyasa 2026",True),
    ("PH.PE.018","HDPE","HDPE100 Boru Montaj Ø710mm","m",16500,"~Piyasa 2026",True),
    # HDPE Demontaj (~%40 montaj)
    ("PH.PE.019","HDPE","HDPE100 Boru Demontaj Ø20-63mm","m",200,"~Piyasa 2026",True),
    ("PH.PE.020","HDPE","HDPE100 Boru Demontaj Ø75-125mm","m",500,"~Piyasa 2026",True),
    ("PH.PE.021","HDPE","HDPE100 Boru Demontaj Ø160-200mm","m",1100,"~Piyasa 2026",True),
    # HDPE Kaynak
    ("PH.PE.022","HDPE","HDPE100 Boru Alın Kaynak Ø20-63mm","Adet",650,"~Piyasa 2026",True),
    ("PH.PE.023","HDPE","HDPE100 Boru Alın Kaynak Ø75-125mm","Adet",1350,"~Piyasa 2026",True),
    ("PH.PE.024","HDPE","HDPE100 Boru Alın Kaynak Ø160-200mm","Adet",2500,"~Piyasa 2026",True),
    ("PH.PE.025","HDPE","HDPE100 Elektrofüzyon Kaynak Ø20-63mm","Adet",500,"~Piyasa 2026",True),
    ("PH.PE.026","HDPE","HDPE100 Elektrofüzyon Kaynak Ø75-125mm","Adet",1100,"~Piyasa 2026",True),
    ("PH.PE.027","HDPE","HDPE100 Elektrofüzyon Kaynak Ø160-200mm","Adet",2100,"~Piyasa 2026",True),
    # HDPE Vana
    ("PH.PE.028","HDPE","HDPE100 Vana Montaj Ø25-63mm","Adet",900,"~Piyasa 2026",True),
    ("PH.PE.029","HDPE","HDPE100 Vana Montaj Ø75-125mm","Adet",2000,"~Piyasa 2026",True),
    ("PH.PE.030","HDPE","HDPE100 Vana Montaj Ø160-200mm","Adet",3800,"~Piyasa 2026",True),
]

DONMEZ_SUTUNLAR = [
    ("Poz No",14),("Kategori",20),("Açıklama",52),
    ("Birim",10),("Birim Fiyat (₺)",18),("Not / Kaynak",34),
]
DONMEZ_DATA = [
    ("PPBF-100-001","KUMLAMA","Yürüme Yolu Çizimi","m",28.30,"",False),
    ("PPBF-100-002","KUMLAMA","Yüzey Kumlama — Bazalt","m²",840.25,"",False),
    ("PPBF-100-003","KUMLAMA","Yüzey Kumlama — Cüruf","m²",1235.25,"",False),
    ("PPBF-100-004","KUMLAMA","Yüzey Kumlama — Grit (Çelik bilye)","m²",1770.25,"",False),
    ("PPBF-100-005","İSKELE","İskele Kurulum/Söküm","m³",685.50,"",False),
    ("PPBF-100-006","İSKELE","Küçük İskele Pozu","m³",685.50,"",False),
    ("PPBF-100-007","YEVMİYE","Yevmiye (Gün)","Gün",4200,"",False),
    ("PPBF-100-008","YEVMİYE","Yevmiye (Saat)","Saat",625,"",False),
    ("PPBF-100-009","KİMYASAL","Kimyasal Temizlik","m²",869.35,"",False),
    ("PPBF-100-010","ŞABLON","Tank Şablon İşlemi","Adet",45925.30,"",False),
    # Boyama kalemleri — piyasa tahmini (dek-mar.com.tr, 2026)
    ("PPBF-100-011","BOYAMA","Boya — 1 Kat (genel, malzeme dahil)","m²",220,"~Piyasa 2026",True),
    ("PPBF-100-012","BOYAMA","Boya — 2 Kat","m²",380,"~Piyasa 2026",True),
    ("PPBF-100-013","BOYAMA","Boya — 3 Kat","m²",540,"~Piyasa 2026",True),
    ("PPBF-100-014","BOYAMA","Boya — Zemin/Epoksi 1 Kat","m²",450,"~Piyasa 2026",True),
    ("PPBF-100-015","BOYAMA","Boya — Zemin/Epoksi 2 Kat","m²",780,"~Piyasa 2026",True),
    ("PPBF-100-016","BOYAMA","Boya Demaj/Sıyırma","m²",200,"~Piyasa 2026",True),
    ("PPBF-100-017","BOYAMA","Çatı Boya 1 Kat","m²",240,"~Piyasa 2026",True),
    ("PPBF-100-018","BOYAMA","Çatı Boya 2 Kat","m²",420,"~Piyasa 2026",True),
    ("PPBF-100-019","KORUMA","Pas Önleyici Astar 1 Kat","m²",175,"~Piyasa 2026",True),
    ("PPBF-100-020","KORUMA","Pas Önleyici Astar 2 Kat","m²",310,"~Piyasa 2026",True),
    ("PPBF-100-021","KORUMA","Epoksi Astar 1 Kat","m²",240,"~Piyasa 2026",True),
    ("PPBF-100-022","KORUMA","Poliüretan Topkat 1 Kat","m²",340,"~Piyasa 2026",True),
    ("PPBF-100-023","KORUMA","Poliüretan Topkat 2 Kat","m²",600,"~Piyasa 2026",True),
    ("PPBF-100-024","KORUMA","Yüksek Sıcaklık Boya 1 Kat (>200°C)","m²",420,"~Piyasa 2026",True),
    ("PPBF-100-025","KORUMA","Yüksek Sıcaklık Boya 2 Kat","m²",760,"~Piyasa 2026",True),
    ("PPBF-100-026","KORUMA","Galvaniz Sonrası Boya","m²",220,"~Piyasa 2026",True),
    ("PPBF-100-027","MALZEME","Rulo Boya Uygulaması (işçilik)","m²",150,"~Piyasa 2026",True),
    ("PPBF-100-028","MALZEME","Airless Püskürtme (işçilik)","m²",130,"~Piyasa 2026",True),
    ("PPBF-100-029","MALZEME","Fırça Uygulaması (işçilik)","m²",160,"~Piyasa 2026",True),
    ("PPBF-100-030","TAMİR","Boya Tamir/Rötuş","m²",420,"~Piyasa 2026",True),
    ("PPBF-100-031","TAMİR","Kaynak Yeri Boya Tamir","m",300,"~Piyasa 2026",True),
    ("PPBF-100-032","NAKLİYE","Ekipman Nakliye (saha içi)","Sefer",3000,"~Piyasa 2026",True),
    ("PPBF-100-033","DİĞER","Renk Kodu Yazma/Etiketleme","Adet",600,"~Piyasa 2026",True),
    ("PPBF-100-034","DİĞER","İşaret/Uyarı Levha Boya","Adet",400,"~Piyasa 2026",True),
    ("PPBF-100-035","DİĞER","Zemin Çizgi Boyama","m",120,"~Piyasa 2026",True),
    ("PPBF-100-036","DİĞER","Zemin Sembol Boyama","Adet",800,"~Piyasa 2026",True),
    ("PPBF-100-037","DİĞER","Flanş/Bağlantı Koruma Kapağı Montaj","Adet",200,"~Piyasa 2026",True),
    ("PPBF-100-038","DİĞER","Paket İş (teklif bazlı)","Adet",None,"Teklif alınmalı",False),
]

NOT_SUTUNLAR = [
    ("Poz No",12),("Kategori",20),("Açıklama",52),
    ("Birim",10),("Birim Fiyat (₺)",18),("Not / Kaynak",34),
]
NOT_DATA = [
    ("PPRC.001","BORU","PPRC Boru Montaj DN20 (3/4\")","m",1600,"",False),
    ("PPRC.002","BORU","PPRC Boru Montaj DN25 (1\")","m",1800,"",False),
    ("PPRC.003","BORU","PPRC Boru Montaj DN32 (1¼\")","m",2100,"",False),
    ("PPRC.004","BORU","PPRC Boru Montaj DN40 (1½\")","m",2350,"",False),
    ("PPRC.005","BORU","PPRC Boru Montaj DN50 (2\")","m",2350,"",False),
    ("PPRC.006","BORU","PPRC Boru Montaj DN63 (2½\")","m",2350,"",False),
    ("PVC.001","BORU","PVC/Koruge Boru Montaj DN75","m",1900,"",False),
    ("PVC.002","BORU","PVC/Koruge Boru Montaj DN100","m",2100,"",False),
    ("PVC.003","BORU","PVC/Koruge Boru Montaj DN150","m",2800,"",False),
    ("PVC.004","BORU","PVC/Koruge Boru Montaj DN200","m",3200,"",False),
    ("PVC.005","BORU","PVC/Koruge Boru Montaj DN300","m",3800,"",False),
    # Armatür — piyasa tahmini
    ("ARM.001","TESİSAT","Armatür — Batarya Değişimi","Adet",1000,"~Piyasa 2026",True),
    ("ARM.002","TESİSAT","Armatür — Lavabo Montaj","Adet",750,"~Piyasa 2026",True),
    ("ARM.003","TESİSAT","Armatür — Klozet Montaj","Adet",850,"~Piyasa 2026",True),
    ("ARM.004","TESİSAT","Rezervuar Değişimi","Adet",600,"~Piyasa 2026",True),
    ("ARM.005","TESİSAT","Duş Kabini Montaj","Adet",1600,"~Piyasa 2026",True),
    # Kaplama — piyasa tahmini (ankaradekorasyontadilat.com, 2026)
    ("KPD.001","KAPLAMA","Seramik Döşeme (≤600x600mm) — işçilik","m²",650,"~Piyasa 2026",True),
    ("KPD.002","KAPLAMA","Seramik Duvar Kaplaması — işçilik","m²",700,"~Piyasa 2026",True),
    ("KPD.003","KAPLAMA","Granit Döşeme — işçilik","m²",900,"~Piyasa 2026",True),
    ("KPD.004","KAPLAMA","Epoksi Zemin Kaplama (self-leveling)","m²",500,"~Piyasa 2026",True),
    # Asma tavan — piyasa tahmini
    ("CAT.001","ASMA TAVAN","Asma Tavan — Alçıpan","m²",480,"~Piyasa 2026",True),
    ("CAT.002","ASMA TAVAN","Asma Tavan — Metal Karkas","m²",380,"~Piyasa 2026",True),
    # Beton/Kazı
    ("BTN.001","BETON/KAZI","El Kazısı","m³",250,"",False),
    ("BTN.002","BETON/KAZI","Moloz Nakli","m³",1000,"",False),
    ("BTN.003","BETON/KAZI","Çelik Hasır (düz)","kg",95,"",False),
    ("BTN.004","BETON/KAZI","Demir Donatı (nervürlü) — işçilik","kg",15,"~Piyasa 2026",True),
    ("BTN.005","BETON/KAZI","C20 Beton Dökme (malzeme+işçilik)","m³",3500,"~Piyasa 2026",True),
    ("BTN.006","BETON/KAZI","C25 Beton Dökme (malzeme+işçilik)","m³",4000,"~Piyasa 2026",True),
    ("BTN.007","BETON/KAZI","Kalıp İşçiliği","m²",280,"~Piyasa 2026",True),
    # Genel
    ("MF.001","GENEL","Yatay Taşıma (saha içi, el arabası)","Adet",150,"",False),
    ("MF.002","GENEL","Genel Temizlik/İşçilik","Saat",1350,"",False),
    ("MF.003","GENEL","Yükleme-Boşaltma","Saat",1350,"",False),
    ("MF.004","GENEL","Pencere/Kapı Silikon Uygulama","m",150,"~Piyasa 2026",True),
    ("MF.005","GENEL","Panel Kapı Montaj","Adet",2200,"~Piyasa 2026",True),
    ("MF.006","GENEL","PVC Pencere Montaj","m²",1600,"~Piyasa 2026",True),
]

MANREL_SUTUNLAR = [
    ("Poz No",14),("Tip",18),("Açıklama",48),
    ("Boyut",12),("Birim Fiyat (₺)",18),("Not / Kaynak",28),
]
MANREL_DATA = [
    ("MNR.PSV.001","PSV/TRV TEST","Emniyet/Tahliye Vanası Test","1/2\"",1000,"",False),
    ("MNR.PSV.002","PSV/TRV TEST","Emniyet/Tahliye Vanası Test","3/4\"",1000,"",False),
    ("MNR.PSV.003","PSV/TRV TEST","Emniyet/Tahliye Vanası Test","1\"",1500,"",False),
    ("MNR.PSV.004","PSV/TRV TEST","Emniyet/Tahliye Vanası Test","1½\"",1500,"",False),
    ("MNR.PSV.005","PSV/TRV TEST","Emniyet/Tahliye Vanası Test","2\"",2000,"",False),
    ("MNR.PSV.006","PSV/TRV TEST","Emniyet/Tahliye Vanası Test","3\"",3000,"",False),
    ("MNR.PSV.007","PSV/TRV TEST","Emniyet/Tahliye Vanası Test","4\"",4000,"",False),
    ("MNR.PSV.008","PSV/TRV TEST","Emniyet/Tahliye Vanası Test","6\"",6000,"",False),
    ("MNR.PSV.009","PSV/TRV TEST","Emniyet/Tahliye Vanası Test","8\"",8000,"",False),
    ("MNR.PSV.010","PSV/TRV TEST","Emniyet/Tahliye Vanası Test","10\"",10000,"",False),
    ("MNR.PSV.011","PSV/TRV TEST","Emniyet/Tahliye Vanası Test","12\"",16000,"",False),
    ("MNR.BAK.001","PSV/TRV BAKIM-A","PSV/TRV Bakım-A (söküm+revizyon+test)","1/2\"",3250,"",False),
    ("MNR.BAK.002","PSV/TRV BAKIM-A","PSV/TRV Bakım-A","3/4\"",3250,"",False),
    ("MNR.BAK.003","PSV/TRV BAKIM-A","PSV/TRV Bakım-A","1\"",4750,"",False),
    ("MNR.BAK.004","PSV/TRV BAKIM-A","PSV/TRV Bakım-A","1½\"",4750,"",False),
    ("MNR.BAK.005","PSV/TRV BAKIM-A","PSV/TRV Bakım-A","2\"",6300,"",False),
    ("MNR.BAK.006","PSV/TRV BAKIM-A","PSV/TRV Bakım-A","3\"",9350,"",False),
    ("MNR.BAK.007","PSV/TRV BAKIM-A","PSV/TRV Bakım-A","4\"",12400,"",False),
    ("MNR.BAK.008","PSV/TRV BAKIM-A","PSV/TRV Bakım-A","6\"",19100,"",False),
    ("MNR.BAK.009","PSV/TRV BAKIM-A","PSV/TRV Bakım-A","8\"",25700,"",False),
    ("MNR.BAK.010","PSV/TRV BAKIM-A","PSV/TRV Bakım-A","10\"",32500,"",False),
    ("MNR.BAK.011","PSV/TRV BAKIM-A","PSV/TRV Bakım-A","12\"",41300,"",False),
    ("MNR.PVV.001","PVV TEST","Patlama Önleme Vanası Test","4\"",3350,"",False),
    ("MNR.PVV.002","PVV TEST","Patlama Önleme Vanası Test","6\"",3350,"",False),
    ("MNR.PVV.003","PVV TEST","Patlama Önleme Vanası Test","8\"",3350,"",False),
    ("MNR.PVV.004","PVV TEST","Patlama Önleme Vanası Test","10\"",3350,"",False),
    ("MNR.PVV.005","PVV TEST","Patlama Önleme Vanası Test","12\"",3350,"",False),
    ("MNR.PVV.006","PVV TEST","Patlama Önleme Vanası Test","16\"",3350,"",False),
    ("MNR.PVV.007","PVV TEST","Patlama Önleme Vanası Test","20\"",3350,"",False),
    ("MNR.PVV.008","PVV TEST","Patlama Önleme Vanası Test","24\"",3350,"",False),
    ("MNR.AT.001","ALEV TUTUCU BAKIM-A","Alev Tutucu Bakım-A","4\"",8300,"",False),
    ("MNR.AT.002","ALEV TUTUCU BAKIM-A","Alev Tutucu Bakım-A","6\"",8300,"",False),
    ("MNR.AT.003","ALEV TUTUCU BAKIM-A","Alev Tutucu Bakım-A","8\"",10400,"",False),
    ("MNR.AT.004","ALEV TUTUCU BAKIM-A","Alev Tutucu Bakım-A","10\"",12550,"",False),
    ("MNR.AT.005","ALEV TUTUCU BAKIM-A","Alev Tutucu Bakım-A","12\"",14650,"",False),
    ("MNR.AT.006","ALEV TUTUCU BAKIM-A","Alev Tutucu Bakım-A","16\"",16750,"",False),
    ("MNR.AT.007","ALEV TUTUCU BAKIM-A","Alev Tutucu Bakım-A","20\"",18500,"",False),
    ("MNR.AT.008","ALEV TUTUCU BAKIM-A","Alev Tutucu Bakım-A","24\"",18500,"",False),
]

# ── EXCEL YARAT ───────────────────────────────────────────────────────────────
os.makedirs(CIKTI, exist_ok=True)
wb = openpyxl.Workbook()

# ÖZET sayfası
ws_ozet = wb.active
ws_ozet.title = "ÖZET"
ws_ozet.sheet_view.showGridLines = False
for harf, gen in [("A",36),("B",20),("C",22),("D",16),("E",16)]:
    ws_ozet.column_dimensions[harf].width = gen

ws_ozet.merge_cells("A1:E1")
ws_ozet["A1"] = "POLİSAN KİMYA — BİRİM FİYAT HAVUZU 2026"
bh(ws_ozet["A1"], RENK["baslik"], boyut=14)
ws_ozet.row_dimensions[1].height = 30

ws_ozet.merge_cells("A2:E2")
ws_ozet["A2"] = (f"Hazırlanma: {date.today().strftime('%d.%m.%Y')}  |  4 Sözleşme  |  KDV Hariç  |  "
                 "🟠 Turuncu = piyasa tahmini — sözleşmeden doğrulayın")
ws_ozet["A2"].font = Font(name="Calibri", size=9, italic=True, color="555555")
ws_ozet["A2"].alignment = Alignment(horizontal="center")
ws_ozet.row_dimensions[2].height = 14
ws_ozet.row_dimensions[3].height = 8

for ci, metin in enumerate(["Sözleşme / Firma","Sözleşme No","Sözleşme Türü","Poz Sayısı","Sayfa"], 1):
    bh(ws_ozet.cell(row=4, column=ci, value=metin), RENK["alt_baslik"], boyut=10)
ws_ozet.row_dimensions[4].height = 22

for ri, (firma, szl_no, tur, poz, sayfa) in enumerate([
    ("Özgül Teknik Metal İnş.","2026/1013","Birim Fiyat",len(OZGUL_DATA),"Özgül Teknik"),
    ("Dönmez Kumlama Boya","2026/1015","Birim Fiyat",len(DONMEZ_DATA),"Dönmez Kumlama"),
    ("NOT İnşaat","2026/1023","Birim Fiyat",len(NOT_DATA),"NOT İnşaat"),
    ("Manrel Mühendislik","2026/1043","Birim Fiyat",len(MANREL_DATA),"Manrel"),
], 5):
    renk = RENK["satir_a"] if ri % 2 == 0 else None
    ws_ozet.row_dimensions[ri].height = 18
    for ci, deger in enumerate([firma, szl_no, tur, poz, sayfa], 1):
        bv(ws_ozet.cell(row=ri, column=ci, value=deger), renk, merkez=(ci >= 3))

ws_ozet.row_dimensions[9].height = 10

uyari = ws_ozet["A10"]
ws_ozet.merge_cells("A10:E10")
uyari.value = ("🟠 Turuncu satırlar piyasa araştırmasına dayalı tahmindir (Kaynak: dek-mar.com.tr, DSİ 2026, "
               "ankaradekorasyontadilat.com). Fatura geldiğinde sözleşme orijinalinden teyit edin.")
uyari.font = Font(name="Calibri", size=9, color="7F3F00", bold=True)
uyari.fill = PatternFill("solid", fgColor="FCB040")
uyari.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
uyari.border = kenar()
ws_ozet.row_dimensions[10].height = 32

# Sözleşme sayfaları
sayfa_olustur(wb,"Özgül Teknik","2026/1013","Özgül Teknik Metal İnş.",OZGUL_SUTUNLAR,OZGUL_DATA)
sayfa_olustur(wb,"Dönmez Kumlama","2026/1015","Dönmez Kumlama Boya İnş.",DONMEZ_SUTUNLAR,DONMEZ_DATA)
sayfa_olustur(wb,"NOT İnşaat","2026/1023","NOT İnşaat",NOT_SUTUNLAR,NOT_DATA)
sayfa_olustur(wb,"Manrel","2026/1043","Manrel Mühendislik",MANREL_SUTUNLAR,MANREL_DATA)

wb.save(DOSYA)

ozgul_t  = sum(1 for r in OZGUL_DATA  if r[-1])
donmez_t = sum(1 for r in DONMEZ_DATA if r[-1])
not_t    = sum(1 for r in NOT_DATA    if r[-1])
print(f"Dosya: {DOSYA}")
print(f"Turuncu (tahmin): Özgül={ozgul_t}, Dönmez={donmez_t}, NOT={not_t}, Manrel=0")
print(f"Toplam poz: {len(OZGUL_DATA)+len(DONMEZ_DATA)+len(NOT_DATA)+len(MANREL_DATA)}")
