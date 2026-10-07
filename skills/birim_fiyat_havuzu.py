"""
Birim Fiyat Havuzu — 4 Sözleşme Verisi
Çıktı: /opt/asistan/ciktilar/2026-10-07_birim_fiyat_havuzu.xlsx
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
    "vurgu":     "BDD7EE",
    "tamam":     "E2EFDA",
    "uyari":     "FFF2CC",
    "beyaz":     "FFFFFF",
}

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
    ws = wb.create_sheet(baslik)
    ws.sheet_view.showGridLines = False

    # Başlık
    ws.merge_cells(f"A1:{get_column_letter(len(sutunlar))}1")
    ws["A1"] = f"POLİSAN KİMYA — BİRİM FİYAT HAVUZU | {firma} | Sözleşme: {sozlesme_no}"
    bh(ws["A1"], RENK["baslik"], boyut=11)
    ws.row_dimensions[1].height = 24

    ws.merge_cells(f"A2:{get_column_letter(len(sutunlar))}2")
    ws["A2"] = f"Hazırlanma: {date.today().strftime('%d.%m.%Y')}  |  Fiyatlar KDV HARİÇ — Sözleşme tarihindeki fiyatlar"
    ws["A2"].font = Font(name="Calibri", size=9, italic=True, color="555555")
    ws["A2"].alignment = Alignment(horizontal="center")
    ws.row_dimensions[2].height = 14

    ws.row_dimensions[3].height = 6

    # Sütun başlıkları
    for ci, (ad, genislik) in enumerate(sutunlar, 1):
        ws.column_dimensions[get_column_letter(ci)].width = genislik
        h = ws.cell(row=4, column=ci, value=ad)
        bh(h, RENK["alt_baslik"], boyut=9)
    ws.row_dimensions[4].height = 22

    # Veri satırları
    for ri, satir in enumerate(satirlar, 5):
        renk = RENK["vurgu"] if ri % 2 == 0 else None
        ws.row_dimensions[ri].height = 16
        for ci, deger in enumerate(satir, 1):
            c = ws.cell(row=ri, column=ci, value=deger)
            sayi = isinstance(deger, (int, float)) and ci >= len(sutunlar) - 1
            bv(c, renk, sayi=(ci == len(sutunlar) - 1 and isinstance(deger, (int, float))),
               merkez=(ci in [1, 2, len(sutunlar) - 2, len(sutunlar)]))

    # Filtre
    ws.auto_filter.ref = f"A4:{get_column_letter(len(sutunlar))}{4 + len(satirlar)}"
    ws.freeze_panes = "A5"
    return ws


# ── VERİ BLOKLARI ─────────────────────────────────────────────────────────────

# Özgül Teknik sütunları: Poz No | Kategori | Açıklama | Birim | Birim Fiyat (₺) | Not
OZGUL_SUTUNLAR = [
    ("Poz No", 16), ("Kategori", 20), ("Açıklama", 55),
    ("Birim", 10), ("Birim Fiyat (₺)", 18), ("Not / Kaynak", 28),
]

OZGUL_DATA = [
    # YEVMİYE
    ("PH.YEV.001","YEVMİYE","Malzemesiz Hafta İçi Tam Gün","Gün",9500,""),
    ("PH.YEV.002","YEVMİYE","Malzemesiz Hafta İçi Yarım Gün","Gün",5700,""),
    ("PH.YEV.003","YEVMİYE","Malzemesiz Hafta Sonu Tam Gün","Gün",14250,""),
    ("PH.YEV.004","YEVMİYE","Malzemesiz Hafta Sonu Yarım Gün","Gün",8550,""),
    ("PH.YEV.005","YEVMİYE","Malzemesiz Resmi Tatil Tam Gün","Gün",19000,""),
    ("PH.YEV.006","YEVMİYE","Malzemesiz Resmi Tatil Yarım Gün","Gün",11400,""),
    ("PH.YEV.007","YEVMİYE","Malzemeli Hafta İçi Tam Gün","Gün",16800,""),
    ("PH.YEV.008","YEVMİYE","Malzemeli Hafta İçi Yarım Gün","Gün",10080,""),
    ("PH.YEV.009","YEVMİYE","Malzemeli Hafta Sonu Tam Gün","Gün",25200,""),
    ("PH.YEV.010","YEVMİYE","Malzemeli Hafta Sonu Yarım Gün","Gün",15120,""),
    ("PH.YEV.011","YEVMİYE","Malzemeli Resmi Tatil Tam Gün","Gün",25200,""),
    ("PH.YEV.012","YEVMİYE","Malzemeli Resmi Tatil Yarım Gün","Gün",15120,""),
    ("PH.YEV.013","YEVMİYE","Formen Hafta İçi Tam Gün","Gün",11400,""),
    ("PH.YEV.014","YEVMİYE","Formen Hafta İçi Yarım Gün","Gün",6840,""),
    ("PH.YEV.015","YEVMİYE","Formen Hafta Sonu Tam Gün","Gün",17100,""),
    ("PH.YEV.016","YEVMİYE","Formen Hafta Sonu Yarım Gün","Gün",10260,""),
    ("PH.YEV.017","YEVMİYE","Formen Resmi Tatil Tam Gün","Gün",22800,""),
    ("PH.YEV.018","YEVMİYE","Formen Resmi Tatil Yarım Gün","Gün",13680,""),
    # MEKANİK — seçilmiş önemli kalemler (tam liste 162 poz)
    ("PH.MEK.001","MEKANİK","CS Boru Fittings Montaj DN15","Adet",420,""),
    ("PH.MEK.002","MEKANİK","CS Boru Fittings Montaj DN20","Adet",420,""),
    ("PH.MEK.003","MEKANİK","CS Boru Fittings Montaj DN25","Adet",490,""),
    ("PH.MEK.004","MEKANİK","CS Boru Fittings Montaj DN32","Adet",490,""),
    ("PH.MEK.005","MEKANİK","CS Boru Fittings Montaj DN40","Adet",560,""),
    ("PH.MEK.006","MEKANİK","CS Boru Fittings Montaj DN50","Adet",630,""),
    ("PH.MEK.007","MEKANİK","CS Boru Fittings Montaj DN65","Adet",700,""),
    ("PH.MEK.008","MEKANİK","CS Boru Fittings Montaj DN80","Adet",840,""),
    ("PH.MEK.009","MEKANİK","CS Boru Fittings Montaj DN100","Adet",980,""),
    ("PH.MEK.010","MEKANİK","CS Boru Fittings Montaj DN125","Adet",1260,""),
    ("PH.MEK.011","MEKANİK","CS Boru Fittings Montaj DN150","Adet",1540,""),
    ("PH.MEK.012","MEKANİK","CS Boru Fittings Montaj DN200","Adet",2100,""),
    ("PH.MEK.013","MEKANİK","CS Boru Fittings Montaj DN250","Adet",2800,""),
    ("PH.MEK.014","MEKANİK","CS Boru Fittings Montaj DN300","Adet",3360,""),
    ("PH.MEK.015","MEKANİK","CS Boru Fittings Montaj DN350","Adet",3920,""),
    ("PH.MEK.016","MEKANİK","CS Boru Fittings Demontaj DN15","Adet",280,""),
    ("PH.MEK.017","MEKANİK","CS Boru Fittings Demontaj DN20","Adet",280,""),
    ("PH.MEK.018","MEKANİK","CS Boru Fittings Demontaj DN25","Adet",315,""),
    ("PH.MEK.019","MEKANİK","CS Boru Fittings Demontaj DN32","Adet",315,""),
    ("PH.MEK.020","MEKANİK","CS Boru Fittings Demontaj DN40","Adet",350,""),
    ("PH.MEK.021","MEKANİK","CS Boru Fittings Demontaj DN50","Adet",385,""),
    ("PH.MEK.022","MEKANİK","CS Boru Fittings Demontaj DN65","Adet",420,""),
    ("PH.MEK.023","MEKANİK","CS Boru Fittings Demontaj DN80","Adet",490,""),
    ("PH.MEK.024","MEKANİK","CS Boru Fittings Demontaj DN100","Adet",560,""),
    ("PH.MEK.025","MEKANİK","CS Boru Fittings Demontaj DN150","Adet",840,""),
    ("PH.MEK.026","MEKANİK","CS Boru Fittings Demontaj DN200","Adet",1120,""),
    ("PH.MEK.027","MEKANİK","CS Boru Fittings Demontaj DN250","Adet",1400,""),
    ("PH.MEK.028","MEKANİK","CS Boru Fittings Demontaj DN300","Adet",1680,""),
    ("PH.MEK.029","MEKANİK","CS Boru Kaynak DN15-DN50 (karbon çelik)","Adet",700,""),
    ("PH.MEK.030","MEKANİK","CS Boru Kaynak DN65-DN100","Adet",1400,""),
    ("PH.MEK.031","MEKANİK","CS Boru Kaynak DN150","Adet",2100,""),
    ("PH.MEK.032","MEKANİK","CS Boru Kaynak DN200","Adet",2800,""),
    ("PH.MEK.033","MEKANİK","CS Boru Kaynak DN250","Adet",3500,""),
    ("PH.MEK.034","MEKANİK","CS Boru Kaynak DN300","Adet",4200,""),
    ("PH.MEK.035","MEKANİK","CS Boru Kaynak DN350","Adet",4900,""),
    ("PH.MEK.036","MEKANİK","SS Boru Fittings Montaj DN15","Adet",630,""),
    ("PH.MEK.037","MEKANİK","SS Boru Fittings Montaj DN25","Adet",735,""),
    ("PH.MEK.038","MEKANİK","SS Boru Fittings Montaj DN50","Adet",945,""),
    ("PH.MEK.039","MEKANİK","SS Boru Fittings Montaj DN80","Adet",1260,""),
    ("PH.MEK.040","MEKANİK","SS Boru Fittings Montaj DN100","Adet",1470,""),
    ("PH.MEK.041","MEKANİK","SS Boru Fittings Montaj DN150","Adet",2310,""),
    ("PH.MEK.042","MEKANİK","SS Boru Fittings Montaj DN200","Adet",3150,""),
    ("PH.MEK.043","MEKANİK","SS Boru Kaynak DN15-DN50 (paslanmaz)","Adet",1050,""),
    ("PH.MEK.044","MEKANİK","SS Boru Kaynak DN65-DN100","Adet",2100,""),
    ("PH.MEK.045","MEKANİK","SS Boru Kaynak DN150","Adet",3150,""),
    ("PH.MEK.046","MEKANİK","SS Boru Kaynak DN200","Adet",4200,""),
    ("PH.MEK.047","MEKANİK","Galvaniz Boru Fittings Montaj DN15","Adet",490,""),
    ("PH.MEK.048","MEKANİK","Galvaniz Boru Fittings Montaj DN25","Adet",560,""),
    ("PH.MEK.049","MEKANİK","Galvaniz Boru Fittings Montaj DN50","Adet",700,""),
    ("PH.MEK.050","MEKANİK","Galvaniz Boru Fittings Montaj DN80","Adet",980,""),
    ("PH.MEK.051","MEKANİK","Vana Montaj CS DN15","Adet",350,""),
    ("PH.MEK.052","MEKANİK","Vana Montaj CS DN20","Adet",350,""),
    ("PH.MEK.053","MEKANİK","Vana Montaj CS DN25","Adet",420,""),
    ("PH.MEK.054","MEKANİK","Vana Montaj CS DN32","Adet",420,""),
    ("PH.MEK.055","MEKANİK","Vana Montaj CS DN40","Adet",490,""),
    ("PH.MEK.056","MEKANİK","Vana Montaj CS DN50","Adet",560,""),
    ("PH.MEK.057","MEKANİK","Vana Montaj CS DN65","Adet",630,""),
    ("PH.MEK.058","MEKANİK","Vana Montaj CS DN80","Adet",770,""),
    ("PH.MEK.059","MEKANİK","Vana Montaj CS DN100","Adet",910,""),
    ("PH.MEK.060","MEKANİK","Vana Montaj CS DN150","Adet",1400,""),
    ("PH.MEK.061","MEKANİK","Vana Montaj CS DN200","Adet",1960,""),
    ("PH.MEK.062","MEKANİK","Vana Montaj CS DN250","Adet",2520,""),
    ("PH.MEK.063","MEKANİK","Vana Montaj CS DN300","Adet",3080,""),
    ("PH.MEK.064","MEKANİK","Vana Demontaj CS DN15","Adet",210,""),
    ("PH.MEK.065","MEKANİK","Vana Demontaj CS DN50","Adet",315,""),
    ("PH.MEK.066","MEKANİK","Vana Demontaj CS DN100","Adet",525,""),
    ("PH.MEK.067","MEKANİK","Vana Demontaj CS DN200","Adet",1120,""),
    ("PH.MEK.068","MEKANİK","Vana Montaj SS DN15","Adet",525,""),
    ("PH.MEK.069","MEKANİK","Vana Montaj SS DN50","Adet",840,""),
    ("PH.MEK.070","MEKANİK","Vana Montaj SS DN100","Adet",1365,""),
    ("PH.MEK.071","MEKANİK","Vana Montaj SS DN200","Adet",2940,""),
    ("PH.MEK.072","MEKANİK","Kör Flanş CS DN15-DN50 Montaj","Adet",280,""),
    ("PH.MEK.073","MEKANİK","Kör Flanş CS DN65-DN100 Montaj","Adet",420,""),
    ("PH.MEK.074","MEKANİK","Kör Flanş CS DN150-DN200 Montaj","Adet",700,""),
    ("PH.MEK.075","MEKANİK","Deluge Vana Montaj 4\"","Adet",5600,""),
    ("PH.MEK.076","MEKANİK","Deluge Vana Montaj 6\"","Adet",7000,""),
    ("PH.MEK.077","MEKANİK","Pompa Montaj (genel, malzemesiz)","Adet",28000,""),
    ("PH.MEK.078","MEKANİK","Pompa Demontaj (genel)","Adet",16800,""),
    # İNŞAAT
    ("PH.INS.001","İNŞAAT","Betopan Levha Montaj","m²",402.58,"DOĞRULANMALI"),
    ("PH.INS.002","İNŞAAT","Alçıpan Levha Montaj","m²",402.58,"DOĞRULANMALI"),
    ("PH.INS.003","İNŞAAT","Trapez Sac Cephe/Çatı Kaplama","m²",644,"DOĞRULANMALI"),
    ("PH.INS.004","İNŞAAT","Metal Kapı Montaj (≤1.5 m²)","Adet",None,"DOĞRULANMALI"),
    ("PH.INS.005","İNŞAAT","Metal Kapı Montaj (>1.5 m²)","Adet",None,"DOĞRULANMALI"),
    # ÇELİK
    ("PH.ÇLK.001","ÇELİK","Çelik Konstrüksiyon İmalat (malzeme hariç)","kg",140,""),
    ("PH.ÇLK.002","ÇELİK","Çelik Konstrüksiyon Montaj","kg",112,""),
    ("PH.ÇLK.003","ÇELİK","Çelik Konstrüksiyon Demontaj","kg",56,""),
    ("PH.ÇLK.004","ÇELİK","Galvaniz Petek/Izgara Döşeme","m²",70,""),
    # HDPE — seçilmiş kalemler (tam 125 poz)
    ("PH.PE.001","HDPE","HDPE100 Boru Montaj Ø20mm","m",None,"DOĞRULANMALI"),
    ("PH.PE.002","HDPE","HDPE100 Boru Montaj Ø25mm","m",None,"DOĞRULANMALI"),
    ("PH.PE.003","HDPE","HDPE100 Boru Montaj Ø32mm","m",None,"DOĞRULANMALI"),
    ("PH.PE.004","HDPE","HDPE100 Boru Montaj Ø40mm","m",None,"DOĞRULANMALI"),
    ("PH.PE.005","HDPE","HDPE100 Boru Montaj Ø50mm","m",None,"DOĞRULANMALI"),
    ("PH.PE.006","HDPE","HDPE100 Boru Montaj Ø63mm","m",None,"DOĞRULANMALI"),
    ("PH.PE.007","HDPE","HDPE100 Boru Montaj Ø75mm","m",None,"DOĞRULANMALI"),
    ("PH.PE.008","HDPE","HDPE100 Boru Montaj Ø90mm","m",None,"DOĞRULANMALI"),
    ("PH.PE.009","HDPE","HDPE100 Boru Montaj Ø110mm","m",None,"DOĞRULANMALI"),
    ("PH.PE.010","HDPE","HDPE100 Boru Montaj Ø125mm","m",None,"DOĞRULANMALI"),
    ("PH.PE.011","HDPE","HDPE100 Boru Montaj Ø160mm","m",None,"DOĞRULANMALI"),
    ("PH.PE.012","HDPE","HDPE100 Boru Montaj Ø200mm","m",None,"DOĞRULANMALI"),
    ("PH.PE.013","HDPE","HDPE100 Boru Montaj Ø250mm","m",None,"DOĞRULANMALI"),
    ("PH.PE.014","HDPE","HDPE100 Boru Montaj Ø315mm","m",None,"DOĞRULANMALI"),
    ("PH.PE.015","HDPE","HDPE100 Boru Montaj Ø400mm","m",None,"DOĞRULANMALI"),
    ("PH.PE.016","HDPE","HDPE100 Boru Montaj Ø500mm","m",None,"DOĞRULANMALI"),
    ("PH.PE.017","HDPE","HDPE100 Boru Montaj Ø630mm","m",None,"DOĞRULANMALI"),
    ("PH.PE.018","HDPE","HDPE100 Boru Montaj Ø710mm","m",None,"DOĞRULANMALI"),
    ("PH.PE.019","HDPE","HDPE100 Boru Demontaj Ø20-63mm","m",None,"DOĞRULANMALI"),
    ("PH.PE.020","HDPE","HDPE100 Boru Demontaj Ø75-125mm","m",None,"DOĞRULANMALI"),
    ("PH.PE.021","HDPE","HDPE100 Boru Demontaj Ø160-200mm","m",None,"DOĞRULANMALI"),
    ("PH.PE.022","HDPE","HDPE100 Boru Alın Kaynak Ø20-63mm","Adet",None,"DOĞRULANMALI"),
    ("PH.PE.023","HDPE","HDPE100 Boru Alın Kaynak Ø75-125mm","Adet",None,"DOĞRULANMALI"),
    ("PH.PE.024","HDPE","HDPE100 Boru Alın Kaynak Ø160-200mm","Adet",None,"DOĞRULANMALI"),
    ("PH.PE.025","HDPE","HDPE100 Elektrofüzyon Kaynak Ø20-63mm","Adet",None,"DOĞRULANMALI"),
    ("PH.PE.026","HDPE","HDPE100 Elektrofüzyon Kaynak Ø75-125mm","Adet",None,"DOĞRULANMALI"),
    ("PH.PE.027","HDPE","HDPE100 Elektrofüzyon Kaynak Ø160-200mm","Adet",None,"DOĞRULANMALI"),
    ("PH.PE.028","HDPE","HDPE100 Vana Montaj Ø25-63mm","Adet",None,"DOĞRULANMALI"),
    ("PH.PE.029","HDPE","HDPE100 Vana Montaj Ø75-125mm","Adet",None,"DOĞRULANMALI"),
    ("PH.PE.030","HDPE","HDPE100 Vana Montaj Ø160-200mm","Adet",None,"DOĞRULANMALI"),
]

# Dönmez Kumlama
DONMEZ_SUTUNLAR = [
    ("Poz No", 16), ("Kategori", 22), ("Açıklama", 55),
    ("Birim", 10), ("Birim Fiyat (₺)", 18), ("Not / Kaynak", 28),
]

DONMEZ_DATA = [
    ("PPBF-100-001","KUMLAMA","Yürüme Yolu Çizimi","m",28.30,""),
    ("PPBF-100-002","KUMLAMA","Yüzey Kumlama — Bazalt","m²",840.25,""),
    ("PPBF-100-003","KUMLAMA","Yüzey Kumlama — Cüruf","m²",1235.25,""),
    ("PPBF-100-004","KUMLAMA","Yüzey Kumlama — Grit (Çelik bilye)","m²",1770.25,""),
    ("PPBF-100-005","İSKELE","İskele Kurulum/Söküm","m³",685.50,""),
    ("PPBF-100-006","İSKELE","Küçük İskele Pozu","m³",685.50,""),
    ("PPBF-100-007","YEVMİYE","Yevmiye (Gün)","Gün",4200,""),
    ("PPBF-100-008","YEVMİYE","Yevmiye (Saat)","Saat",625,""),
    ("PPBF-100-009","KİMYASAL","Kimyasal Temizlik","m²",869.35,""),
    ("PPBF-100-010","ŞABLON","Tank Şablon İşlemi","Adet",45925.30,""),
    ("PPBF-100-011","BOYAMA","Boya — 1 Kat (genel)","m²",None,"DOĞRULANMALI"),
    ("PPBF-100-012","BOYAMA","Boya — 2 Kat","m²",None,"DOĞRULANMALI"),
    ("PPBF-100-013","BOYAMA","Boya — 3 Kat","m²",None,"DOĞRULANMALI"),
    ("PPBF-100-014","BOYAMA","Boya — Zemin/Epoksi 1 Kat","m²",None,"DOĞRULANMALI"),
    ("PPBF-100-015","BOYAMA","Boya — Zemin/Epoksi 2 Kat","m²",None,"DOĞRULANMALI"),
    ("PPBF-100-016","BOYAMA","Boya Demaj/Sıyırma","m²",None,"DOĞRULANMALI"),
    ("PPBF-100-017","BOYAMA","Çatı Boya 1 Kat","m²",None,"DOĞRULANMALI"),
    ("PPBF-100-018","BOYAMA","Çatı Boya 2 Kat","m²",None,"DOĞRULANMALI"),
    ("PPBF-100-019","KORUMA","Pas Önleyici Astar 1 Kat","m²",None,"DOĞRULANMALI"),
    ("PPBF-100-020","KORUMA","Pas Önleyici Astar 2 Kat","m²",None,"DOĞRULANMALI"),
    ("PPBF-100-021","KORUMA","Epoksi Astar 1 Kat","m²",None,"DOĞRULANMALI"),
    ("PPBF-100-022","KORUMA","Poliüretan Topkat 1 Kat","m²",None,"DOĞRULANMALI"),
    ("PPBF-100-023","KORUMA","Poliüretan Topkat 2 Kat","m²",None,"DOĞRULANMALI"),
    ("PPBF-100-024","KORUMA","Yüksek Sıcaklık Boya 1 Kat (>200°C)","m²",None,"DOĞRULANMALI"),
    ("PPBF-100-025","KORUMA","Yüksek Sıcaklık Boya 2 Kat","m²",None,"DOĞRULANMALI"),
    ("PPBF-100-026","KORUMA","Galvaniz Sonrası Boya","m²",None,"DOĞRULANMALI"),
    ("PPBF-100-027","MALZEME","Rulo Boya Uygulaması","m²",None,"DOĞRULANMALI"),
    ("PPBF-100-028","MALZEME","Airless Püskürtme","m²",None,"DOĞRULANMALI"),
    ("PPBF-100-029","MALZEME","Fırça Uygulaması","m²",None,"DOĞRULANMALI"),
    ("PPBF-100-030","TAMİR","Boya Tamir/Rötuş","m²",None,"DOĞRULANMALI"),
    ("PPBF-100-031","TAMİR","Kaynak Yeri Boya Tamir","m",None,"DOĞRULANMALI"),
    ("PPBF-100-032","NAKLİYE","Ekipman Nakliye (saha içi)","Sefer",None,"DOĞRULANMALI"),
    ("PPBF-100-033","DİĞER","Renk Kodu Yazma/Etiketleme","Adet",None,"DOĞRULANMALI"),
    ("PPBF-100-034","DİĞER","İşaret/Uyarı Levha Boya","Adet",None,"DOĞRULANMALI"),
    ("PPBF-100-035","DİĞER","Zemin Çizgi Boyama","m",None,"DOĞRULANMALI"),
    ("PPBF-100-036","DİĞER","Zemin Sembol Boyama","Adet",None,"DOĞRULANMALI"),
    ("PPBF-100-037","DİĞER","Flanş/Bağlantı Koruma Kapağı Montaj","Adet",None,"DOĞRULANMALI"),
    ("PPBF-100-038","DİĞER","Paket İş (teklif bazlı)","Adet",None,"DOĞRULANMALI"),
]

# NOT İnşaat
NOT_SUTUNLAR = [
    ("Poz No", 14), ("Kategori", 22), ("Açıklama", 55),
    ("Birim", 10), ("Birim Fiyat (₺)", 18), ("Not / Kaynak", 28),
]

NOT_DATA = [
    ("PPRC.001","BORU","PPRC Boru Montaj DN20 (3/4\")","m",1600,""),
    ("PPRC.002","BORU","PPRC Boru Montaj DN25 (1\")","m",1800,""),
    ("PPRC.003","BORU","PPRC Boru Montaj DN32 (1¼\")","m",2100,""),
    ("PPRC.004","BORU","PPRC Boru Montaj DN40 (1½\")","m",2350,""),
    ("PPRC.005","BORU","PPRC Boru Montaj DN50 (2\")","m",2350,""),
    ("PPRC.006","BORU","PPRC Boru Montaj DN63 (2½\")","m",2350,""),
    ("PVC.001","BORU","PVC/Koruge Boru Montaj DN75","m",1900,""),
    ("PVC.002","BORU","PVC/Koruge Boru Montaj DN100","m",2100,""),
    ("PVC.003","BORU","PVC/Koruge Boru Montaj DN150","m",2800,""),
    ("PVC.004","BORU","PVC/Koruge Boru Montaj DN200","m",3200,""),
    ("PVC.005","BORU","PVC/Koruge Boru Montaj DN300","m",3800,""),
    ("ARM.001","TESİSAT","Armatür — Batarya Değişimi (banyo/mutfak)","Adet",None,"DOĞRULANMALI"),
    ("ARM.002","TESİSAT","Armatür — Lavabo Montaj","Adet",None,"DOĞRULANMALI"),
    ("ARM.003","TESİSAT","Armatür — Klozet Montaj","Adet",None,"DOĞRULANMALI"),
    ("ARM.004","TESİSAT","Rezervuar Değişimi","Adet",None,"DOĞRULANMALI"),
    ("ARM.005","TESİSAT","Duş Kabini Montaj","Adet",None,"DOĞRULANMALI"),
    ("KPD.001","KAPLAMA","Seramik Döşeme (≤600x600mm)","m²",None,"DOĞRULANMALI"),
    ("KPD.002","KAPLAMA","Seramik Duvar Kaplaması","m²",None,"DOĞRULANMALI"),
    ("KPD.003","KAPLAMA","Granit Döşeme","m²",None,"DOĞRULANMALI"),
    ("KPD.004","KAPLAMA","Epoksi Zemin Kaplama","m²",None,"DOĞRULANMALI"),
    ("CAT.001","ASMA TAVAN","Asma Tavan — Alçıpan","m²",None,"DOĞRULANMALI"),
    ("CAT.002","ASMA TAVAN","Asma Tavan — Metal Karkas","m²",None,"DOĞRULANMALI"),
    ("BTN.001","BETON/KAZI","El Kazısı","m³",250,""),
    ("BTN.002","BETON/KAZI","Moloz Nakli","m³",1000,""),
    ("BTN.003","BETON/KAZI","Çelik Hasır (düz)","kg",95,""),
    ("BTN.004","BETON/KAZI","Demir Donatı (nervürlü)","kg",None,"DOĞRULANMALI"),
    ("BTN.005","BETON/KAZI","C20 Beton Dökme","m³",None,"DOĞRULANMALI"),
    ("BTN.006","BETON/KAZI","C25 Beton Dökme","m³",None,"DOĞRULANMALI"),
    ("BTN.007","BETON/KAZI","Kalıp İşçiliği","m²",None,"DOĞRULANMALI"),
    ("MF.001","GENEL","Yatay Taşıma (saha içi, el arabası)","Adet",150,""),
    ("MF.002","GENEL","Genel Temizlik/İşçilik","Saat",1350,""),
    ("MF.003","GENEL","Yükleme-Boşaltma","Saat",1350,""),
    ("MF.004","GENEL","Pencere/Kapı Silikon Uygulama","m",None,"DOĞRULANMALI"),
    ("MF.005","GENEL","Panel Kapı Montaj","Adet",None,"DOĞRULANMALI"),
    ("MF.006","GENEL","PVC Pencere Montaj","m²",None,"DOĞRULANMALI"),
]

# Manrel
MANREL_SUTUNLAR = [
    ("Poz No", 14), ("Tip", 15), ("Açıklama", 50),
    ("Boyut", 12), ("Birim Fiyat (₺)", 18), ("Not / Kaynak", 28),
]

MANREL_DATA = [
    ("MNR.PSV.001","PSV/TRV TEST","Emniyet/Tahliye Vanası Test","1/2\"",1000,""),
    ("MNR.PSV.002","PSV/TRV TEST","Emniyet/Tahliye Vanası Test","3/4\"",1000,""),
    ("MNR.PSV.003","PSV/TRV TEST","Emniyet/Tahliye Vanası Test","1\"",1500,""),
    ("MNR.PSV.004","PSV/TRV TEST","Emniyet/Tahliye Vanası Test","1½\"",1500,""),
    ("MNR.PSV.005","PSV/TRV TEST","Emniyet/Tahliye Vanası Test","2\"",2000,""),
    ("MNR.PSV.006","PSV/TRV TEST","Emniyet/Tahliye Vanası Test","3\"",3000,""),
    ("MNR.PSV.007","PSV/TRV TEST","Emniyet/Tahliye Vanası Test","4\"",4000,""),
    ("MNR.PSV.008","PSV/TRV TEST","Emniyet/Tahliye Vanası Test","6\"",6000,""),
    ("MNR.PSV.009","PSV/TRV TEST","Emniyet/Tahliye Vanası Test","8\"",8000,""),
    ("MNR.PSV.010","PSV/TRV TEST","Emniyet/Tahliye Vanası Test","10\"",10000,""),
    ("MNR.PSV.011","PSV/TRV TEST","Emniyet/Tahliye Vanası Test","12\"",16000,""),
    ("MNR.BAK.001","PSV/TRV BAKIM-A","PSV/TRV Bakım-A (söküm+revizyon+test)","1/2\"",3250,""),
    ("MNR.BAK.002","PSV/TRV BAKIM-A","PSV/TRV Bakım-A","3/4\"",3250,""),
    ("MNR.BAK.003","PSV/TRV BAKIM-A","PSV/TRV Bakım-A","1\"",4750,""),
    ("MNR.BAK.004","PSV/TRV BAKIM-A","PSV/TRV Bakım-A","1½\"",4750,""),
    ("MNR.BAK.005","PSV/TRV BAKIM-A","PSV/TRV Bakım-A","2\"",6300,""),
    ("MNR.BAK.006","PSV/TRV BAKIM-A","PSV/TRV Bakım-A","3\"",9350,""),
    ("MNR.BAK.007","PSV/TRV BAKIM-A","PSV/TRV Bakım-A","4\"",12400,""),
    ("MNR.BAK.008","PSV/TRV BAKIM-A","PSV/TRV Bakım-A","6\"",19100,""),
    ("MNR.BAK.009","PSV/TRV BAKIM-A","PSV/TRV Bakım-A","8\"",25700,""),
    ("MNR.BAK.010","PSV/TRV BAKIM-A","PSV/TRV Bakım-A","10\"",32500,""),
    ("MNR.BAK.011","PSV/TRV BAKIM-A","PSV/TRV Bakım-A","12\"",41300,""),
    ("MNR.PVV.001","PVV TEST","Patlama Önleme Vanası Test","4\"",3350,""),
    ("MNR.PVV.002","PVV TEST","Patlama Önleme Vanası Test","6\"",3350,""),
    ("MNR.PVV.003","PVV TEST","Patlama Önleme Vanası Test","8\"",3350,""),
    ("MNR.PVV.004","PVV TEST","Patlama Önleme Vanası Test","10\"",3350,""),
    ("MNR.PVV.005","PVV TEST","Patlama Önleme Vanası Test","12\"",3350,""),
    ("MNR.PVV.006","PVV TEST","Patlama Önleme Vanası Test","16\"",3350,""),
    ("MNR.PVV.007","PVV TEST","Patlama Önleme Vanası Test","20\"",3350,""),
    ("MNR.PVV.008","PVV TEST","Patlama Önleme Vanası Test","24\"",3350,""),
    ("MNR.AT.001","ALEV TUTUCU BAKIM-A","Alev Tutucu Bakım-A","4\"",8300,""),
    ("MNR.AT.002","ALEV TUTUCU BAKIM-A","Alev Tutucu Bakım-A","6\"",8300,""),
    ("MNR.AT.003","ALEV TUTUCU BAKIM-A","Alev Tutucu Bakım-A","8\"",10400,""),
    ("MNR.AT.004","ALEV TUTUCU BAKIM-A","Alev Tutucu Bakım-A","10\"",12550,""),
    ("MNR.AT.005","ALEV TUTUCU BAKIM-A","Alev Tutucu Bakım-A","12\"",14650,""),
    ("MNR.AT.006","ALEV TUTUCU BAKIM-A","Alev Tutucu Bakım-A","16\"",16750,""),
    ("MNR.AT.007","ALEV TUTUCU BAKIM-A","Alev Tutucu Bakım-A","20\"",18500,""),
    ("MNR.AT.008","ALEV TUTUCU BAKIM-A","Alev Tutucu Bakım-A","24\"",18500,""),
]

# ── EXCEL YARAT ───────────────────────────────────────────────────────────────

os.makedirs(CIKTI, exist_ok=True)
wb = openpyxl.Workbook()

# 1. ÖZET sayfası
ws_ozet = wb.active
ws_ozet.title = "ÖZET"
ws_ozet.sheet_view.showGridLines = False
ws_ozet.column_dimensions["A"].width = 36
ws_ozet.column_dimensions["B"].width = 20
ws_ozet.column_dimensions["C"].width = 22
ws_ozet.column_dimensions["D"].width = 16
ws_ozet.column_dimensions["E"].width = 16

ws_ozet.merge_cells("A1:E1")
ws_ozet["A1"] = "POLİSAN KİMYA — BİRİM FİYAT HAVUZU 2026"
bh(ws_ozet["A1"], RENK["baslik"], boyut=14)
ws_ozet.row_dimensions[1].height = 30

ws_ozet.merge_cells("A2:E2")
ws_ozet["A2"] = f"Hazırlanma: {date.today().strftime('%d.%m.%Y')}  |  4 Sözleşme  |  KDV Hariç Fiyatlar  |  DOĞRULANMALI hücreleri kontrol edin"
ws_ozet["A2"].font = Font(name="Calibri", size=9, italic=True, color="555555")
ws_ozet["A2"].alignment = Alignment(horizontal="center")
ws_ozet.row_dimensions[2].height = 14
ws_ozet.row_dimensions[3].height = 8

baslik_r = 4
for ci, metin in enumerate(["Sözleşme / Firma", "Sözleşme No", "Sözleşme Türü", "Poz Sayısı", "Sayfa"], 1):
    h = ws_ozet.cell(row=baslik_r, column=ci, value=metin)
    bh(h, RENK["alt_baslik"], boyut=10)
ws_ozet.row_dimensions[baslik_r].height = 22

sozlesmeler = [
    ("Özgül Teknik Metal İnş.", "2026/1013", "Birim Fiyat", len(OZGUL_DATA), "Özgül Teknik"),
    ("Dönmez Kumlama Boya", "2026/1015", "Birim Fiyat", len(DONMEZ_DATA), "Dönmez Kumlama"),
    ("NOT İnşaat", "2026/1023", "Birim Fiyat", len(NOT_DATA), "NOT İnşaat"),
    ("Manrel Mühendislik", "2026/1043", "Birim Fiyat", len(MANREL_DATA), "Manrel"),
]

for ri, (firma, szl_no, tur, poz_sayisi, sayfa) in enumerate(sozlesmeler, 5):
    renk = RENK["vurgu"] if ri % 2 == 0 else None
    ws_ozet.row_dimensions[ri].height = 18
    for ci, deger in enumerate([firma, szl_no, tur, poz_sayisi, sayfa], 1):
        c = ws_ozet.cell(row=ri, column=ci, value=deger)
        bv(c, renk, merkez=(ci >= 3))

ws_ozet.row_dimensions[9].height = 10

uyari = ws_ozet.cell(row=10, column=1,
    value="⚠ 'DOĞRULANMALI' yazan hücreler: sözleşme orijinalinden teyit edilmemiş kalemler. "
          "PDF'de sayısal değer okunamadı — sözleşme ekinden doğrulayın.")
ws_ozet.merge_cells("A10:E10")
uyari.font = Font(name="Calibri", size=9, color="7F6000", bold=True)
uyari.fill = PatternFill("solid", fgColor=RENK["uyari"])
uyari.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
uyari.border = kenar()
ws_ozet.row_dimensions[10].height = 28

# 2-5. Sözleşme sayfaları
sayfa_olustur(wb, "Özgül Teknik", "2026/1013", "Özgül Teknik Metal İnş.", OZGUL_SUTUNLAR, OZGUL_DATA)
sayfa_olustur(wb, "Dönmez Kumlama", "2026/1015", "Dönmez Kumlama Boya İnş.", DONMEZ_SUTUNLAR, DONMEZ_DATA)
sayfa_olustur(wb, "NOT İnşaat", "2026/1023", "NOT İnşaat", NOT_SUTUNLAR, NOT_DATA)
sayfa_olustur(wb, "Manrel", "2026/1043", "Manrel Mühendislik", MANREL_SUTUNLAR, MANREL_DATA)

wb.save(DOSYA)
print(f"Dosya oluşturuldu: {DOSYA}")
print(f"Toplam poz: Özgül={len(OZGUL_DATA)}, Dönmez={len(DONMEZ_DATA)}, NOT={len(NOT_DATA)}, Manrel={len(MANREL_DATA)}")
