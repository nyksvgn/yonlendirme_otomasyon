 
import streamlit as st
import openpyxl
import io
import re
from datetime import datetime

st.set_page_config(page_title="Yönlendirme Aktarımı", page_icon="📊")

st.title("📦 Yönlendirme Otomasyonu")

uploaded_po = st.file_uploader("ZTM003 dosyasını yükleyin (.xlsx)", type=["xlsx"])
uploaded_yon = st.file_uploader("Yönlendirme şablon dosyasını yükleyin (.xlsx)", type=["xlsx"])

def normalize(val):
    val = re.sub(r"[\r\n\t]", "", str(val)).lower()
    return re.sub(r"\s+", "", val)

if uploaded_po and uploaded_yon:
    if st.button("🚀 Verileri Aktar ve Dosyayı Oluştur"):
        mapping = {
            normalize("Alıcı"): "Sipariş veren bayi/dist Kodu",
            normalize("Üretim yeri"): "Yönlendirme Yapılan Fabrika Kodu (2. SN)",
            normalize("Kapı Çıkış Tarihi"): "Fatura Tarihi",
            normalize("Ürün"): "Ürün Kodu (SKU)",
            normalize("Teslimat Miktarı"): "Adet (Tava\Koli\Kasa)",
            normalize("yönlendirme nedeni"): "Yönlendirme yapma nedeni"
        }

        nakliye_kod_map = {
            ("ZTIR", "Gidiş"): "ZTIR01",
            ("ZTIR", "GidişDönüş"): "ZTIR02",
            ("ZKMY", "Gidiş"): "ZKMY01",
            ("ZKMY", "GidişDönüş"): "ZKMY02"
         }

        try:
            wb_src = openpyxl.load_workbook(uploaded_po, data_only=True)
            ws_src = wb_src["Data"]
            wb_dst = openpyxl.load_workbook(uploaded_yon)
            ws_dst = wb_dst["Ana_sayfa"]

            src_headers = {normalize(cell.value): idx for idx, cell in enumerate(ws_src[1])}
            dst_headers = {normalize(cell.value): idx for idx, cell in enumerate(ws_dst[1])}
            sevk_idx = src_headers.get("teslimatmiktarı")
            dst_row = 2
            aktarilan = 0

            for row in ws_src.iter_rows(min_row=2, values_only=True):
                sevk_miktar = row[sevk_idx] if sevk_idx is not None else None
                if sevk_miktar in [0, 0.0, "0", "0.0", None, ""]:
                    continue

                for src_col, dst_col in mapping.items():
                    src_idx = src_headers.get(src_col)
                    dst_idx = dst_headers.get(normalize(dst_col))
                    if src_idx is None or dst_idx is None:
                        continue
                    ws_dst.cell(row=dst_row, column=dst_idx+1, value=row[src_idx])

                nk_idx = src_headers.get("Nakliye araçları")
                yon_idx = src_headers.get("Nakliye Tipi Tanımı")
                dst_nt_idx = dst_headers.get("Nakliye Tipi")
                if None not in (nk_idx, yon_idx, dst_nt_idx):
                    def clean(val):
                       return re.sub(r"\s+", "", str(val).strip().upper())
                    def clean_title(val):
                        return re.sub(r"[-–]", "", str(val).strip().title())  
                    nk_val = clean(row[nk_idx]) if row[nk_idx] else ""
                    yon_val = clean_title(row[yon_idx]) if row[yon_idx] else ""
                    combined = nakliye_kod_map.get((nk_val, yon_val), f"{nk_val}{yon_val}")
                    ws_dst.cell(row=dst_row, column=dst_nt_idx+1, value=combined)

                dst_row += 1
                aktarilan += 1

            buffer = io.BytesIO()
            wb_dst.save(buffer)
            buffer.seek(0)

            filename = f"Yönlendirme_{datetime.today().strftime('%d%m%Y')}.xlsx"
            st.success(f"✅ {aktarilan} satır başarıyla aktarıldı.")
            st.download_button("📥 Dosyayı indir", buffer, file_name=filename)

        except Exception as e:
            st.error(f"❌ Hata oluştu: {e}")
