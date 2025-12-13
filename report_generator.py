from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Alignment, Font
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.page import PageMargins
from PIL import Image as PILImage
from datetime import datetime
import os
import re
import subprocess
import shutil
import tempfile
import uuid

# -------------------------------------------------
# YOLLAR
# -------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_FILE = os.path.join(BASE_DIR, "rapor.xlsx")
OUTPUT_DIR = os.path.join(BASE_DIR, "generated_pdfs")

os.makedirs(OUTPUT_DIR, exist_ok=True)

# -------------------------------------------------
# FOTOĞRAF ALANLARI
# -------------------------------------------------
PHOTO_AREAS = {
    1: ("B", 29, "I", 47),
    2: ("J", 29, "P", 47),
    3: ("B", 52, "I", 70),
    4: ("J", 52, "P", 70),
    5: ("B", 75, "I", 93),
    6: ("J", 75, "P", 93),
    7: ("B", 98, "I", 116),
    8: ("J", 98, "P", 116),
}

# -------------------------------------------------
# TÜRKÇE NORMALIZE
# -------------------------------------------------
def normalize(text: str) -> str:
    return (
        text.lower()
        .replace("ı", "i")
        .replace("ğ", "g")
        .replace("ü", "u")
        .replace("ş", "s")
        .replace("ö", "o")
        .replace("ç", "c")
        .replace(":", "")
        .strip()
    )

# -------------------------------------------------
# FOTOĞRAF ALAN BOYUTU (PX)
# -------------------------------------------------
def area_pixel_size(ws, area):
    sc, sr, ec, er = area

    width = sum(
        (ws.column_dimensions[ws.cell(1, c).column_letter].width or 8) * 7
        for c in range(column_index_from_string(sc), column_index_from_string(ec) + 1)
    )

    height = sum(
        (ws.row_dimensions[r].height or 15) * 1.25
        for r in range(sr, er + 1)
    )

    return int(width * 0.95), int(height * 0.95)

# -------------------------------------------------
# GERÇEK MERKEZ ANCHOR
# -------------------------------------------------
def smart_anchor(ws, area, img_height_px):
    sc, sr, ec, er = area
    sc_i = column_index_from_string(sc)
    ec_i = column_index_from_string(ec)

    center_col = sc_i + (ec_i - sc_i) // 2
    center_row = sr + (er - sr) // 2

    px_per_row = 18
    row_offset = int((img_height_px / 2) / px_per_row)

    final_row = max(sr + 1, center_row - row_offset)
    col_letter = ws.cell(row=1, column=center_col).column_letter

    return f"{col_letter}{final_row}"

# -------------------------------------------------
# PRINT AREA VE SAYFA AYARLARI
# -------------------------------------------------
def setup_print_area(ws):
    """
    Excel sayfasına print area (yazdırma alanı) ve sayfa ayarları yapar.
    Tüm içeriği ve resimleri PDF'e dahil etmek için.
    """
    # En son dolu satırı bul
    max_row = ws.max_row
    # Resimlerin olduğu alanları kontrol et
    for area in PHOTO_AREAS.values():
        _, _, _, er = area
        if er > max_row:
            max_row = er
    
    # En son dolu sütunu bul
    max_col = ws.max_column
    # Resimlerin olduğu alanları kontrol et
    for area in PHOTO_AREAS.values():
        _, _, ec, _ = area
        ec_i = column_index_from_string(ec)
        if ec_i > max_col:
            max_col = ec_i
    
    # Print area'yı ayarla (A1'den en son dolu hücreye kadar)
    # Son foto alanının altına kadar
    print_end_col = get_column_letter(max_col)
    print_end_row = max_row + 2  # Biraz ekstra alan
    
    print_area = f"A1:{print_end_col}{print_end_row}"
    ws.print_area = print_area
    
    # Sayfa ayarları
    ws.page_setup.orientation = "portrait"  # "portrait" veya "landscape"
    ws.page_setup.paperSize = 9  # 9 = A4
    
    # Scale'i dinamik olarak hesapla - içeriğin yüksekliğine göre
    # A4 sayfa yüksekliği: ~29.7 cm = ~842 point
    # Excel satır yüksekliği ortalama: ~15 point
    # İçerik yüksekliği: max_row * 15 point (yaklaşık)
    
    # İçeriğin toplam yüksekliğini hesapla
    total_height_points = 0
    for row in range(1, max_row + 1):
        row_height = ws.row_dimensions[row].height
        if row_height:
            total_height_points += row_height * 1.25  # Excel'de point to pixel dönüşümü
        else:
            total_height_points += 15 * 1.25  # Varsayılan satır yüksekliği
    
    # A4 sayfa yüksekliği (point cinsinden, kenar boşlukları hariç)
    # A4: 29.7 cm = 842 point, kenar boşlukları: 0.4 inch = ~28 point (üst+alt)
    usable_page_height = 842 - 56  # ~786 point kullanılabilir alan
    
    # Genişliği de hesapla (yana taşmayı önlemek için)
    total_width_points = 0
    for col in range(1, max_col + 1):
        col_letter = get_column_letter(col)
        col_width = ws.column_dimensions[col_letter].width
        if col_width:
            total_width_points += col_width * 7  # Excel'de genişlik to point dönüşümü
        else:
            total_width_points += 8 * 7  # Varsayılan sütun genişliği
    
    # A4 sayfa genişliği (point cinsinden, kenar boşlukları hariç)
    # A4: 21 cm = 595 point, kenar boşlukları: 0.4 inch = ~28 point (sol+sağ)
    usable_page_width = 595 - 56  # ~539 point kullanılabilir alan
    
    # Scale hesapla: hem yüksekliğe hem genişliğe göre
    scale_by_height = 100
    scale_by_width = 100
    
    if total_height_points > 0:
        scale_by_height = int((usable_page_height / total_height_points) * 100)
    
    if total_width_points > 0:
        scale_by_width = int((usable_page_width / total_width_points) * 100)
    
    # İkisinden küçük olanı al (hem yüksekliğe hem genişliğe sığsın)
    calculated_scale = min(scale_by_height, scale_by_width)
    
    # Biraz daha küçült (güvenlik için %10 daha küçük)
    calculated_scale = int(calculated_scale * 0.90)
    
    # Scale'i %40-%85 arasında tut (çok küçük veya çok büyük olmasın)
    calculated_scale = max(40, min(85, calculated_scale))
    
    # Scale'i ayarla
    try:
        ws.page_setup.scale = calculated_scale
    except Exception:
        # Scale ayarlanamazsa varsayılan değer
        try:
            ws.page_setup.scale = 60
        except Exception:
            pass
    
    # fitToPage'i kaldır - sadece scale kullan
    try:
        if hasattr(ws, 'sheet_properties') and ws.sheet_properties:
            if hasattr(ws.sheet_properties, 'pageSetUpPr'):
                if ws.sheet_properties.pageSetUpPr is not None:
                    ws.sheet_properties.pageSetUpPr.fitToPage = False
                    ws.sheet_properties.pageSetUpPr.fitToWidth = 0
                    ws.sheet_properties.pageSetUpPr.fitToHeight = 0
    except Exception:
        pass
    
    # Kenar boşlukları (margins) - tek sayfaya sığdırmak için küçük
    try:
        ws.page_margins = PageMargins(
            left=0.2,
            right=0.2,
            top=0.2,
            bottom=0.2,
            header=0.1,
            footer=0.1
        )
    except Exception:
        pass  # Bazı şablonlarda çalışmayabilir, devam et
    
    return print_area

# -------------------------------------------------
# EXCEL → PDF (xlsx2pdf - önerilen)
# -------------------------------------------------
def excel_to_pdf_xlsx2pdf(excel_file, pdf_file):
    """
    Excel dosyasını PDF'e çevirir.
    xlsx2pdf kütüphanesi kullanır (sayfa ayarlarını daha iyi okur).
    """
    try:
        # xlsx2pdf 1.0.4 için farklı import denemeleri
        try:
            from xlsx2pdf import convert
            convert(excel_file, pdf_file)
        except (ImportError, AttributeError):
            # Alternatif import yöntemi
            try:
                import xlsx2pdf
                xlsx2pdf.convert(excel_file, pdf_file)
            except (AttributeError, TypeError):
                # xlsx2pdf 1.0.4 için yeni API
                from xlsx2pdf.converter import convert
                convert(excel_file, pdf_file)
        
        if os.path.exists(pdf_file) and os.path.getsize(pdf_file) > 0:
            return True
        return False
    except ImportError as e:
        print(f"xlsx2pdf ImportError: {e}")
        return False
    except Exception as e:
        print(f"xlsx2pdf Exception: {type(e).__name__}: {e}")
        import traceback
        traceback.print_exc()
        return False

# -------------------------------------------------
# EXCEL → PDF (LibreOffice)
# -------------------------------------------------
def excel_to_pdf_libreoffice(excel_file, pdf_file):
    libreoffice_cmd = (
        shutil.which("soffice")
        or "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    )

    if not libreoffice_cmd or not os.path.exists(libreoffice_cmd):
        return False

    cmd = [
        libreoffice_cmd,
        "--headless",
        "--convert-to", "pdf",
        "--outdir", os.path.dirname(excel_file),
        excel_file
    ]

    subprocess.run(cmd, capture_output=True)

    generated = excel_file.replace(".xlsx", ".pdf")
    if os.path.exists(generated):
        os.replace(generated, pdf_file)
        return True

    return False

# -------------------------------------------------
# METİN PARSE ET
# -------------------------------------------------
def parse_text(text):
    """
    Form'dan gelen metni parse eder.
    Terminal script'indeki mantıkla aynı.
    """
    lines = text.strip().split("\n")
    
    data = {"rapor_no": "", "tarih": "", "yapilan_isler": []}
    mode = None

    for raw in lines:
        clean = normalize(raw)
        if clean == "rapor_no":
            mode = "rapor_no"
            continue
        if clean in ("tarih", "tarih_no"):
            mode = "tarih"
            continue
        if clean in ("yapilan_isler", "yapilan_is"):
            mode = "yapilan_isler"
            continue

        if mode == "rapor_no" and not data["rapor_no"]:
            data["rapor_no"] = raw
        elif mode == "tarih" and not data["tarih"]:
            data["tarih"] = raw
        elif mode == "yapilan_isler" and raw.startswith("-"):
            data["yapilan_isler"].append(raw[1:].strip())

    if not data["tarih"]:
        data["tarih"] = datetime.now().strftime("%d.%m.%Y")

    return data

# -------------------------------------------------
# ANA FONKSİYON
# -------------------------------------------------
def generate_report(text, photos):
    """
    Form'dan gelen metin ve fotoğrafları kullanarak PDF oluşturur.
    
    Args:
        text: Form'dan gelen rapor metni
        photos: Flask FileStorage listesi (fotoğraflar)
    
    Returns:
        PDF dosyasının yolu
    """
    # Metni parse et
    data = parse_text(text)
    safe_date = re.sub(r"[^\d.]", "", data["tarih"])
    
    # Excel şablonunu yükle
    if not os.path.exists(TEMPLATE_FILE):
        raise FileNotFoundError(f"Excel şablonu bulunamadı: {TEMPLATE_FILE}")
    
    wb = load_workbook(TEMPLATE_FILE)
    ws = wb.active

    # Excel'i doldur
    for row in ws.iter_rows():
        for cell in row:
            if cell.value == "Rapor_No":
                cell.value = data["rapor_no"]
            elif cell.value == "Tarih_No":
                cell.value = data["tarih"]
            elif isinstance(cell.value, str) and normalize(cell.value).startswith("foto"):
                cell.value = ""

    # Yapılan işleri ekle
    for i, text_item in enumerate(data["yapilan_isler"]):
        ws[f"B{8 + i}"].value = text_item

    # Fotoğrafları işle
    temp_files = []
    temp_dir = tempfile.mkdtemp()
    
    try:
        # Fotoğrafları geçici dizine kaydet
        photo_files = []
        for i, photo in enumerate(photos[:8]):
            if photo.filename:
                # Dosya uzantısını al
                ext = os.path.splitext(photo.filename)[1] or ".jpg"
                temp_photo_path = os.path.join(temp_dir, f"photo_{i+1}{ext}")
                photo.save(temp_photo_path)
                photo_files.append(temp_photo_path)

        # Fotoğrafları Excel'e ekle
        for i, photo_path in enumerate(photo_files, start=1):
            area = PHOTO_AREAS[i]
            max_w, max_h = area_pixel_size(ws, area)

            pil = PILImage.open(photo_path)
            # PIL versiyonuna göre uyumlu thumbnail
            try:
                pil.thumbnail((max_w, max_h), PILImage.Resampling.LANCZOS)
            except AttributeError:
                # Eski PIL versiyonları için
                pil.thumbnail((max_w, max_h), PILImage.LANCZOS)

            tmp = os.path.join(temp_dir, f"_tmp_{i}.jpg")
            pil.save(tmp)
            temp_files.append(tmp)

            img = Image(tmp)
            img.anchor = smart_anchor(ws, area, pil.height)
            ws.add_image(img)

        # Print area ve sayfa ayarları
        setup_print_area(ws)

        # Geçici Excel dosyası oluştur
        temp_xlsx = os.path.join(temp_dir, f"rapor-{safe_date}-{uuid.uuid4().hex[:8]}.xlsx")
        wb.save(temp_xlsx)

        # PDF oluştur
        pdf_filename = f"rapor-{safe_date}-{uuid.uuid4().hex[:8]}.pdf"
        pdf_filepath = os.path.join(OUTPUT_DIR, pdf_filename)

        # Önce xlsx2pdf ile dene
        pdf_created = False
        error_messages = []
        
        print(f"PDF oluşturma başlıyor: {temp_xlsx} -> {pdf_filepath}")
        
        try:
            result = excel_to_pdf_xlsx2pdf(temp_xlsx, pdf_filepath)
            print(f"xlsx2pdf sonucu: {result}")
            if result:
                pdf_created = True
                print("PDF başarıyla oluşturuldu (xlsx2pdf)")
        except Exception as e:
            error_msg = f"xlsx2pdf hatası: {type(e).__name__}: {str(e)}"
            error_messages.append(error_msg)
            print(error_msg)
            import traceback
            traceback.print_exc()
        
        if not pdf_created:
            print("LibreOffice deneniyor...")
            try:
                result = excel_to_pdf_libreoffice(temp_xlsx, pdf_filepath)
                print(f"LibreOffice sonucu: {result}")
                if result:
                    pdf_created = True
                    print("PDF başarıyla oluşturuldu (LibreOffice)")
            except Exception as e:
                error_msg = f"LibreOffice hatası: {type(e).__name__}: {str(e)}"
                error_messages.append(error_msg)
                print(error_msg)
        
        if not pdf_created:
            error_msg = "PDF oluşturulamadı. "
            if error_messages:
                error_msg += " ".join(error_messages)
            else:
                error_msg += "xlsx2pdf veya LibreOffice gerekli."
            print(f"PDF oluşturma başarısız: {error_msg}")
            raise Exception(error_msg)

        return pdf_filepath

    finally:
        # Geçici dosyaları temizle
        for f in temp_files:
            if os.path.exists(f):
                try:
                    os.remove(f)
                except:
                    pass
        
        # Geçici dizini temizle
        try:
            shutil.rmtree(temp_dir)
        except:
            pass
