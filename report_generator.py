from PIL import Image as PILImage
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as ReportLabImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from datetime import datetime
import os
import re
import tempfile
import uuid
import shutil

# -------------------------------------------------
# YOLLAR
# -------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "generated_pdfs")

os.makedirs(OUTPUT_DIR, exist_ok=True)

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
# REPORTLAB İLE PDF OLUŞTUR (Render-compatible)
# -------------------------------------------------
def create_pdf_with_reportlab(data, photo_files, pdf_filepath):
    """
    ReportLab kullanarak direkt PDF oluşturur.
    Excel → PDF dönüştürme yapmaz, Render-compatible.
    """
    try:
        # PDF dokümanı oluştur
        doc = SimpleDocTemplate(
            pdf_filepath,
            pagesize=A4,
            rightMargin=1*cm,
            leftMargin=1*cm,
            topMargin=1*cm,
            bottomMargin=1*cm
        )
        
        # Stil tanımlamaları
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Heading1'],
            fontSize=16,
            textColor=colors.HexColor('#000000'),
            alignment=TA_CENTER,
            spaceAfter=20
        )
        
        normal_style = ParagraphStyle(
            'CustomNormal',
            parent=styles['Normal'],
            fontSize=11,
            textColor=colors.HexColor('#000000'),
            alignment=TA_LEFT,
            spaceAfter=12
        )
        
        # PDF içeriği
        story = []
        
        # Başlık: Rapor No ve Tarih
        if data["rapor_no"]:
            story.append(Paragraph(f"<b>Rapor No:</b> {data['rapor_no']}", normal_style))
        if data["tarih"]:
            story.append(Paragraph(f"<b>Tarih:</b> {data['tarih']}", normal_style))
        
        story.append(Spacer(1, 0.5*cm))
        
        # Yapılan İşler başlığı
        story.append(Paragraph("<b>Yapılan İşler:</b>", normal_style))
        story.append(Spacer(1, 0.3*cm))
        
        # Yapılan işler listesi
        for i, is_item in enumerate(data["yapilan_isler"], 1):
            story.append(Paragraph(f"{i}. {is_item}", normal_style))
        
        story.append(Spacer(1, 0.5*cm))
        
        # Fotoğraflar
        if photo_files:
            story.append(Paragraph("<b>Fotoğraflar:</b>", normal_style))
            story.append(Spacer(1, 0.3*cm))
            
            # Fotoğrafları 2x4 grid'de göster (8 fotoğraf için)
            photos_per_row = 2
            photo_width = (A4[0] - 2*cm) / photos_per_row - 0.5*cm
            photo_height = 6*cm
            
            for i in range(0, len(photo_files), photos_per_row):
                row_photos = photo_files[i:i+photos_per_row]
                row_data = []
                
                for photo_path in row_photos:
                    try:
                        # Fotoğrafı boyutlandır
                        pil_img = PILImage.open(photo_path)
                        pil_img.thumbnail((int(photo_width*2), int(photo_height*2)), PILImage.Resampling.LANCZOS)
                        
                        # Geçici dosyaya kaydet
                        temp_photo = os.path.join(os.path.dirname(photo_path), f"_pdf_{os.path.basename(photo_path)}")
                        pil_img.save(temp_photo)
                        
                        img = ReportLabImage(temp_photo, width=photo_width, height=photo_height)
                        row_data.append(img)
                    except Exception as e:
                        print(f"Fotoğraf yüklenemedi {photo_path}: {e}")
                        row_data.append(Paragraph("Fotoğraf yüklenemedi", normal_style))
                
                # Eksik sütunları doldur
                while len(row_data) < photos_per_row:
                    row_data.append(Spacer(1, 1))
                
                # Tablo oluştur
                photo_table = Table([row_data], colWidths=[photo_width]*photos_per_row)
                photo_table.setStyle(TableStyle([
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ]))
                story.append(photo_table)
                story.append(Spacer(1, 0.3*cm))
        
        # PDF'i oluştur
        doc.build(story)
        
        if os.path.exists(pdf_filepath) and os.path.getsize(pdf_filepath) > 0:
            print(f"PDF başarıyla oluşturuldu (ReportLab): {pdf_filepath}")
            return True
        return False
        
    except Exception as e:
        print(f"ReportLab PDF oluşturma hatası: {type(e).__name__}: {e}")
        import traceback
        traceback.print_exc()
        return False

# -------------------------------------------------
# ESKİ KODLAR KALDIRILDI - Artık Excel → PDF dönüştürme yapmıyoruz
# Render-compatible: Sadece ReportLab kullanıyoruz
# Tüm excel_to_pdf_xlsx2pdf ve excel_to_pdf_libreoffice fonksiyonları kaldırıldı
# -------------------------------------------------

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
    
    # Geçici dizin oluştur
    temp_dir = tempfile.mkdtemp()
    temp_files = []
    
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

        # PDF oluştur - ReportLab ile direkt (Render-compatible)
        # Excel kullanmıyoruz, sadece ReportLab ile PDF oluşturuyoruz
        pdf_filename = f"rapor-{safe_date}-{uuid.uuid4().hex[:8]}.pdf"
        pdf_filepath = os.path.join(OUTPUT_DIR, pdf_filename)
        
        print(f"PDF oluşturma başlıyor (ReportLab): {pdf_filepath}")
        
        # ReportLab ile PDF oluştur
        pdf_created = create_pdf_with_reportlab(data, photo_files, pdf_filepath)
        
        if not pdf_created:
            raise Exception("PDF oluşturulamadı. ReportLab hatası.")

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
            import shutil
            shutil.rmtree(temp_dir)
        except:
            pass
