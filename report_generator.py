from PIL import Image as PILImage
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm, mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as ReportLabImage, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
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
# Excel şablonu formatında, tek sayfalık
# -------------------------------------------------
def create_pdf_with_reportlab(data, photo_files, pdf_filepath):
    """
    ReportLab kullanarak direkt PDF oluşturur.
    Excel şablonu formatında, tek sayfalık düzen.
    Türkçe karakter desteği ile.
    """
    try:
        # Türkçe karakter desteği için Helvetica kullan (ReportLab'in varsayılan fontu Türkçe destekler)
        # Eğer sorun olursa, DejaVu Sans gibi Unicode font kullanabiliriz
        
        # PDF dokümanı oluştur - Excel şablonuna benzer margin'ler
        doc = SimpleDocTemplate(
            pdf_filepath,
            pagesize=A4,
            rightMargin=0.5*cm,
            leftMargin=0.5*cm,
            topMargin=0.5*cm,
            bottomMargin=0.5*cm
        )
        
        # Stil tanımlamaları - Excel formatına uygun
        styles = getSampleStyleSheet()
        
        # Türkçe karakter desteği için encoding ayarı
        # ReportLab Paragraph HTML benzeri format kullanır ve UTF-8 destekler
        
        # Proje başlığı stili (D2 hücresi - üstte sol)
        project_style = ParagraphStyle(
            'ProjectStyle',
            parent=styles['Normal'],
            fontSize=10,
            textColor=colors.HexColor('#000000'),
            alignment=TA_LEFT,
            spaceAfter=5,
            fontName='Helvetica'
        )
        
        # Rapor No/Tarih stili (N2, N3 hücreleri - üstte sağ)
        header_style = ParagraphStyle(
            'HeaderStyle',
            parent=styles['Normal'],
            fontSize=10,
            textColor=colors.HexColor('#000000'),
            alignment=TA_LEFT,
            spaceAfter=3,
            fontName='Helvetica'
        )
        
        # Ana başlık stili (B5 hücresi)
        main_title_style = ParagraphStyle(
            'MainTitleStyle',
            parent=styles['Normal'],
            fontSize=14,
            textColor=colors.HexColor('#000000'),
            alignment=TA_CENTER,
            spaceAfter=10,
            fontName='Helvetica-Bold'
        )
        
        # İmalatın Cinsi başlığı (B7 hücresi)
        section_style = ParagraphStyle(
            'SectionStyle',
            parent=styles['Normal'],
            fontSize=11,
            textColor=colors.HexColor('#000000'),
            alignment=TA_LEFT,
            spaceAfter=5,
            fontName='Helvetica-Bold'
        )
        
        # Yapılan işler stili (B8-B11 hücreleri)
        work_item_style = ParagraphStyle(
            'WorkItemStyle',
            parent=styles['Normal'],
            fontSize=10,
            textColor=colors.HexColor('#000000'),
            alignment=TA_LEFT,
            spaceAfter=4,
            leftIndent=10,
            fontName='Helvetica'
        )
        
        # PDF içeriği
        story = []
        
        # Üst kısım: Proje adı (D2) ve Rapor No/Tarih (N2, N3)
        # Excel'de D2 sol tarafta, N2-N3 sağ tarafta
        header_table_data = []
        left_col = []
        right_col = []
        
        # Sol kolon: Proje adı (D2) - Türkçe karakterler için doğrudan kullan
        project_name = "FETİHTEPE CAMİİ GÜÇLENDİRME VE YENİLEME PROJESİ"
        left_col.append(Paragraph(project_name, project_style))
        left_col.append(Spacer(1, 0.2*cm))
        
        # Sağ kolon: Rapor No ve Tarih (N2, N3)
        if data["rapor_no"]:
            right_col.append(Paragraph(f"Günlük Rapor No: {data['rapor_no']}", header_style))
        if data["tarih"]:
            right_col.append(Paragraph(f"Tarih: {data['tarih']}", header_style))
        
        # Header tablosu
        header_table = Table([[left_col, right_col]], colWidths=[12*cm, 6*cm])
        header_table.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('ALIGN', (0, 0), (0, -1), 'LEFT'),
            ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
        ]))
        story.append(header_table)
        story.append(Spacer(1, 0.3*cm))
        
        # Ana başlık: GÜNLÜK FAALİYET RAPORU (B5)
        story.append(Paragraph("GÜNLÜK FAALİYET RAPORU", main_title_style))
        story.append(Spacer(1, 0.4*cm))
        
        # İMALATIN CİNSİ başlığı (B7)
        story.append(Paragraph("İMALATIN CİNSİ", section_style))
        story.append(Spacer(1, 0.2*cm))
        
        # Yapılan işler listesi (B8-B11) - Excel'de "-" ile başlıyor
        for i, is_item in enumerate(data["yapilan_isler"]):
            story.append(Paragraph(f"- {is_item}", work_item_style))
        
        story.append(Spacer(1, 0.3*cm))
        
        # Fotoğraflar - Excel'deki gibi 2x4 grid (8 fotoğraf)
        if photo_files:
            # Fotoğrafları Excel'deki gibi 2 sütun, 4 satır grid'de göster
            photos_per_row = 2
            # Excel'de fotoğraflar B29-I47, J29-P47 gibi alanlarda
            # A4 sayfasında yaklaşık 8cm x 5cm per fotoğraf
            photo_width = 8*cm
            photo_height = 5*cm
            
            for i in range(0, len(photo_files), photos_per_row):
                row_photos = photo_files[i:i+photos_per_row]
                row_data = []
                
                for photo_path in row_photos:
                    try:
                        # Fotoğrafı boyutlandır
                        pil_img = PILImage.open(photo_path)
                        # Aspect ratio'yu koru
                        pil_img.thumbnail((int(photo_width*2), int(photo_height*2)), PILImage.Resampling.LANCZOS)
                        
                        # Geçici dosyaya kaydet
                        temp_photo = os.path.join(os.path.dirname(photo_path), f"_pdf_{os.path.basename(photo_path)}")
                        pil_img.save(temp_photo)
                        
                        img = ReportLabImage(temp_photo, width=photo_width, height=photo_height, kind='proportional')
                        row_data.append(img)
                    except Exception as e:
                        print(f"Fotoğraf yüklenemedi {photo_path}: {e}")
                        row_data.append(Paragraph("Fotoğraf yüklenemedi", work_item_style))
                
                # Eksik sütunları doldur
                while len(row_data) < photos_per_row:
                    row_data.append(Spacer(1, 1))
                
                # Tablo oluştur - Excel'deki gibi 2 sütun
                photo_table = Table([row_data], colWidths=[photo_width]*photos_per_row)
                photo_table.setStyle(TableStyle([
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                    ('LEFTPADDING', (0, 0), (-1, -1), 2),
                    ('RIGHTPADDING', (0, 0), (-1, -1), 2),
                    ('TOPPADDING', (0, 0), (-1, -1), 2),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
                ]))
                story.append(photo_table)
                story.append(Spacer(1, 0.2*cm))
        
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
