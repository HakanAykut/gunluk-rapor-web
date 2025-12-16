from PIL import Image as PILImage
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm, mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as ReportLabImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
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
LOGO_FILE = os.path.join(BASE_DIR, "Resim1.png")

os.makedirs(OUTPUT_DIR, exist_ok=True)

# -------------------------------------------------
# TÜRKÇE FONT AYARLARI
# -------------------------------------------------
def setup_font():
    """Türkçe karakter desteği için font ayarla"""
    # Helvetica Türkçe karakterleri destekler (çoğu durumda)
    # DejaVu Sans daha iyi ama yüklü olmayabilir
    try:
        # DejaVu Sans font dosyası varsa kullan
        dejavu_path = os.path.join(BASE_DIR, "DejaVuSans.ttf")
        if os.path.exists(dejavu_path):
            pdfmetrics.registerFont(TTFont('DejaVuSans', dejavu_path))
            return 'DejaVuSans'
    except:
        pass
    
    # Helvetica kullan (Türkçe karakterler çoğu zaman çalışır)
    return 'Helvetica'

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
# EXCEL TABLO YAPISINI TAKLİT EDEN PDF OLUŞTUR
# -------------------------------------------------
def create_pdf_with_reportlab(data, photo_files, pdf_filepath, logo_path=None):
    """
    Excel'deki tablo yapısını birebir taklit ederek PDF oluşturur.
    Ana layout bir Table grid'dir (16 sütun: A-P).
    """
    try:
        # Font ayarla
        font_name = setup_font()
        
        # PDF dokümanı oluştur - A4, küçük margin'ler
        doc = SimpleDocTemplate(
            pdf_filepath,
            pagesize=A4,
            rightMargin=0.5*cm,
            leftMargin=0.5*cm,
            topMargin=0.5*cm,
            bottomMargin=0.5*cm
        )
        
        if logo_path is None:
            logo_path = LOGO_FILE
        
        styles = getSampleStyleSheet()
        
        # Excel'deki gibi sütun genişlikleri (16 sütun: A-P)
        # Toplam genişlik: A4 genişliği - margin'ler = ~19cm
        col_widths = [
            1.0*cm,  # A
            1.2*cm,  # B
            1.2*cm,  # C
            1.2*cm,  # D
            1.2*cm,  # E
            1.2*cm,  # F
            1.2*cm,  # G
            1.2*cm,  # H
            1.2*cm,  # I
            1.2*cm,  # J
            1.2*cm,  # K
            1.2*cm,  # L
            1.2*cm,  # M
            1.2*cm,  # N
            1.2*cm,  # O
            1.2*cm,  # P
        ]
        
        # Tablo verileri - Excel'deki gibi satır satır
        table_data = []
        
        # Satır 1: Boş
        table_data.append([''] * 16)
        
        # Satır 2: Header - Logo (A2-C2), Proje başlığı (D2-M2), Rapor No (N2), Tarih (O2-P2)
        row2 = [''] * 16
        
        # Logo - A2-C2 (birleştirilmiş görünüm için A2'ye koy, C2'ye kadar boş)
        if logo_path and os.path.exists(logo_path):
            try:
                logo_img = ReportLabImage(logo_path, width=3*cm, height=1.5*cm, kind='proportional')
                row2[0] = logo_img  # A sütunu
            except Exception as e:
                print(f"Logo yüklenemedi: {e}")
                row2[0] = Paragraph("LOGO", ParagraphStyle('Logo', parent=styles['Normal'], 
                                                          fontSize=8, fontName=font_name))
        else:
            row2[0] = Paragraph("LOGO", ParagraphStyle('Logo', parent=styles['Normal'], 
                                                      fontSize=8, fontName=font_name))
        
        # Proje başlığı - D2-M2 (birleştirilmiş görünüm için D2'ye koy)
        project_title = data.get("proje_basligi", "FETİHTEPE MERKEZ CAMİ'İ GÜÇLENDİRME VE YENİLEME PROJESİ")
        project_style = ParagraphStyle(
            'ProjectTitle',
            parent=styles['Normal'],
            fontSize=12,
            fontName=f'{font_name}-Bold',
            textColor=colors.black,
            alignment=TA_CENTER
        )
        row2[3] = Paragraph(project_title, project_style)  # D sütunu
        
        # Rapor No - N2
        rapor_no_text = f"<b>Günlük Rapor No:</b><br/>{data.get('rapor_no', '')}"
        header_style = ParagraphStyle(
            'Header',
            parent=styles['Normal'],
            fontSize=9,
            fontName=font_name,
            alignment=TA_RIGHT
        )
        row2[13] = Paragraph(rapor_no_text, header_style)  # N sütunu
        
        table_data.append(row2)
        
        # Satır 3: Tarih (N3)
        row3 = [''] * 16
        tarih_text = f"<b>Tarih:</b><br/>{data.get('tarih', '')}"
        row3[13] = Paragraph(tarih_text, header_style)  # N sütunu
        table_data.append(row3)
        
        # Satır 4: Boş
        table_data.append([''] * 16)
        
        # Satır 5: "GÜNLÜK FAALİYET RAPORU" gri bar (B5-P5)
        row5 = [''] * 16
        daily_report_style = ParagraphStyle(
            'DailyReport',
            parent=styles['Normal'],
            fontSize=12,
            fontName=f'{font_name}-Bold',
            textColor=colors.white,
            alignment=TA_CENTER
        )
        row5[1] = Paragraph("GÜNLÜK FAALİYET RAPORU", daily_report_style)  # B sütunu
        table_data.append(row5)
        
        # Satır 6: Boş
        table_data.append([''] * 16)
        
        # Satır 7: "YAPILAN İŞLER:" başlığı (B7)
        row7 = [''] * 16
        works_title_style = ParagraphStyle(
            'WorksTitle',
            parent=styles['Normal'],
            fontSize=11,
            fontName=f'{font_name}-Bold',
            alignment=TA_LEFT
        )
        row7[1] = Paragraph("YAPILAN İŞLER:", works_title_style)  # B sütunu
        table_data.append(row7)
        
        # Satır 8-13: Yapılan işler listesi (B8-B13)
        work_item_style = ParagraphStyle(
            'WorkItem',
            parent=styles['Normal'],
            fontSize=9,
            fontName=font_name,
            alignment=TA_LEFT,
            leftIndent=0.3*cm,
            leading=11
        )
        yapilan_isler = data.get("yapilan_isler", [])
        for i in range(6):
            row = [''] * 16
            if i < len(yapilan_isler):
                row[1] = Paragraph(f"• {yapilan_isler[i]}", work_item_style)  # B sütunu
            else:
                # Boş satır - altı çizgili görünüm için
                row[1] = Paragraph("_", ParagraphStyle('EmptyLine', parent=styles['Normal'], 
                                                      fontSize=8, fontName=font_name, 
                                                      textColor=colors.grey))
            table_data.append(row)
        
        # Satır 14-27: Boş satırlar (altı çizgili görünüm için)
        for _ in range(14):
            row = [''] * 16
            row[1] = Paragraph("_", ParagraphStyle('EmptyLine', parent=styles['Normal'], 
                                                  fontSize=8, fontName=font_name, 
                                                  textColor=colors.grey))
            table_data.append(row)
        
        # Satır 28: "İMALAT FOTOĞRAFLARI" başlığı (B28-P28)
        row28 = [''] * 16
        photos_title_style = ParagraphStyle(
            'PhotosTitle',
            parent=styles['Normal'],
            fontSize=11,
            fontName=f'{font_name}-Bold',
            alignment=TA_CENTER
        )
        row28[1] = Paragraph("İMALAT FOTOĞRAFLARI", photos_title_style)  # B sütunu
        table_data.append(row28)
        
        # Fotoğraflar için satır sayısı hesapla
        # Her fotoğraf yaklaşık 18 satır yüksekliğinde (Excel'deki gibi)
        # Fotoğraf pozisyonları: B29-I47, J29-P47, B52-I70, J52-P70, B75-I93, J75-P93, B98-I116, J98-P116
        photo_start_rows = [28, 28, 51, 51, 74, 74, 97, 97]  # Satır indeksleri (0-based)
        photo_start_cols = [1, 9, 1, 9, 1, 9, 1, 9]  # B=1, J=9
        photo_end_cols = [8, 15, 8, 15, 8, 15, 8, 15]  # I=8, P=15
        
        # Fotoğraf boyutları
        photo_box_width = sum(col_widths[1:9])  # B-I arası genişlik
        photo_box_height = 4.5*cm  # Yaklaşık 18 satır yüksekliği
        
        # Fotoğrafları hazırla ve tabloya ekle
        for photo_idx in range(8):
            start_row = photo_start_rows[photo_idx]
            start_col = photo_start_cols[photo_idx]
            end_col = photo_end_cols[photo_idx]
            
            # Eksik satırları ekle
            while len(table_data) <= start_row + 18:
                table_data.append([''] * 16)
            
            if photo_idx < len(photo_files):
                photo_path = photo_files[photo_idx]
                try:
                    # Fotoğrafı yükle ve boyutlandır
                    pil_img = PILImage.open(photo_path)
                    pil_img.thumbnail((int(photo_box_width*2), int(photo_box_height*2)), 
                                    PILImage.Resampling.LANCZOS)
                    
                    # Geçici dosyaya kaydet
                    temp_photo = os.path.join(os.path.dirname(photo_path), 
                                             f"_pdf_temp_{os.path.basename(photo_path)}")
                    pil_img.save(temp_photo)
                    
                    # ReportLab Image oluştur
                    img = ReportLabImage(temp_photo, width=photo_box_width, 
                                       height=photo_box_height, kind='proportional')
                    
                    # Fotoğraf ve etiket için iç tablo
                    photo_cell_data = [
                        [img],
                        [Paragraph(f"FOTO-{photo_idx + 1}", 
                                  ParagraphStyle('PhotoLabel', parent=styles['Normal'], 
                                               fontSize=8, alignment=TA_CENTER, fontName=font_name))]
                    ]
                    photo_cell = Table(photo_cell_data, 
                                     colWidths=[photo_box_width], 
                                     rowHeights=[photo_box_height, 0.3*cm])
                    photo_cell.setStyle(TableStyle([
                        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                        ('VALIGN', (0, 0), (0, 0), 'MIDDLE'),
                        ('LEFTPADDING', (0, 0), (-1, -1), 2),
                        ('RIGHTPADDING', (0, 0), (-1, -1), 2),
                        ('TOPPADDING', (0, 0), (0, 0), 2),
                        ('BOTTOMPADDING', (0, 0), (0, 0), 2),
                    ]))
                    
                    # Fotoğrafı tabloya ekle (birleştirilmiş hücre simülasyonu)
                    # İlk hücreye fotoğrafı koy, diğer hücreleri boş bırak
                    table_data[start_row][start_col] = photo_cell
                    
                except Exception as e:
                    print(f"Fotoğraf yüklenemedi {photo_path}: {e}")
        
        # Ana tablo oluştur
        main_table = Table(table_data, colWidths=col_widths, repeatRows=0)
        
        # Tablo stil ayarları
        table_style = [
            # Grid çizgileri (Excel'deki gibi)
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            # Hizalama
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            # Padding
            ('LEFTPADDING', (0, 0), (-1, -1), 2),
            ('RIGHTPADDING', (0, 0), (-1, -1), 2),
            ('TOPPADDING', (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
            # Gri bar: "GÜNLÜK FAALİYET RAPORU" (Satır 5, B-P)
            ('BACKGROUND', (1, 4), (15, 4), colors.HexColor('#808080')),
            # Gri bar: "İMALAT FOTOĞRAFLARI" (Satır 28, B-P)
            ('BACKGROUND', (1, 27), (15, 27), colors.HexColor('#D3D3D3')),
            # Proje başlığı ortalanmış (D2-M2)
            ('ALIGN', (3, 1), (12, 1), 'CENTER'),
            # Rapor No ve Tarih sağa hizalı (N2-N3)
            ('ALIGN', (13, 1), (13, 2), 'RIGHT'),
        ]
        
        main_table.setStyle(TableStyle(table_style))
        
        # Story oluştur
        story = []
        story.append(main_table)
        
        # Footer
        story.append(Spacer(1, 0.2*cm))
        footer_style = ParagraphStyle(
            'Footer',
            parent=styles['Normal'],
            fontSize=7,
            fontName=font_name,
            textColor=colors.HexColor('#666666'),
            alignment=TA_LEFT
        )
        story.append(Paragraph(
            "İşbu dokümanda HASSAS bilgi bulunmamaktadır. / This document does not contain SENSITIVE information.",
            footer_style
        ))
        
        # PDF'i oluştur
        doc.build(story)
        
        # Geçici dosyaları temizle
        for photo_path in photo_files:
            temp_photo = os.path.join(os.path.dirname(photo_path), 
                                     f"_pdf_temp_{os.path.basename(photo_path)}")
            if os.path.exists(temp_photo):
                try:
                    os.remove(temp_photo)
                except:
                    pass
        
        if os.path.exists(pdf_filepath) and os.path.getsize(pdf_filepath) > 0:
            print(f"PDF başarıyla oluşturuldu: {pdf_filepath}")
            return True
        return False
        
    except Exception as e:
        print(f"PDF oluşturma hatası: {type(e).__name__}: {e}")
        import traceback
        traceback.print_exc()
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
def generate_report(data, photos):
    """
    Form'dan gelen data ve fotoğrafları kullanarak PDF oluşturur.
    
    Args:
        data: Dict - {"tarih": "...", "rapor_no": "...", "yapilan_isler": [...]}
        photos: Flask FileStorage listesi (fotoğraflar)
    
    Returns:
        PDF dosyasının yolu
    """
    safe_date = re.sub(r"[^\d.]", "", data["tarih"])
    
    # Geçici dizin oluştur
    temp_dir = tempfile.mkdtemp()
    temp_files = []
    
    try:
        # Fotoğrafları geçici dizine kaydet
        photo_files = []
        for i, photo in enumerate(photos[:8]):
            if photo.filename:
                ext = os.path.splitext(photo.filename)[1] or ".jpg"
                temp_photo_path = os.path.join(temp_dir, f"photo_{i+1}{ext}")
                photo.save(temp_photo_path)
                photo_files.append(temp_photo_path)

        # PDF oluştur
        pdf_filename = f"rapor-{safe_date}-{uuid.uuid4().hex[:8]}.pdf"
        pdf_filepath = os.path.join(OUTPUT_DIR, pdf_filename)
        
        print(f"PDF oluşturma başlıyor: {pdf_filepath}")
        
        pdf_created = create_pdf_with_reportlab(data, photo_files, pdf_filepath)
        
        if not pdf_created:
            raise Exception("PDF oluşturulamadı.")

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
