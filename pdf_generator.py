from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from pdf_layout import (
    PAGE_WIDTH, PAGE_HEIGHT,
    HEADER_HEIGHT, MARGIN_LEFT, MARGIN_RIGHT, MARGIN_TOP, MARGIN_BOTTOM,
    HEADER_COL1_WIDTH, HEADER_COL2_WIDTH, HEADER_COL3_WIDTH,
    HEADER_TABLE_CELL_HEIGHT,
    BAND_HEIGHT,
    WORKS_TITLE_HEIGHT, WORKS_ROW_HEIGHT, WORKS_MAX_ROWS,
    PHOTO_GRID_COLS, PHOTO_GRID_ROWS, PHOTOS_PER_PAGE, PHOTO_LABEL_HEIGHT,
    FONT_SIZE_TITLE, FONT_SIZE_HEADER, FONT_SIZE_NORMAL, FONT_SIZE_SMALL,
    setup_fonts, draw_box, draw_text, draw_text_multiline, draw_image_fit,
    calculate_text_height
)
import os
import tempfile
import shutil
import uuid
import re

# Base dizin
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "generated_pdfs")
LOGO_FILE = os.path.join(BASE_DIR, "Resim1.png")

os.makedirs(OUTPUT_DIR, exist_ok=True)

def generate_pdf(data, photo_files, pdf_filepath, logo_path=None, base_dir=None):
    """
    Canvas ile manuel koordinatlarla PDF oluşturur.
    
    Args:
        data: Dict - {"tarih": "...", "rapor_no": "...", "yapilan_isler": [...], "proje_basligi": "..."}
        photo_files: List[str] - Fotoğraf dosya yolları
        pdf_filepath: str - Çıktı PDF yolu
        logo_path: str - Logo dosya yolu (opsiyonel)
        base_dir: str - Proje base dizini (font yükleme için)
    
    Returns:
        bool - Başarılı ise True
    """
    try:
        if base_dir is None:
            base_dir = BASE_DIR
        
        # Fontları yükle
        font_regular, font_bold = setup_fonts(base_dir)
        
        if logo_path is None:
            logo_path = LOGO_FILE
        
        # Canvas oluştur
        c = canvas.Canvas(pdf_filepath, pagesize=A4)
        
        # İlk sayfa koordinatları
        current_y = PAGE_HEIGHT - MARGIN_TOP
        photo_index = 0
        total_photos = len(photo_files)
        page_num = 0
        
        # ============================================================
        # İLK SAYFA: HEADER VE İÇERİK
        # ============================================================
        # HEADER (3 kolonlu) - sağdan da margin var, sol margin kadar
        header_y = current_y
        header_x = MARGIN_LEFT
        # Header'ın toplam genişliği: sayfa genişliği - sol margin - sağ margin
        header_total_width = PAGE_WIDTH - MARGIN_LEFT - MARGIN_RIGHT
        # Kolon genişliklerini header toplam genişliğine göre hesapla
        header_col1_width = header_total_width * 0.25
        header_col2_width = header_total_width * 0.50
        header_col3_width = header_total_width * 0.25
        
        # Sol kolon: Logo
        col1_x = header_x
        col1_y = header_y - HEADER_HEIGHT
        draw_box(c, col1_x, col1_y, header_col1_width, HEADER_HEIGHT)
        
        if logo_path and os.path.exists(logo_path):
            # Logo ortalanmış - alanı daha iyi kullan, fotoğraftaki gibi büyük
            logo_padding = 0.1*cm  # Çok az padding
            logo_width = header_col1_width - 2*logo_padding
            logo_height = HEADER_HEIGHT - 2*logo_padding
            logo_x = col1_x + logo_padding
            logo_y = col1_y + logo_padding
            draw_image_fit(c, logo_x, logo_y, logo_width, logo_height, logo_path)
        
        # Orta kolon: Proje başlığı
        col2_x = col1_x + header_col1_width
        col2_y = col1_y
        draw_box(c, col2_x, col2_y, header_col2_width, HEADER_HEIGHT)
        
        project_title = data.get("proje_basligi", "FETİHTEPE MERKEZ CAMİ'İ GÜÇLENDİRME VE YENİLEME PROJESİ")
        # Metni orta kolon genişliğine sığdır ve ortala
        title_max_width = header_col2_width - 0.3*cm
        
        # Metni kelimelere böl ve satırlara ayır
        words = project_title.split()
        lines = []
        current_line = ""
        
        for word in words:
            test_line = current_line + (" " if current_line else "") + word
            if c.stringWidth(test_line, font_bold, FONT_SIZE_TITLE) <= title_max_width:
                current_line = test_line
            else:
                if current_line:
                    lines.append(current_line)
                current_line = word
        if current_line:
            lines.append(current_line)
        
        # Metni dikey olarak ortala
        total_text_height = len(lines) * FONT_SIZE_TITLE * 1.3
        start_y = col2_y + HEADER_HEIGHT / 2 + total_text_height / 2 - FONT_SIZE_TITLE * 1.3
        
        # Her satırı ortala ve çiz
        for i, line in enumerate(lines):
            line_y = start_y - i * FONT_SIZE_TITLE * 1.3
            draw_text(c, col2_x + header_col2_width / 2, line_y, line, font_bold, FONT_SIZE_TITLE,
                     alignment='center', bold=True)
        
        # Sağ kolon: Rapor bilgileri (2x2 tablo) - açık gri arka plan
        col3_x = col2_x + header_col2_width
        col3_y = col1_y
        draw_box(c, col3_x, col3_y, header_col3_width, HEADER_HEIGHT,
                fill_color=colors.HexColor('#F5F5F5'))
        
        # İç tablo çizgileri
        cell_width = header_col3_width / 2
        # Dikey çizgi - sağdan taşmaması için header_col3_width kullan
        c.setStrokeColor(colors.black)
        c.setLineWidth(0.5)
        c.line(col3_x + cell_width, col3_y, col3_x + cell_width, col3_y + HEADER_HEIGHT)
        # Yatay çizgi - sağdan taşmaması için header_col3_width kullan
        c.line(col3_x, col3_y + HEADER_TABLE_CELL_HEIGHT, col3_x + header_col3_width, col3_y + HEADER_TABLE_CELL_HEIGHT)
        
        # Hücre içerikleri - küçük fontlar kullan, değerler sağdaki karşıdaki kutucuklarda
        cell_padding = 0.08*cm
        header_small_font = 6  # Küçük font boyutu (küçültüldü)
        
        # Satır 1, Sol hücre: "Günlük Rapor No" - küçük font
        draw_text(c, col3_x + cell_padding, col3_y + HEADER_TABLE_CELL_HEIGHT + HEADER_TABLE_CELL_HEIGHT/2 - header_small_font/3, 
                 "Günlük Rapor No", font_regular, header_small_font, alignment='left')
        # Satır 1, Sağ hücre: Rapor No değeri (karşıdaki kutucukta, ortalanmış, küçük font)
        rapor_no_text = str(data.get("rapor_no", ""))
        draw_text(c, col3_x + cell_width + cell_width/2, col3_y + HEADER_TABLE_CELL_HEIGHT + HEADER_TABLE_CELL_HEIGHT/2 - header_small_font/3,
                 rapor_no_text, font_regular, header_small_font, alignment='center')
        
        # Satır 2, Sol hücre: "Tarih" - küçük font
        draw_text(c, col3_x + cell_padding, col3_y + HEADER_TABLE_CELL_HEIGHT/2 - header_small_font/3,
                 "Tarih", font_regular, header_small_font, alignment='left')
        # Satır 2, Sağ hücre: Tarih değeri (karşıdaki kutucukta, ortalanmış, küçük font)
        tarih_text = data.get("tarih", "")
        draw_text(c, col3_x + cell_width + cell_width/2, col3_y + HEADER_TABLE_CELL_HEIGHT/2 - header_small_font/3,
                 tarih_text, font_regular, header_small_font, alignment='center')
        
        current_y = col1_y  # Header'dan sonra boşluk yok, bitişik
        
        # ============================================================
        # GRI BANT: "GÜNLÜK FAALİYET RAPORU"
        # Header tablosu ile sağdan hizalı olmalı
        # ============================================================
        band_y = current_y - BAND_HEIGHT
        # Header'ın sağ kenarı: col3_x + header_col3_width
        # Gri band header ile sağdan hizalı olmalı (aynı genişlikte)
        header_right_edge = col3_x + header_col3_width
        band_width = header_total_width
        draw_box(c, MARGIN_LEFT, band_y, band_width, BAND_HEIGHT,
                fill_color=colors.HexColor('#D3D3D3'))
        # Metni gri bandın ortasına yerleştir
        band_center_x = MARGIN_LEFT + band_width / 2
        draw_text(c, band_center_x, band_y + BAND_HEIGHT / 2 - FONT_SIZE_HEADER / 3,
                 "GÜNLÜK FAALİYET RAPORU", font_bold, FONT_SIZE_HEADER,
                 alignment='center', bold=True)
        current_y = band_y  # Bitişik, boşluk yok
        
        # ============================================================
        # YAPILAN İŞLER BÖLÜMÜ
        # Header ile aynı genişlikte olmalı
        # ============================================================
        # Başlık - açık gri arka plan
        works_title_y = current_y - WORKS_TITLE_HEIGHT
        draw_box(c, MARGIN_LEFT, works_title_y, band_width, WORKS_TITLE_HEIGHT,
                fill_color=colors.HexColor('#F5F5F5'))
        draw_text(c, MARGIN_LEFT + 0.1*cm, works_title_y + WORKS_TITLE_HEIGHT / 2 - FONT_SIZE_NORMAL / 3,
                 "YAPILAN İŞLER:", font_bold, FONT_SIZE_NORMAL, bold=True, alignment='left')
        current_y = works_title_y  # Bitişik, boşluk yok
        
        # İşler listesi - dinamik yükseklik ile
        yapilan_isler = data.get("yapilan_isler", [])
        
        # İşler için alan genişliği (padding dahil)
        works_text_width = band_width - 0.2*cm  # Sol ve sağ padding
        works_text_x = MARGIN_LEFT + 0.1*cm
        
        # Maksimum toplam yükseklik (genel düzeni korumak için)
        max_total_height = WORKS_ROW_HEIGHT * WORKS_MAX_ROWS
        min_row_height = WORKS_ROW_HEIGHT  # Minimum satır yüksekliği
        
        # İşleri dinamik olarak çiz
        current_works_y = current_y
        total_used_height = 0
        
        for i, is_text in enumerate(yapilan_isler):
            # Maksimum alanı aşmamak için kontrol
            remaining_height = max_total_height - total_used_height
            if remaining_height < min_row_height:
                # Alan doldu, kalan işleri atla
                break
            
            # Madde işareti ile metin
            full_text = f"• {is_text}"
            
            # Hücre içi padding (üst ve alt)
            cell_padding_vertical = 0.15*cm  # Üst ve alttan boşluk
            
            # Metnin yüksekliğini hesapla (padding olmadan)
            text_height = calculate_text_height(c, full_text, font_regular, FONT_SIZE_NORMAL, works_text_width)
            # Hücre yüksekliği = metin yüksekliği + üst padding + alt padding
            row_height = max(text_height + 2*cell_padding_vertical, min_row_height)
            
            # Eğer kalan alan yetersizse, mevcut alanı kullan
            if total_used_height + row_height > max_total_height:
                row_height = remaining_height
            
            # Hücreyi çiz
            row_y = current_works_y - row_height
            draw_box(c, MARGIN_LEFT, row_y, band_width, row_height)
            
            # Metni word wrap ile hücrenin içine çiz
            # Metnin gerçek yüksekliğini hesapla (satır sayısı * satır yüksekliği)
            line_height = FONT_SIZE_NORMAL * 1.2
            
            # Gelişmiş satır bölme: Çok uzun kelimeleri de karakter karakter böler
            def get_wrapped_lines(text, font, size, width):
                words = text.split(' ')
                lines = []
                current_line = ''
                for word in words:
                    # Kelime tek başına genişlikten büyükse karakter karakter böl
                    if c.stringWidth(word, font, size) > width:
                        if current_line:
                            lines.append(current_line)
                            current_line = ''
                        
                        temp_word = ''
                        for char in word:
                            if c.stringWidth(temp_word + char, font, size) <= width:
                                temp_word += char
                            else:
                                lines.append(temp_word)
                                temp_word = char
                        current_line = temp_word
                    else:
                        test_line = (current_line + ' ' + word).strip()
                        if c.stringWidth(test_line, font, size) <= width:
                            current_line = test_line
                        else:
                            if current_line:
                                lines.append(current_line)
                            current_line = word
                if current_line:
                    lines.append(current_line)
                return lines

            wrapped_lines = get_wrapped_lines(full_text, font_regular, FONT_SIZE_NORMAL, works_text_width)
            
            # Metni hücrenin üstünden padding ile başlat
            # ReportLab'de y pozisyonu baseline'dır, bu yüzden font boyutunun bir kısmını ekliyoruz
            text_start_y = row_y + row_height - cell_padding_vertical - (FONT_SIZE_NORMAL * 0.3)
            
            # Metni satır satır çiz (yukarıdan aşağıya)
            for line in wrapped_lines:
                draw_text(c, works_text_x, text_start_y, line, font_regular, FONT_SIZE_NORMAL, alignment='left')
                text_start_y -= line_height
            
            # Sonraki iş için y pozisyonunu güncelle
            current_works_y = row_y
            total_used_height += row_height
        
        # Boş satırları çiz (kalan alan için)
        remaining_height = max_total_height - total_used_height
        empty_rows_count = int(remaining_height / min_row_height)
        
        for i in range(empty_rows_count):
            row_y = current_works_y - min_row_height
            draw_box(c, MARGIN_LEFT, row_y, band_width, min_row_height)
            
            # Boş satır - altı çizgili görünüm (sadece boş hücrelerde)
            c.setStrokeColor(colors.grey)
            c.setLineWidth(0.3)
            line_y = row_y + min_row_height / 2
            c.line(works_text_x, line_y, 
                   MARGIN_LEFT + band_width - 0.1*cm, line_y)
            
            current_works_y = row_y
        
        current_y = current_works_y  # Bitişik, boşluk yok
        
        # ============================================================
        # GRI BANT: "İMALAT FOTOĞRAFLARI"
        # Header ile aynı genişlikte olmalı
        # ============================================================
        photos_band_y = current_y - BAND_HEIGHT
        draw_box(c, MARGIN_LEFT, photos_band_y, band_width, BAND_HEIGHT,
                fill_color=colors.HexColor('#D3D3D3'))
        photos_band_center_x = MARGIN_LEFT + band_width / 2
        draw_text(c, photos_band_center_x, photos_band_y + BAND_HEIGHT / 2 - FONT_SIZE_HEADER / 3,
                 "İMALAT FOTOĞRAFLARI", font_bold, FONT_SIZE_HEADER,
                 alignment='center', bold=True)
        current_y = photos_band_y  # Bitişik, boşluk yok
        
        # ============================================================
        # FOTOĞRAF GRID (2x4) - Tüm sayfalar için
        # Header ile aynı genişlikte olmalı
        # ============================================================
        while photo_index < total_photos:
            available_width = band_width  # Header ile hizalı
            available_height = current_y - MARGIN_BOTTOM
            photo_cell_width = available_width / PHOTO_GRID_COLS
            photo_cell_height = (available_height - PHOTO_LABEL_HEIGHT * PHOTO_GRID_ROWS) / PHOTO_GRID_ROWS
            photo_image_height = photo_cell_height - PHOTO_LABEL_HEIGHT
            
            grid_start_y = current_y
            photos_on_this_page = 0
            
            for row in range(PHOTO_GRID_ROWS):
                for col in range(PHOTO_GRID_COLS):
                    if photo_index >= total_photos:
                        break
                    
                    cell_x = MARGIN_LEFT + col * photo_cell_width
                    cell_y = grid_start_y - (row + 1) * photo_cell_height
                    
                    # Hücre kenarlığı
                    draw_box(c, cell_x, cell_y, photo_cell_width, photo_cell_height)
                    
                    # Fotoğraf - çok az padding ekle (yukarı ve aşağıdan)
                    photo_path = photo_files[photo_index]
                    photo_padding = 0.05*cm  # Çok az padding
                    image_y = cell_y + PHOTO_LABEL_HEIGHT + photo_padding
                    image_height = photo_image_height - 2*photo_padding
                    if os.path.exists(photo_path):
                        draw_image_fit(c, cell_x + photo_padding, image_y, photo_cell_width - 2*photo_padding, image_height, photo_path)
                    
                    # Fotoğraf etiketi - üstünde çizgi ile kutunun içindeymiş gibi
                    label_y = cell_y
                    label_box_height = PHOTO_LABEL_HEIGHT
                    # Etiket kutusu çiz
                    draw_box(c, cell_x, label_y, photo_cell_width, label_box_height)
                    # Etiket metni
                    label_text = f"FOTO-{photo_index + 1}"
                    draw_text(c, cell_x + photo_cell_width / 2, label_y + label_box_height / 2 - FONT_SIZE_SMALL / 3,
                             label_text, font_regular, FONT_SIZE_SMALL, alignment='center')
                    
                    photo_index += 1
                    photos_on_this_page += 1
                
                if photo_index >= total_photos:
                    break
            
            # Sayfayı bitir
            c.showPage()
            page_num += 1
            
            # Eğer daha fazla fotoğraf varsa yeni sayfa başlat
            if photo_index < total_photos:
                current_y = PAGE_HEIGHT - MARGIN_TOP
            else:
                break
        
        # PDF'i kaydet
        c.save()
        
        if os.path.exists(pdf_filepath) and os.path.getsize(pdf_filepath) > 0:
            print(f"PDF başarıyla oluşturuldu: {pdf_filepath}")
            return True
        return False
        
    except Exception as e:
        print(f"PDF oluşturma hatası: {type(e).__name__}: {e}")
        import traceback
        traceback.print_exc()
        return False

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
    
    try:
        # Fotoğrafları geçici dizine kaydet
        photo_files = []
        for i, photo in enumerate(photos[:8]):
            if photo and photo.filename:
                # Dosya uzantısını al veya varsayılan .jpg kullan
                ext = os.path.splitext(photo.filename)[1].lower()
                if ext not in ['.jpg', '.jpeg', '.png', '.webp']:
                    ext = '.jpg'
                
                temp_photo_path = os.path.join(temp_dir, f"photo_{i+1}{ext}")
                
                # Dosyayı kaydet ve hemen kapat (memory için)
                photo.save(temp_photo_path)
                photo.close() # Flask FileStorage nesnesini kapat
                
                photo_files.append(temp_photo_path)

        # PDF oluştur
        pdf_filename = f"rapor-{safe_date}-{uuid.uuid4().hex[:8]}.pdf"
        pdf_filepath = os.path.join(OUTPUT_DIR, pdf_filename)
        
        print(f"PDF oluşturma başlıyor: {pdf_filepath}")
        
        pdf_created = generate_pdf(data, photo_files, pdf_filepath)
        
        if not pdf_created:
            raise Exception("PDF oluşturulamadı.")

        return pdf_filepath

    finally:
        # Geçici dizini ve içindeki her şeyi temizle
        try:
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
        except Exception as e:
            print(f"Geçici dizin temizleme hatası: {e}")

