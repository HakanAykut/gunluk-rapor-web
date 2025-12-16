from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm, mm
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.utils import ImageReader
from PIL import Image as PILImage
import os

# A4 boyutları
PAGE_WIDTH, PAGE_HEIGHT = A4

# ============================================================
# LAYOUT SABİTLERİ
# ============================================================

# Header yüksekliği (yaklaşık 1/3 küçültüldü)
HEADER_HEIGHT = 1.2 * cm

# Margin'ler
MARGIN_LEFT = 0.5 * cm
MARGIN_RIGHT = 0.5 * cm
MARGIN_TOP = 0.5 * cm
MARGIN_BOTTOM = 0.5 * cm

# Header kolon genişlikleri (toplam sayfa genişliğinin yüzdeleri)
# Logo kolonunu biraz büyüt, orta kolonu küçült
HEADER_COL1_WIDTH = PAGE_WIDTH * 0.25  # Sol: Logo (biraz büyük)
HEADER_COL2_WIDTH = PAGE_WIDTH * 0.50  # Orta: Başlık (küçültüldü)
HEADER_COL3_WIDTH = PAGE_WIDTH * 0.25  # Sağ: Rapor bilgileri

# Header iç tablo (sağ kolon için)
HEADER_TABLE_CELL_HEIGHT = HEADER_HEIGHT / 2

# Gri bant yüksekliği (daraltıldı)
BAND_HEIGHT = 0.6 * cm

# Yapılan işler bölümü
# Gri bantlardan alınan alanı yapılan işler bölümüne veriyoruz
WORKS_TITLE_HEIGHT = 0.6 * cm
WORKS_ROW_HEIGHT = 0.45 * cm
WORKS_MAX_ROWS = 15  # Maksimum satır sayısı (gri bantlar daraltıldığı için artırıldı)

# Fotoğraf grid
PHOTO_GRID_COLS = 2
PHOTO_GRID_ROWS = 4
PHOTOS_PER_PAGE = PHOTO_GRID_COLS * PHOTO_GRID_ROWS
PHOTO_LABEL_HEIGHT = 0.4 * cm

# Font boyutları (tüm fontlar küçültüldü)
FONT_SIZE_TITLE = 7  # Orta başlık
FONT_SIZE_HEADER = 8  # Gri bant başlıkları
FONT_SIZE_NORMAL = 7  # Normal metin
FONT_SIZE_SMALL = 6  # Küçük metin

# ============================================================
# FONT YÜKLEME
# ============================================================

def setup_fonts(base_dir):
    """DejaVuSans fontlarını yükle"""
    dejavu_regular = os.path.join(base_dir, "DejaVuSans.ttf")
    dejavu_bold = os.path.join(base_dir, "DejaVuSans-Bold.ttf")
    
    if os.path.exists(dejavu_regular):
        pdfmetrics.registerFont(TTFont('DejaVuSans', dejavu_regular))
    else:
        raise FileNotFoundError("DejaVuSans.ttf bulunamadı!")
    
    if os.path.exists(dejavu_bold):
        pdfmetrics.registerFont(TTFont('DejaVuSans-Bold', dejavu_bold))
    else:
        raise FileNotFoundError("DejaVuSans-Bold.ttf bulunamadı!")
    
    return 'DejaVuSans', 'DejaVuSans-Bold'

# ============================================================
# HELPER FONKSİYONLAR
# ============================================================

def draw_box(canvas, x, y, width, height, border_color=colors.black, border_width=0.5, fill_color=None):
    """Kenarlıklı kutu çiz"""
    if fill_color:
        canvas.setFillColor(fill_color)
        canvas.rect(x, y, width, height, fill=1, stroke=0)
    
    canvas.setStrokeColor(border_color)
    canvas.setLineWidth(border_width)
    canvas.rect(x, y, width, height, fill=0, stroke=1)

def draw_text(canvas, x, y, text, font_name, font_size, color=colors.black, 
              alignment='left', width=None, bold=False):
    """Metin çiz (hizalama desteği ile)"""
    if bold and 'Bold' not in font_name:
        font_name = font_name.replace('DejaVuSans', 'DejaVuSans-Bold')
    
    canvas.setFont(font_name, font_size)
    canvas.setFillColor(color)
    
    if alignment == 'center':
        if width:
            canvas.drawCentredString(x + width/2, y, text)
        else:
            canvas.drawCentredString(x, y, text)
    elif alignment == 'right':
        if width:
            canvas.drawRightString(x + width, y, text)
        else:
            canvas.drawRightString(x, y, text)
    else:  # left
        canvas.drawString(x, y, text)

def draw_text_multiline(canvas, x, y, text, font_name, font_size, color=colors.black, 
                        line_height=None, max_width=None, alignment='left'):
    """Çok satırlı metin çiz"""
    if line_height is None:
        line_height = font_size * 1.2
    
    lines = text.split('\n')
    current_y = y
    
    for line in lines:
        if max_width:
            # Metni max_width'e sığdır (basit word wrap)
            words = line.split(' ')
            current_line = ''
            for word in words:
                test_line = current_line + (' ' if current_line else '') + word
                if canvas.stringWidth(test_line, font_name, font_size) <= max_width:
                    current_line = test_line
                else:
                    if current_line:
                        draw_text(canvas, x, current_y, current_line, font_name, font_size, color, alignment)
                        current_y -= line_height
                    current_line = word
            if current_line:
                draw_text(canvas, x, current_y, current_line, font_name, font_size, color, alignment)
                current_y -= line_height
        else:
            draw_text(canvas, x, current_y, line, font_name, font_size, color, alignment)
            current_y -= line_height
    
    return current_y

def draw_image_fit(canvas, x, y, width, height, image_path):
    """Görseli oranı koruyarak sığdır (contain)"""
    try:
        pil_img = PILImage.open(image_path)
        img_width, img_height = pil_img.size
        
        # Oranları hesapla
        scale_w = width / img_width
        scale_h = height / img_height
        scale = min(scale_w, scale_h)
        
        new_width = img_width * scale
        new_height = img_height * scale
        
        # Ortala
        offset_x = x + (width - new_width) / 2
        offset_y = y + (height - new_height) / 2
        
        # ReportLab ImageReader kullan
        img_reader = ImageReader(image_path)
        canvas.drawImage(img_reader, offset_x, offset_y, width=new_width, height=new_height)
        
        return True
    except Exception as e:
        print(f"Görsel yüklenemedi {image_path}: {e}")
        return False

def calculate_text_height(text, font_name, font_size, max_width):
    """Metnin yüksekliğini hesapla (çok satırlı)"""
    # Basit hesaplama - gerçekte canvas.stringWidth kullanılabilir
    lines = text.split('\n')
    line_height = font_size * 1.2
    return len(lines) * line_height

