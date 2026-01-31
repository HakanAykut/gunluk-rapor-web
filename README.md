# Günlük Faaliyet Raporu Web Uygulaması

Flask tabanlı PDF rapor oluşturma uygulaması.

## Deployment Seçenekleri

### 1. Render.com (Önerilen - Ücretsiz)

1. [Render.com](https://render.com) hesabı oluşturun
2. "New +" → "Web Service" seçin
3. GitHub repo'nuzu bağlayın veya direkt deploy edin
4. Ayarlar:
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn app:app`
   - **Environment**: Python 3
   - **Plan**: Free

### 2. Heroku

1. [Heroku CLI](https://devcenter.heroku.com/articles/heroku-cli) yükleyin
2. Terminal'de:
```bash
heroku login
heroku create gunluk-rapor-app
git init
git add .
git commit -m "Initial commit"
git push heroku main
```

### 3. Railway.app

1. [Railway.app](https://railway.app) hesabı oluşturun
2. "New Project" → "Deploy from GitHub repo"
3. Repo'nuzu seçin
4. Otomatik deploy başlar

## Gereksinimler

- Python 3.9+
- Flask 3.1.2
- ReportLab 4.4.6
- Pillow 11.3.0
- Gunicorn 23.0.0

## Lokal Çalıştırma

Lokalde test yapmak için şu adımları izleyin:

1.  **Sanal Ortam Oluşturun (Opsiyonel ama Önerilir):**
    ```bash
    python -m venv venv
    source venv/bin/activate  # macOS/Linux
    # veya
    venv\Scripts\activate     # Windows
    ```

2.  **Bağımlılıkları Yükleyin:**
    ```bash
    pip install -r requirements.txt
    ```

3.  **Uygulamayı Başlatın:**
    ```bash
    python app.py
    ```

Uygulama `http://127.0.0.1:5000` adresinde çalışacaktır. Değişiklikleri lokalde test ettikten sonra canlıya (git push) alabilirsiniz.

## Performans ve Bellek Optimizasyonları

Bu uygulama, 512MB RAM gibi kısıtlı kaynaklarda çalışacak şekilde optimize edilmiştir:
- **Client-Side Resizing:** Fotoğraflar tarayıcıda 1000px boyutuna düşürülüp JPEG formatında gönderilir.
- **Memory Management:** Sunucu tarafında Pillow nesneleri işlendikten sonra hemen kapatılır ve `gc.collect()` ile bellek temizlenir.
- **One-by-One Processing:** Fotoğraflar PDF'e eklenirken tek tek işlenerek bellek kullanımı minimize edilir.

## Önemli Notlar

- Font dosyaları (`DejaVuSans.ttf`, `DejaVuSans-Bold.ttf`) proje kök dizininde olmalı
- Logo dosyası (`Resim1.png`) proje kök dizininde olmalı
- `generated_pdfs/` klasörü otomatik oluşturulur

