from flask import Flask, render_template, request, send_file, jsonify
from pdf_generator import generate_report
import os

app = Flask(__name__)

@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")

@app.route("/view-pdf/<filename>", methods=["GET"])
def view_pdf(filename):
    """PDF görüntüleme sayfası"""
    tarih = request.args.get("tarih", "")
    return render_template("view_pdf.html", filename=filename, tarih=tarih)

@app.route("/download-pdf/<filename>", methods=["GET"])
def download_pdf(filename):
    """PDF indirme endpoint'i"""
    try:
        from pdf_generator import OUTPUT_DIR
        filepath = os.path.join(OUTPUT_DIR, filename)
        if not os.path.exists(filepath):
            return jsonify({"error": "PDF bulunamadı"}), 404
        
        # İndirme dosya adını oluştur: Günlük_Rapor_{Tarih}.pdf
        tarih = request.args.get("tarih", "")
        if tarih:
            # Tarih formatını dosya adı için uygun hale getir (noktaları alt çizgi ile değiştir)
            tarih_for_filename = tarih.replace(".", "_")
            download_filename = f"Günlük_Rapor_{tarih_for_filename}.pdf"
        else:
            # Tarih yoksa orijinal dosya adını kullan
            download_filename = filename
        
        return send_file(
            filepath,
            as_attachment=True,  # İndirme
            download_name=download_filename,
            mimetype="application/pdf"
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/pdf/<filename>", methods=["GET"])
def serve_pdf(filename):
    """PDF dosyasını tarayıcıda göster"""
    try:
        from pdf_generator import OUTPUT_DIR
        filepath = os.path.join(OUTPUT_DIR, filename)
        if not os.path.exists(filepath):
            return jsonify({"error": "PDF bulunamadı"}), 404
        return send_file(
            filepath,
            as_attachment=False,  # Tarayıcıda görüntüle
            mimetype="application/pdf"
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/generator-test", methods=["POST"])
def generator_test():
    try:
        # Form alanlarından direkt al
        proje = request.form.get("proje", "Fetihtepe")  # Varsayılan: Fetihtepe
        tarih = request.form.get("tarih", "")
        tarih_tipi = request.form.get("tarih_tipi", "gunluk")
        rapor_no = request.form.get("rapor_no", "")
        yapilan_isler_text = request.form.get("yapilan_isler", "")
        photos = request.files.getlist("photos")

        # Validasyon
        if not tarih:
            return jsonify({"error": "Tarih boş olamaz"}), 400
        if not rapor_no:
            return jsonify({"error": "Rapor No boş olamaz"}), 400
        if not yapilan_isler_text.strip():
            return jsonify({"error": "Yapılan İşler boş olamaz"}), 400

        # Tarihi işle
        from datetime import datetime, timedelta
        try:
            tarih_obj = datetime.strptime(tarih, "%Y-%m-%d")
            
            if tarih_tipi == "3gunluk":
                # 3 günlük format (Hafta sonunu atla): 12/13/14.01.2026
                def get_next_workday(d):
                    next_d = d + timedelta(days=1)
                    while next_d.weekday() >= 5: # 5=Cumartesi, 6=Pazar
                        next_d += timedelta(days=1)
                    return next_d
                
                t2 = get_next_workday(tarih_obj)
                t3 = get_next_workday(t2)
                tarih_formatted = f"{tarih_obj.strftime('%d')}/{t2.strftime('%d')}/{t3.strftime('%d.%m.%Y')}"
            elif tarih_tipi == "3gunluk_ozel":
                # Özel 3 günlük format: Kullanıcının seçtiği 3 tarih
                tarih2 = request.form.get("tarih2", "")
                tarih3 = request.form.get("tarih3", "")
                if not tarih2 or not tarih3:
                    return jsonify({"error": "Özel 3 günlük rapor için tüm tarihler seçilmelidir"}), 400
                
                t2_obj = datetime.strptime(tarih2, "%Y-%m-%d")
                t3_obj = datetime.strptime(tarih3, "%Y-%m-%d")
                tarih_formatted = f"{tarih_obj.strftime('%d')}/{t2_obj.strftime('%d')}/{t3_obj.strftime('%d.%m.%Y')}"
            else:
                # Günlük format: 12.01.2026
                tarih_formatted = tarih_obj.strftime("%d.%m.%Y")
        except Exception as e:
            print(f"Tarih işleme hatası: {e}")
            tarih_formatted = tarih

        # Yapılan işleri satırlara böl
        yapilan_isler = []
        for line in yapilan_isler_text.strip().split("\n"):
            line = line.strip()
            if line:
                # Başında • varsa kaldır, yoksa ekle
                if line.startswith("•"):
                    yapilan_isler.append(line[1:].strip())
                else:
                    yapilan_isler.append(line)

        # Proje seçimine göre başlığı belirle
        if proje == "Arap Camii":
            proje_basligi = "Arap Camii Kuran Kursu Güçlendirme Projesi"
        else:  # Fetihtepe (varsayılan)
            proje_basligi = "FETİHTEPE MERKEZ CAMİ'İ GÜÇLENDİRME VE YENİLEME PROJESİ"

        # Data dict oluştur
        data = {
            "tarih": tarih_formatted,
            "rapor_no": rapor_no,
            "yapilan_isler": yapilan_isler,
            "proje_basligi": proje_basligi
        }

        filepath = generate_report(data, photos)
        
        # PDF dosya adını al (URL için)
        filename = os.path.basename(filepath)
        
        # PDF görüntüleme sayfasına yönlendir (tarih bilgisini de gönder)
        from flask import redirect, url_for
        return redirect(url_for('view_pdf', filename=filename, tarih=tarih_formatted))
    except FileNotFoundError as e:
        return jsonify({"error": str(e)}), 404
    except Exception as e:
        import traceback
        error_msg = str(e)
        traceback_str = traceback.format_exc()
        print(f"Error: {error_msg}\n{traceback_str}")
        return jsonify({"error": f"PDF oluşturulurken hata oluştu: {error_msg}"}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    # Production için 0.0.0.0, development için 127.0.0.1
    host = "0.0.0.0" if os.environ.get("PORT") else "127.0.0.1"
    debug = os.environ.get("FLASK_ENV") != "production"
    app.run(host=host, port=port, debug=debug)
