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
    return render_template("view_pdf.html", filename=filename)

@app.route("/download-pdf/<filename>", methods=["GET"])
def download_pdf(filename):
    """PDF indirme endpoint'i"""
    try:
        from pdf_generator import OUTPUT_DIR
        filepath = os.path.join(OUTPUT_DIR, filename)
        if not os.path.exists(filepath):
            return jsonify({"error": "PDF bulunamadı"}), 404
        return send_file(
            filepath,
            as_attachment=True,  # İndirme
            download_name=filename,
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
        tarih = request.form.get("tarih", "")
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

        # Tarihi DD.MM.YYYY formatına çevir
        from datetime import datetime
        try:
            tarih_obj = datetime.strptime(tarih, "%Y-%m-%d")
            tarih_formatted = tarih_obj.strftime("%d.%m.%Y")
        except:
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

        # Data dict oluştur
        data = {
            "tarih": tarih_formatted,
            "rapor_no": rapor_no,
            "yapilan_isler": yapilan_isler
        }

        filepath = generate_report(data, photos)
        
        # PDF dosya adını al (URL için)
        filename = os.path.basename(filepath)
        
        # PDF görüntüleme sayfasına yönlendir
        from flask import redirect, url_for
        return redirect(url_for('view_pdf', filename=filename))
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
