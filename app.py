from flask import Flask, render_template, request, send_file, jsonify
from report_generator import generate_report
import os

app = Flask(__name__)

@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")

@app.route("/generator-test", methods=["POST"])
def generator_test():
    try:
        text = request.form.get("report_text", "")
        photos = request.files.getlist("photos")

        if not text.strip():
            return jsonify({"error": "Rapor metni boş olamaz"}), 400

        filepath = generate_report(text, photos)

        return send_file(
            filepath,
            as_attachment=True,
            download_name=os.path.basename(filepath),
            mimetype="application/pdf"
        )
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
    app.run(host="0.0.0.0", port=port, debug=False)
