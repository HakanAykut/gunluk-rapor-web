from flask import Flask, render_template, request, send_file
from report_generator import generate_report
import os

app = Flask(__name__)

@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")

@app.route("/generator-test", methods=["POST"])
def generator_test():
    text = request.form.get("report_text", "")
    photos = request.files.getlist("photos")

    filepath = generate_report(text, photos)

    return send_file(
        filepath,
        as_attachment=True,
        download_name=os.path.basename(filepath),
        mimetype="application/pdf"
    )

if __name__ == "__main__":
    app.run(debug=True)
