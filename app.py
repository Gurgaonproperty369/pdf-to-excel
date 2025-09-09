# app.py
import os
from flask import Flask, request, redirect, url_for, render_template, send_from_directory, flash
from werkzeug.utils import secure_filename
from converter import pdf_to_excel

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"
ALLOWED_EXTENSIONS = {"pdf"}
MAX_CONTENT_LENGTH = 50 * 1024 * 1024  # 50 MB limit

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app = Flask(__name__)
app.secret_key = "replace-this-with-a-secure-random-secret"
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["OUTPUT_FOLDER"] = OUTPUT_FOLDER
app.config["MAX_CONTENT_LENGTH"] = MAX_CONTENT_LENGTH


def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/convert", methods=["POST"])
def convert():
    # file
    if "pdf_file" not in request.files:
        flash("No file part")
        return redirect(request.url)
    file = request.files["pdf_file"]
    if file.filename == "":
        flash("No selected file")
        return redirect(request.url)
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        pdf_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
        file.save(pdf_path)

        # parameters
        method = request.form.get("method", "auto")  # auto, camelot, tabula, pdfplumber
        pages = request.form.get("pages", "all")
        camelot_flavor = request.form.get("camelot_flavor", "lattice")

        out_basename = os.path.splitext(filename)[0] + ".xlsx"
        excel_out_path = os.path.join(app.config["OUTPUT_FOLDER"], out_basename)

        try:
            outpath, used_method = pdf_to_excel(
                pdf_path,
                excel_out_path,
                method=method,
                pages=pages,
                camelot_flavor=camelot_flavor,
            )
            download_url = url_for("download_file", filename=os.path.basename(outpath))
            return render_template("result.html", download_url=download_url, method=used_method)
        except Exception as e:
            # report error
            flash(f"Conversion failed: {e}")
            return redirect(url_for("index"))
    else:
        flash("Invalid file type. Only PDF allowed.")
        return redirect(url_for("index"))


@app.route("/downloads/<path:filename>")
def download_file(filename):
    return send_from_directory(app.config["OUTPUT_FOLDER"], filename, as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)

