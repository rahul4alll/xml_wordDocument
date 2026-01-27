from flask import Flask, render_template, request, jsonify, send_file
import os
import tempfile
from threading import Timer
from decipher_api import lookup_survey, fetch_survey_xml
from pqr_exporter import export_word_from_xml_file
from config import Config
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, "input")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

os.makedirs(INPUT_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)
app = Flask(__name__)
app.secret_key = Config.FLASK_SECRET_KEY

@app.route("/")
def home():
    return render_template("index.html")

@app.route("/api/lookup", methods=["POST"])
def api_lookup():
    survey_id = request.json.get("survey_id")
    print(f"Received survey_id: {survey_id}")  # Log the received survey_id
    response = lookup_survey(survey_id=survey_id)
    print(f"Lookup response: {response}")  # Log the response from lookup_survey
    return jsonify(response)
# Cleanup function to remove temporary files
def delete_file(path):
    try:
        if os.path.exists(path):
            os.remove(path)
    except Exception as e:
        print(f"Cleanup failed for {path}: {e}")

@app.route("/api/export", methods=["POST"])
def api_export():
    survey_id = request.json.get("survey_id")

    # ---------- 1️⃣ Download XML ----------
    xml_content = fetch_survey_xml(survey_id)

    xml_filename = f"{survey_id}.xml"
    xml_path = os.path.join(INPUT_DIR, xml_filename)

    with open(xml_path, "w", encoding="utf-8") as f:
        f.write(xml_content)

    # ---------- 2️⃣ Generate Word ----------
    word_filename = f"survey_{survey_id}.docx"
    word_path = os.path.join(OUTPUT_DIR, word_filename)

    export_word_from_xml_file(xml_path, word_path)
    # ---------- Schedule cleanup ----------
    Timer(30, delete_file, args=[xml_path]).start()
    Timer(30, delete_file, args=[word_path]).start()
    # ---------- 3️⃣ Download Word ----------
    return send_file(
        word_path,
        as_attachment=True,
        download_name=word_filename
    )


    return response

if __name__ == "__main__":
    app.run(debug=True)
