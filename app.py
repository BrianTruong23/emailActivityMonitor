from flask import Flask, render_template, jsonify
import subprocess
from flask_cors import CORS
from flask import send_file
from main import load_render_credentials, get_all_unread_emails, EXCEL_PATH, build


app = Flask(__name__)
CORS(app)


from flask import send_file

@app.route("/get-excel")
def get_excel():
    return send_file("result/log2.xlsx", as_attachment=False)


@app.route("/get-log")
def get_log():
    return send_file("result/log2.xlsx", as_attachment=False)

@app.route("/download-excel")
def download_excel():
    return send_file("result/log2.xlsx", as_attachment=True)

@app.route("/")
def index():
    return render_template("index.html")

from flask import jsonify
import subprocess

@app.route("/run-script", methods=["POST"])
def run_script():
    try:
        # Run main.py
        result = subprocess.run(
            ["python3", "main.py"],
            check=True,
            capture_output=True,
            text=True
        )
        return jsonify({
            "status": "success",
            "stdout": result.stdout
        })
    except subprocess.CalledProcessError as e:
        return jsonify({
            "status": "error",
            "message": str(e),
            "stdout": e.stdout,
            "stderr": e.stderr
        }), 500


@app.route("/get-emails", methods=["GET"])
def get_emails():
    print("✅ /get-emails route was called!")

    creds = load_render_credentials()
    print("✅ Credentials loaded.")

    gmail_service = build('gmail', 'v1', credentials=creds)
    print("✅ Gmail service built.")

    emails = get_all_unread_emails(gmail_service)
    print(f"✅ Fetched {len(emails)} emails.")

    simplified = []
    for email in emails:
        simplified.append({
            "from": email.get("From", "(No Sender)"),
            "subject": email.get("Subject", "(No Subject)"),
            "date": email.get("Date", "")
        })

    print("✅ Prepared JSON response.")
    return jsonify(simplified)




if __name__ == "__main__":
    app.run(debug=True)
