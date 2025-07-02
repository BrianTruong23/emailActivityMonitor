from flask import Flask, render_template, jsonify, request, send_file
import subprocess
from flask_cors import CORS
from flask_sqlalchemy import SQLAlchemy
from main import load_credentials, get_all_unread_emails, EXCEL_PATH, build
import pandas as pd

app = Flask(__name__)
CORS(app)

# Database config
app.config["SQLALCHEMY_DATABASE_URI"] = "sqlite:////var/data/status.db"
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
db = SQLAlchemy(app)

# Database model
class Message(db.Model):
    __tablename__ = "messages"
    message_id = db.Column(db.String, primary_key=True)
    email_sender = db.Column(db.String)
    date = db.Column(db.String)
    time_received = db.Column(db.String)
    wait_time = db.Column(db.String)
    status = db.Column(db.String)

# Routes
@app.route("/")
def index():
    return render_template("index.html")

@app.route("/get-excel")
def get_excel():
    return send_file("result/log2.xlsx", as_attachment=False)

@app.route("/get-log")
def get_log():
    return send_file("result/log2.xlsx", as_attachment=False)

@app.route("/download-excel")
def download_excel():
    return send_file("result/log2.xlsx", as_attachment=True)

@app.route("/run-script", methods=["POST"])
def run_script():
    try:
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

@app.route("/update_status", methods=["POST"])
def update_status():
    data = request.get_json()
    message_id = data["message_id"]
    new_status = data["status"]

    msg = Message.query.get(message_id)
    if msg:
        msg.status = new_status
        db.session.commit()
        return jsonify({"success": True})
    else:
        return jsonify({"success": False, "error": "Message not found"})

@app.route("/get-status")
def get_status():
    messages = Message.query.all()
    data = []
    for msg in messages:
        data.append({
            "Message_ID": msg.message_id,
            "Email Sender": msg.email_sender,
            "Date": msg.date,
            "Time Received": msg.time_received,
            "Wait Time": msg.wait_time,
            "Status": msg.status
        })
    return jsonify(data)

@app.route("/get-emails", methods=["GET"])
def get_emails():
    print("✅ /get-emails route was called!")
    creds = load_credentials()
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
