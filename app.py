from flask import Flask, render_template, jsonify
import subprocess
from flask_cors import CORS
from flask import send_file

app = Flask(__name__)
CORS(app)


@app.route("/get-log")
def get_log():
    return send_file("result/log2.xlsx", as_attachment=False)


@app.route("/")
def index():
    return render_template("index.html")

@app.route("/run-script", methods=["POST"])
def run_script():
    try:
        # Run your main.py script
        subprocess.run(["python3", "main.py"], check=True)
        return jsonify({"status": "success"})
    except subprocess.CalledProcessError as e:
        return jsonify({"status": "error", "message": str(e)}), 500

if __name__ == "__main__":
    app.run(debug=True)
