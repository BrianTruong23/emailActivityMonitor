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
        result = subprocess.run(
            ["python3", "main.py"],
            check=True,
            capture_output=True,
            text=True
        )
        return jsonify({"status": "success", "stdout": result.stdout})
    except subprocess.CalledProcessError as e:
        # Return stdout and stderr so you can debug
        return jsonify({
            "status": "error",
            "message": str(e),
            "stdout": e.stdout,
            "stderr": e.stderr
        }), 500

if __name__ == "__main__":
    app.run(debug=True)
