from flask import Flask, jsonify
import subprocess
from flask_cors import CORS

app = Flask(__name__)
CORS(app)  # enable cross-origin requests from Netlify

@app.route('/run-script', methods=['POST'])
def run_script():
    try:
        subprocess.run(['python3', 'main.py'], check=True)
        return jsonify({"status": "success"})
    except subprocess.CalledProcessError as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/')
def home():
    return "✅ Backend is running."

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
