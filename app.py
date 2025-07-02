from flask import Flask, jsonify
import subprocess

app = Flask(__name__)

@app.route('/run-script', methods=['POST'])
def run_script():
    try:
        # Run your main.py script
        subprocess.run(['python3', 'main.py'], check=True)
        return jsonify({"status": "success"})
    except subprocess.CalledProcessError as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/')
def hello():
    return "✅ Flask backend is running."

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
