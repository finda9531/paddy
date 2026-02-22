import os
from flask import Flask, render_template, request, jsonify
import pyautogui

app = Flask(__name__)

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/api/command", methods=["POST"])
def execute_command():
    data = request.json
    action_type = data.get("type")
    target = data.get("target")

    try:
        if action_type == "shortcut":
            # 複数キーの同時押しを実行
            keys = target.split(",")
            pyautogui.hotkey(*keys)
            return jsonify({"status": "success", "message": f"Shortcut {target} executed"}), 200
        
        elif action_type == "launch":
            # アプリの起動 “”Windows環境を想定“”
            os.startfile(target)
            return jsonify({"status": "success", "message": f"App {target} launched"}), 200
            
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

if __name__ == "__main__":
    # iPadなどローカルネットワーク内の他端末からアクセスできるように 0.0.0.0 を指定
    app.run(host="0.0.0.0", port=5000)