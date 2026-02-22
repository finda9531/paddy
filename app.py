import os
import subprocess
import asyncio
from flask import Flask, render_template, request, jsonify
import pyautogui
from winrt.windows.media.control import GlobalSystemMediaTransportControlsSessionManager
import requests

from ctypes import cast, POINTER
import comtypes
from comtypes import CLSCTX_ALL, CoInitialize, CoUninitialize, CoCreateInstance
from pycaw.pycaw import IMMDeviceEnumerator, IAudioEndpointVolume

app = Flask(__name__)

# --- Windowsのミュート状態を取得 ---
def get_mute_state():
    CoInitialize()
    try:
        CLSID_MMDeviceEnumerator = comtypes.GUID("{BCDE0395-E52F-467C-8E3D-C4579291692E}")
        enumerator = CoCreateInstance(
            CLSID_MMDeviceEnumerator,
            interface=IMMDeviceEnumerator,
            clsctx=CLSCTX_ALL
        )
        endpoint = enumerator.GetDefaultAudioEndpoint(0, 1)
        interface = endpoint.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = cast(interface, POINTER(IAudioEndpointVolume))
        return volume.GetMute() == 1
    except Exception as e:
        print(f"Audio Error: {e}")
        return False
    finally:
        CoUninitialize()

async def get_media_info():
    try:
        manager = await GlobalSystemMediaTransportControlsSessionManager.request_async()
        session = manager.get_current_session()
        if session:
            info = await session.try_get_media_properties_async()
            return {"title": info.title, "artist": info.artist}
        return {"title": "Not Playing", "artist": ""}
    except Exception:
        return {"title": "Not Playing", "artist": ""}

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/api/now_playing", methods=["GET"])
def now_playing():
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    try:
        info = loop.run_until_complete(get_media_info())
    finally:
        loop.close()
        
    info["is_muted"] = get_mute_state()
    return jsonify(info), 200

@app.route("/api/weather", methods=["GET"])
def get_weather():
    try:
        url = "https://api.open-meteo.com/v1/forecast?latitude=36.23&longitude=137.97&current_weather=true&timezone=Asia%2FTokyo"
        response = requests.get(url)
        data = response.json()
        current = data.get("current_weather", {})
        temp = current.get("temperature")
        code = current.get("weathercode")
        is_day = current.get("is_day", 1) # 昼夜の判定データを取得
        
        weather_map = {0: "Clear", 1: "Mainly Clear", 2: "Partly Cloudy", 3: "Overcast", 45: "Fog", 48: "Fog", 51: "Drizzle", 61: "Rain", 71: "Snow", 95: "Thunderstorm"}
        status = weather_map.get(code, "Cloudy")
        return jsonify({"temp": temp, "status": status, "is_day": is_day}), 200
    except Exception:
        return jsonify({"temp": "--", "status": "Unknown", "is_day": 1}), 500

@app.route("/api/command", methods=["POST"])
def execute_command():
    data = request.json
    action_type = data.get("type")
    target = data.get("target")

    try:
        if action_type == "shortcut":
            keys = target.split(",")
            pyautogui.hotkey(*keys)
        elif action_type == "launch":
            os.startfile(target)
        elif action_type == "media":
            pyautogui.press(target)
        elif action_type == "launch_url":
            browser = target.get("browser")
            url = target.get("url")
            if browser == "firefox":
                subprocess.Popen(["start", "firefox", url], shell=True)
            elif browser == "chrome_app":
                subprocess.Popen(["start", "chrome", f"--app={url}"], shell=True)
            elif browser == "chrome":
                subprocess.Popen(["start", "chrome", url], shell=True)
        return jsonify({"status": "success"}), 200
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route("/api/bluetooth", methods=["POST"])
def connect_bluetooth():
    try:
        result1 = subprocess.run(['btcom', '-n', 'AirPods', '-c'], capture_output=True, text=True)
        if result1.returncode == 0:
            return jsonify({"status": "success", "message": "AirPods connected"}), 200
        result2 = subprocess.run(['btcom', '-n', 'AirPods Pro', '-c'], capture_output=True, text=True)
        if result2.returncode == 0:
            return jsonify({"status": "success", "message": "AirPods Pro connected"}), 200
        return jsonify({"status": "error", "message": "Connection failed"}), 500
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
