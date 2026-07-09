# -*- coding: utf-8 -*-
"""
TechOps Mail Tracker — Backend Server
======================================
Run this on the dedicated machine ONCE. It stays running in the background
and listens for requests from the team's webpage.

Requirements:
    pip install flask flask-cors

Run:
    python server.py

The server listens on port 5000. Leave this terminal window open.
The webpage at index.html connects to http://EXLAPLPX4dnzrAk:5000
"""

import subprocess
import sys
import os
from flask import Flask, jsonify
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

# Path to the tracker script — must be in the same folder as server.py
SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
SCRIPT_PATH = os.path.join(SCRIPT_DIR, "track_outlook_mails_com.py")

# Stop file — when this exists, the tracker shuts down cleanly
STOP_FILE = os.path.join(SCRIPT_DIR, ".tracker_stop")

tracker_process = None


def is_running():
    global tracker_process
    if tracker_process is None:
        return False
    return tracker_process.poll() is None


@app.route("/status", methods=["GET"])
def status():
    return jsonify({
        "running": is_running(),
        "message": "Tracker is running" if is_running() else "Tracker is stopped"
    })


@app.route("/run", methods=["POST"])
def run():
    global tracker_process
    if is_running():
        return jsonify({"success": False, "message": "Tracker is already running"})
    if os.path.exists(STOP_FILE):
        os.remove(STOP_FILE)
    try:
        tracker_process = subprocess.Popen([sys.executable, SCRIPT_PATH])
        return jsonify({"success": True, "message": "Tracker started successfully"})
    except Exception as e:
        return jsonify({"success": False, "message": f"Failed to start: {e}"}), 500


@app.route("/stop", methods=["POST"])
def stop():
    global tracker_process
    if not is_running():
        return jsonify({"success": False, "message": "Tracker is not running"})

    # Create stop file — script checks for this every poll cycle and exits cleanly
    with open(STOP_FILE, "w") as f:
        f.write("stop")

    # Wait up to 60s for the script to finish its clean shutdown
    try:
        tracker_process.wait(timeout=60)
        tracker_process = None
        return jsonify({"success": True, "message": "Tracker stopped — dashboard updated"})
    except subprocess.TimeoutExpired:
        tracker_process.kill()
        tracker_process = None
        if os.path.exists(STOP_FILE):
            os.remove(STOP_FILE)
        return jsonify({"success": False,
                        "message": "Stopped forcefully — dashboard may not have updated"})


if __name__ == "__main__":
    if os.path.exists(STOP_FILE):
        os.remove(STOP_FILE)
    print("TechOps Tracker Server starting...")
    print(f"Script path: {SCRIPT_PATH}")
    print("Listening on http://EXLAPLPX4dnzrAk:5000")
    print("Keep this window open. Press Ctrl+C to shut down the server.\n")
    app.run(host="0.0.0.0", port=5000, debug=False)
