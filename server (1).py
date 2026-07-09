# -*- coding: utf-8 -*-
"""
TechOps Mail Tracker — Backend Server
======================================
Run this on the dedicated machine ONCE. It stays running in the background
and listens for requests from the team's webpage.

Requirements:
    pip install flask

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
CORS(app)  # allows the HTML page to call this server from a browser

# Path to the tracker script — must be in the same folder as server.py
SCRIPT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                           "track_outlook_mails_com.py")

# The running script process (None if not started)
tracker_process = None


def is_running():
    """Returns True if the tracker script is currently running."""
    global tracker_process
    if tracker_process is None:
        return False
    # poll() returns None if process is still running, otherwise the exit code
    return tracker_process.poll() is None


@app.route("/status", methods=["GET"])
def status():
    """Returns whether the tracker is currently running."""
    return jsonify({
        "running": is_running(),
        "message": "Tracker is running" if is_running() else "Tracker is stopped"
    })


@app.route("/run", methods=["POST"])
def run():
    """Starts the tracker script if it is not already running."""
    global tracker_process
    if is_running():
        return jsonify({
            "success": False,
            "message": "Tracker is already running"
        })
    try:
        tracker_process = subprocess.Popen(
            [sys.executable, SCRIPT_PATH],
            creationflags=subprocess.CREATE_NEW_PROCESS_GROUP,
        )
        return jsonify({
            "success": True,
            "message": "Tracker started successfully"
        })
    except Exception as e:
        return jsonify({
            "success": False,
            "message": f"Failed to start tracker: {e}"
        }), 500


@app.route("/stop", methods=["POST"])
def stop():
    """Stops the tracker script gracefully via SIGINT (same as Ctrl+C)
    so the script's KeyboardInterrupt handler runs and updates the dashboard."""
    global tracker_process
    if not is_running():
        return jsonify({
            "success": False,
            "message": "Tracker is not running"
        })
    try:
        import signal
        # Send SIGINT (Ctrl+C equivalent) so the script's KeyboardInterrupt
        # handler runs — this triggers the final dashboard update before exit
        if sys.platform == "win32":
            tracker_process.send_signal(signal.CTRL_C_EVENT)
        else:
            tracker_process.send_signal(signal.SIGINT)
        tracker_process.wait(timeout=30)  # give it time to finish dashboard update
        tracker_process = None
        return jsonify({
            "success": True,
            "message": "Tracker stopped — dashboard updated"
        })
    except Exception as e:
        # Fall back to terminate if signal fails
        try:
            tracker_process.terminate()
            tracker_process = None
        except Exception:
            pass
        return jsonify({
            "success": False,
            "message": f"Stopped forcefully (dashboard may not have updated): {e}"
        })


if __name__ == "__main__":
    print("TechOps Tracker Server starting...")
    print(f"Script path: {SCRIPT_PATH}")
    print("Listening on http://EXLAPLPX4dnzrAk:5000")
    print("Keep this window open. Press Ctrl+C to shut down the server.\n")
    app.run(host="0.0.0.0", port=5000, debug=False)
