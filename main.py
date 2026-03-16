"""
HR Attendance Analytics System – Launcher
Opens the Flask app inside a native pywebview window.
Falls back to opening in browser if pywebview is not installed.
"""
import sys
import os
import threading
import webbrowser
import time

# Add the app directory to path
APP_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(APP_DIR)

def run_flask():
    """Start Flask in a background thread."""
    from app import app
    app.run(host='127.0.0.1', port=5000, debug=False, use_reloader=False)


def main():
    # Start Flask server
    server = threading.Thread(target=run_flask, daemon=True)
    server.start()
    time.sleep(1.5)  # Wait for server to start

    try:
        import webview
        webview.create_window(
            'HR Attendance Analytics System',
            'http://127.0.0.1:5000',
            width=1400,
            height=900,
            min_size=(1000, 700),
            resizable=True,
            text_select=True,
        )
        webview.start()
    except ImportError:
        print("pywebview not installed. Opening in your default browser...")
        print("Install it with: pip install pywebview")
        webbrowser.open('http://127.0.0.1:5000')
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            print("\nShutting down...")


if __name__ == '__main__':
    main()
