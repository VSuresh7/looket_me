import threading
import time
import webview
from app import app  # Your main Flask app
from waitress import serve

def start_server():
    """Runs the Waitress production server in a background thread."""
    # threads=1 keeps your Excel database safe from multi-request crashes
    serve(app, host="127.0.0.1", port=8585, threads=1)

if __name__ == '__main__':
    # 1. Start the Flask server in the background
    server_thread = threading.Thread(target=start_server)
    server_thread.daemon = True
    server_thread.start()
    
    # 2. Give the background server a brief second to initialize
    time.sleep(1)
    
    # 3. Create a dedicated desktop app window pointing to your local server
    # This removes the browser tabs and gives you a premium, clean look
    webview.create_window(
        title="LOOKET ME - Men's & Kid's Wear", 
        url="http://127.0.0.1:8585",
        width=1280,
        height=800,
        resizable=True
    )
    
    # 4. Start the GUI loop
    webview.start()