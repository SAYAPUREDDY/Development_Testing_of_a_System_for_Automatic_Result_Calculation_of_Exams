import subprocess
import threading

def start_backend():
    subprocess.run(["uvicorn", "main:app", "--reload"])

def start_frontend():
    subprocess.run(["streamlit", "run", "app.py"])

if __name__ == "__main__":
    # Run backend in a separate thread
    threading.Thread(target=start_backend, daemon=True).start()

    # Run frontend in the main thread
    start_frontend()
