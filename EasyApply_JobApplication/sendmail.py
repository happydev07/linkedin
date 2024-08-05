import os
import time
import subprocess
import winsound
import logging

logging.basicConfig(level=logging.INFO)

def play_sound():
    # Play Windows default sound
    winsound.MessageBeep()

def main():
    file_path = r"C:\Users\devra\Desktop\EasyApply_JobApplication\prompt.txt"
    last_updated = os.path.getmtime(file_path)

    while True:
        time.sleep(1)  # Check for updates every 1 second
        updated = os.path.getmtime(file_path)
        if updated > last_updated:
            last_updated = updated
            logging.info("File updated. Executing resume generation script.")
            script_path = r"C:\Users\devra\Desktop\EasyApply_JobApplication\gpt.py"
            try:
                subprocess.run(['python', script_path], check=True)
                logging.info("Script executed successfully.")
                print("Resume updated!")  # Display message after updating resume
                play_sound()  # Play sound after updating resume
            except subprocess.CalledProcessError as e:
                logging.error(f"Failed to execute script: {e}")

if __name__ == "__main__":
    main()
