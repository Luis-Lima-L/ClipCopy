import shutil
import os
import subprocess
from openpyxl import Workbook
FFPROBE_PATH = "/usr/local/bin/ffprobe"
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox

def get_video_metadata(file_path):
    try:
        command = [
            FFPROBE_PATH, "-v", "error", "-show_entries", "stream=width,height,r_frame_rate",
            "-select_streams", "v:0", "-of", "default=noprint_wrappers=1:nokey=1", file_path
        ]
        
        result = subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)

        if result.stderr:
            raise Exception(f"FFprobe error: {result.stderr.strip()}")

        output = result.stdout.splitlines()

        if len(output) >= 2:
            width = int(output[0].strip()) if output[0].strip() else "Error"
            height = int(output[1].strip()) if output[1].strip() else "Error"

            if width != "Error" and height != "Error" and height != 0:
                aspect_ratio = round(width / height, 2)
            else:
                aspect_ratio = "Error"
        else:
            width = height = aspect_ratio = "Error"
        
        frame_rate_str = output[2].strip() if len(output) >= 3 and output[2] else "Error"
        try:
            if '/' in frame_rate_str:
                num, denom = frame_rate_str.split('/')
                frame_rate = round(float(num) / float(denom), 2)
            else:
                frame_rate = float(frame_rate_str)
        except Exception:
            frame_rate = "Error"


        command_duration = [
            FFPROBE_PATH, "-v", "error", "-show_entries", "format=duration",
            "-of", "default=noprint_wrappers=1:nokey=1", file_path
        ]
        
        result_duration = subprocess.run(command_duration, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        if result_duration.stderr:
            raise Exception(f"FFprobe error (duration): {result_duration.stderr.strip()}")
        
        duration = result_duration.stdout.strip()
        try:
            duration = float(duration) if duration else "Error"
        except ValueError:
            duration = "Error"

        return duration, frame_rate, width, height, aspect_ratio

    
    except Exception as e:
        print(f"Error processing file {file_path}: {e}")
        return str(e), str(e), "Error", "Error", "Error"

def get_folder_path(prompt="Select folder"):
    root = tk.Tk()
    root.withdraw() 
    folder_path = filedialog.askdirectory(title=prompt)
    if not folder_path:
        print("No folder selected. Exiting.")
        exit()
    root.quit()  
    root.destroy()  
    return folder_path


def ask_preserve_subfolders():
    root = tk.Tk()
    root.withdraw()
    result = messagebox.askyesno("Preserve Subfolders", "Do you want to keep the subfolder structure in the destination folder?")
    root.quit()  
    root.destroy()
    return result

def ask_for_extension():
    root = tk.Tk()
    root.withdraw()  
    ext = simpledialog.askstring("File Extension", "Enter the file extension you want to process (e.g., mp4, mov):")
    if not ext:
        messagebox.showwarning("Missing Extension", "No extension entered. Exiting.")
        root.quit()  
        root.destroy()  
        exit()
    ext = ext.strip().lower()
    if not ext.startswith('.'):
        ext = '.' + ext
    root.quit()  
    root.destroy() 
    return ext

if __name__ == '__main__':
    source_folder = get_folder_path("Select the source folder")
    destination_folder = get_folder_path("Select the destination folder")

    preserve_subfolders = ask_preserve_subfolders()

    extension_to_process = ask_for_extension()

    copied_files = 0
    omitted_files = 0
    copied_filenames = []
    omitted_filenames = []

    copied_files_excel_path = os.path.join(destination_folder, "Files_Copied.xlsx")
    omitted_files_excel_path = os.path.join(destination_folder, "Files_Omitted.xlsx")

    wb_copied = Workbook()
    ws_copied = wb_copied.active
    ws_copied.title = "Copied Files Metadata"
    ws_copied.append(['Filename', 'Duration (s)', 'Frame Rate (fps)', 'Width', 'Height', 'Aspect Ratio'])

    wb_omitted = Workbook()
    ws_omitted = wb_omitted.active
    ws_omitted.title = "Omitted Files"
    ws_omitted.append(['Filename', 'Extension'])

    for dirpath, dirnames, filenames in os.walk(source_folder):
        for item in filenames:
            item_path = os.path.join(dirpath, item)
            file_ext = os.path.splitext(item)[1].lower()

            if file_ext == extension_to_process:
                if preserve_subfolders:
                    relative_path = os.path.relpath(item_path, source_folder)
                    destination_path = os.path.join(destination_folder, relative_path)
                    os.makedirs(os.path.dirname(destination_path), exist_ok=True)
                else:
                    destination_path = os.path.join(destination_folder, item)
                    os.makedirs(destination_folder, exist_ok=True)

                shutil.copy(item_path, destination_path)
                copied_files += 1
                print(f"Copied: {item_path}")

                copied_filenames.append(item)

                duration, frame_rate, width, height, aspect_ratio = get_video_metadata(item_path)
                ws_copied.append([item, duration, frame_rate, width, height, aspect_ratio])

            else:
                omitted_files += 1
                omitted_filenames.append((item, file_ext))

    wb_copied.save(copied_files_excel_path)

    for omitted_item, omitted_ext in omitted_filenames:
        ws_omitted.append([omitted_item, omitted_ext])

    wb_omitted.save(omitted_files_excel_path)

    print(f"\nProcess complete.")
    print(f"Total files copied: {copied_files}")
    print(f"Total files omitted (not matching '{extension_to_process}'): {omitted_files}")
    print(f"\nList of copied files and metadata saved to '{copied_files_excel_path}'")
    print(f"List of omitted files saved to '{omitted_files_excel_path}'")
