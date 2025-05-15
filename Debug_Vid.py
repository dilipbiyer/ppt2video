import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import os
import pandas as pd

# Initialize Tkinter window
root = tk.Tk()
root.title("Debugging Tool")

images_folder = None
excel_file = None
target_folder = None

# Debug Output Box (Scrollable)
debug_output = scrolledtext.ScrolledText(root, width=80, height=20, wrap=tk.WORD)
debug_output.pack(pady=10)

def log_message(message):
    """Displays debug messages in the UI"""
    debug_output.insert(tk.END, message + "\n")
    debug_output.see(tk.END)  # Auto-scroll to latest message

def select_images_folder():
    """User selects the folder where slide images are stored"""
    global images_folder
    images_folder = filedialog.askdirectory()
    
    if images_folder:
        log_message(f"✅ Selected Images Folder: {images_folder}")
    else:
        log_message("❌ No images folder selected.")

def select_excel_file():
    """User selects an Excel file containing speaker notes"""
    global excel_file
    excel_file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])

    if excel_file:
        log_message(f"✅ Selected Excel File: {os.path.basename(excel_file)}")
    else:
        log_message("❌ No Excel file selected.")

def select_target_folder():
    """User selects the folder where audio and video will be stored"""
    global target_folder
    target_folder = filedialog.askdirectory()

    if target_folder:
        os.makedirs(target_folder, exist_ok=True)  # Ensure folder exists
        log_message(f"✅ Selected Target Folder: {target_folder}")
    else:
        log_message("❌ No target folder selected.")

def run_debug_checks():
    """Comprehensive checks for images, Excel, and output folder."""

    debug_output.delete("1.0", tk.END)  # Clear previous messages
    log_message("🔍 Running Debug Checks...\n")

    # ✅ Check if Images Folder Exists
    if not images_folder or not os.path.exists(images_folder):
        log_message(f"❌ ERROR: Images folder does not exist -> {images_folder}")
        return

    # ✅ Check if Excel File Exists
    if not excel_file or not os.path.exists(excel_file):
        log_message(f"❌ ERROR: Excel file does not exist -> {excel_file}")
        return

    # ✅ Check if Target Folder Exists
    if not target_folder or not os.path.exists(target_folder):
        log_message(f"⚠️ Target folder missing. Creating -> {target_folder}")
        os.makedirs(target_folder, exist_ok=True)

    # ✅ List Images & Check Readability (Case-insensitive extension handling)
    slides = sorted([os.path.join(images_folder, f) for f in os.listdir(images_folder) if f.lower().endswith(".png") or f.lower().endswith(".jpg")])
    log_message(f"📷 Found {len(slides)} slide images.")
    if not slides:
        log_message("❌ ERROR: No images detected. Ensure correct folder selection.")
        return
    for img in slides:
        if not os.path.exists(img):
            log_message(f"❌ ERROR: Missing image file -> {img}")

    # ✅ Read Speaker Notes from Excel
    try:
        df = pd.read_excel(excel_file, usecols="A", skiprows=1)  # Read column A starting from A2
        speaker_notes = df.iloc[:, 0].dropna().tolist()  # Remove empty rows
    except Exception as e:
        log_message(f"❌ ERROR: Failed to read Excel file -> {str(e)}")
        return

    log_message(f"📖 Found {len(speaker_notes)} speaker notes.")
    if not speaker_notes:
        log_message("❌ ERROR: No speaker notes detected. Check the Excel sheet formatting.")
        return

    # ✅ Verify Matching Count Between Slides & Speaker Notes
    if len(slides) != len(speaker_notes):
        log_message(f"❌ ERROR: Mismatch detected! {len(slides)} images vs {len(speaker_notes)} speaker notes.")
        return
    log_message("✅ Images and speaker notes count match!")

    # ✅ Validate Target Folder Writing Ability
    test_file = os.path.join(target_folder, "test_write.txt")
    try:
        with open(test_file, "w") as f:
            f.write("Test file creation successful.")
        os.remove(test_file)
        log_message("✅ Target folder writable.")
    except Exception as e:
        log_message(f"❌ ERROR: Unable to write files in target folder -> {str(e)}")
        return

    log_message("\n✅ Debugging Complete: All checks passed!")

# UI Components
tk.Button(root, text="Select Images Folder", command=select_images_folder).pack(pady=5)
tk.Button(root, text="Select Excel File", command=select_excel_file).pack(pady=5)
tk.Button(root, text="Select Target Folder", command=select_target_folder).pack(pady=5)
tk.Button(root, text="Run Debug Checks", command=run_debug_checks).pack(pady=10)

root.mainloop()
