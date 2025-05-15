import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import pyttsx3
import pandas as pd
from moviepy.editor import *

# Initialize Tkinter window
root = tk.Tk()
root.title("Automated Video Creator")

images_folder = None
excel_file = None
target_folder = None

# User Instruction Label
instruction_label = tk.Label(root, text="Step 1: Select folder containing images.", padx=10, pady=5, fg="blue")
instruction_label.pack()

# Progress Bar
progress = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate")
progress.pack(pady=10)

def update_instruction(new_text):
    """Updates the instruction label for the user"""
    instruction_label.config(text=new_text)

def select_images_folder():
    """User selects the folder where slide images are stored"""
    global images_folder
    images_folder = filedialog.askdirectory()
    
    if images_folder:
        folder_label.config(text=f"Images Folder: {images_folder}")
        update_instruction("Step 2: Select Excel file with speaker notes.")
    else:
        folder_label.config(text="No folder selected")

def select_excel_file():
    """User selects an Excel file containing speaker notes"""
    global excel_file
    excel_file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])

    if excel_file:
        excel_label.config(text=f"Excel File: {os.path.basename(excel_file)}")
        update_instruction("Step 3: Select Target Folder to store generated files.")
    else:
        excel_label.config(text="No file selected")

def select_target_folder():
    """User selects the folder where audio and video will be stored"""
    global target_folder
    target_folder = filedialog.askdirectory()

    if target_folder:
        os.makedirs(target_folder, exist_ok=True)  # Ensure folder exists
        target_label.config(text=f"Target Folder: {target_folder}")
        update_instruction("Step 4: Click 'Generate Video' to create final output.")
    else:
        target_label.config(text="No folder selected")

def generate_video():
    """Generates audio from speaker notes and stitches slides into a video with correct timing"""
    if not images_folder or not excel_file or not target_folder:
        messagebox.showerror("Error", "Ensure images folder, Excel file, and target folder are selected.")
        return

    update_instruction("Generating audio files... Please wait.")

    # Read speaker notes from Excel file
    try:
        df = pd.read_excel(excel_file, usecols="A", skiprows=1)  # Read Column A from A2
        speaker_notes = df.iloc[:, 0].dropna().tolist()  # Remove empty rows
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read Excel file: {str(e)}")
        return

    engine = pyttsx3.init()
    engine.setProperty("rate", 150)  # Reduce speech rate for slower, natural narration

    progress["value"] = 0

    # Detect slide images (case-insensitive)
    slides = sorted([os.path.join(images_folder, f) for f in os.listdir(images_folder) if f.lower().endswith(".png") or f.lower().endswith(".jpg")])

    if len(slides) != len(speaker_notes):
        messagebox.showerror("Error", f"Mismatch detected! {len(slides)} images vs {len(speaker_notes)} speaker notes.")
        return

    audio_files = []
    
    # Generate audio for each slide
    for i, (slide, note) in enumerate(zip(slides, speaker_notes)):
        audio_file = os.path.join(target_folder, f"slide_{i+1}.mp3")
        engine.save_to_file(note, audio_file)
        audio_files.append(audio_file)

        progress["value"] = ((i + 1) / len(slides)) * 50  # Progress halfway for audio generation
        root.update_idletasks()

    engine.runAndWait()
    messagebox.showinfo("Success", "Audio files generated!")

    update_instruction("Stitching video together... Please wait.")

    clips = []
    for img, audio in zip(slides, audio_files):
        if os.path.exists(img) and os.path.exists(audio):
            audio_clip = AudioFileClip(audio)
            slide_duration = audio_clip.duration  # Get the audio duration dynamically
            clip = ImageClip(img, duration=slide_duration).set_audio(audio_clip)
            clips.append(clip)

            # Add a 2-second transition delay
            clips.append(ImageClip(img, duration=2))  

    if not clips:
        messagebox.showerror("Error", "No valid images/audio detected. Ensure files exist.")
        return

    final_video = concatenate_videoclips(clips, method="compose")
    final_video.write_videofile(os.path.join(target_folder, "final_video.mp4"), fps=24)

    progress["value"] = 100  # Completion
    messagebox.showinfo("Success", "Video created successfully!")
    update_instruction("Video generated! You can now find it in the target folder.")

# UI Components
folder_label = tk.Label(root, text="No images folder selected", padx=10, pady=5)
folder_label.pack()
tk.Button(root, text="Select Images Folder", command=select_images_folder).pack(pady=5)

excel_label = tk.Label(root, text="No Excel file selected", padx=10, pady=5)
excel_label.pack()
tk.Button(root, text="Select Excel File", command=select_excel_file).pack(pady=5)

target_label = tk.Label(root, text="No target folder selected", padx=10, pady=5)
target_label.pack()
tk.Button(root, text="Select Target Folder", command=select_target_folder).pack(pady=5)

tk.Button(root, text="Generate Video", command=generate_video).pack(pady=10)

root.mainloop()
