import tkinter as tk
from tkinter import filedialog
import os
import ctypes
import shutil
from PIL import Image, ImageTk
import requests
import re
import json

# Initialize the main window
root = tk.Tk()
root.title("Wallpaper Manager")
root.geometry("800x600")
root.minsize(500, 400)

SPI_SETDESKWALLPAPER = 20  # Constant for setting wallpaper
selected_image = None  # Store selected image path
selected_image_filename = ""  # Store the selected image filename

# Define the folder for Bing Wallpapers
wp_image_folder = os.path.expandvars(r"%userprofile%\Pictures")
theme_wallpaper_path = os.path.expandvars(r"%userprofile%\AppData\Roaming\Microsoft\Windows\Themes\TranscodedWallpaper")

# Configure the root grid layout
root.rowconfigure(0, weight=1)  # Image Frame (Expandable)
root.rowconfigure(1, weight=0)  # Filename Label
root.rowconfigure(2, weight=0)  # Info Label
root.rowconfigure(3, weight=0)  # Button Frame
root.columnconfigure(0, weight=1)

# Create a Frame for Image Preview
image_frame = tk.Frame(root, bg="gray")
image_frame.grid(row=0, column=0, sticky="nsew")

# Create a Label inside the Frame for displaying the image
image_label = tk.Label(image_frame, text="No image loaded", bg="gray")
image_label.pack(expand=True)

# Create a label to display the selected image filename
filename_label = tk.Label(root, text="No image loaded", font=("Helvetica", 12))
filename_label.grid(row=1, column=0, sticky="ew", pady=5)

# Create Labels for displaying image information
info_label = tk.Label(root, text="No information available", bg="white", font=("Helvetica", 10))
info_label.grid(row=2, column=0, sticky="ew")

# Create a Frame for Buttons
button_frame = tk.Frame(root)
button_frame.grid(row=3, column=0, sticky="ew", pady=10)
button_frame.columnconfigure([0, 1, 2], weight=1)

resize_delay = None  # Global variable to track resize events


def downloadBingWallpaper():
    """Downloads Bing's wallpaper of the day and saves it."""
    global selected_image_filename  

    bing_url = "https://www.bing.com/HPImageArchive.aspx?format=js&idx=0&n=1"
    response = requests.get(bing_url)
    data = response.json()

    image_url = "https://www.bing.com" + data["images"][0]["url"]
    image_url_4k = re.sub(r"_\d+x\d+\.jpg", "_UHD.jpg", image_url)
    copyright_info = data["images"][0].get("copyright", "No copyright information available")
    start_date = data["images"][0].get("startdate", "No start date available")
    
    filename = f"{start_date}.jpg"
    file_path = os.path.join(wp_image_folder, filename)

    img_data = requests.get(image_url_4k).content
    with open(file_path, 'wb') as f:
        f.write(img_data)

    selected_image_filename = filename
    previewImage(file_path)

    global selected_image
    selected_image = file_path  
    filename_label.config(text=f"{selected_image_filename}")  
    info_label.config(text=f"Copyright: {copyright_info}")

def saveAsTranscodedWallpaper(image_path):
    """Copies the selected image and renames it to TranscodedWallpaper."""
    try:
        shutil.copy(image_path, theme_wallpaper_path)
        print(f"Image saved as TranscodedWallpaper at: {theme_wallpaper_path}")
    except Exception as e:
        print(f"Error copying file: {e}")

def previewImage(image_path):
    """Loads and displays the selected image while keeping aspect ratio."""
    img = Image.open(image_path)

    # Get available space inside the image_frame
    image_frame.update_idletasks()
    frame_width = image_frame.winfo_width()
    frame_height = image_frame.winfo_height()

    # Prevent extreme values (like fullscreen issues)
    if frame_width < 50 or frame_height < 50:
        return

    # Aspect ratio calculation
    img_ratio = img.width / img.height
    frame_ratio = frame_width / frame_height

    if img_ratio > frame_ratio:
        new_width = frame_width
        new_height = int(frame_width / img_ratio)
    else:
        new_height = frame_height
        new_width = int(frame_height * img_ratio)

    # Resize image properly
    img = img.resize((new_width, new_height), Image.LANCZOS)
    img = ImageTk.PhotoImage(img)

    # Apply image
    image_label.config(image=img, text="")
    image_label.image = img  

def setWallpaper():
    """Sets the selected image as the wallpaper."""
    if selected_image:
        saveAsTranscodedWallpaper(selected_image)
        ctypes.windll.user32.SystemParametersInfoW(SPI_SETDESKWALLPAPER, 0, theme_wallpaper_path, 3)
        print(f"Wallpaper set to: {theme_wallpaper_path}")
    else:
        print("No image selected to set as wallpaper.")

def onResize(event):
    """Adjust image preview when window is resized."""
    global resize_delay
    if resize_delay:
        root.after_cancel(resize_delay)  # Cancel previous scheduled update
    
    # Schedule an update after 300ms (adjust delay if needed)
    resize_delay = root.after(300, lambda: previewImage(selected_image) if selected_image else None)

def loadImage():
    """Manually load an image using a file dialog and preview it."""
    global selected_image
    filepath = filedialog.askopenfilename(
        initialdir=wp_image_folder,
        title="Select an Image", 
        filetypes=[("JPEG", "*.jpg;*.jpeg"), ("PNG", "*.png"), ("All Files", "*.*")]
    )
    
    if filepath:
        selected_image = filepath
        selected_image_filename = os.path.basename(filepath)
        previewImage(filepath)  
        filename_label.config(text=f"{selected_image_filename}")  
        info_label.config(text="No information available")

# Buttons (Grid-based to keep them visible in full-screen mode)
downloadBingButton = tk.Button(button_frame, text="Download Bing Wallpaper", fg="white", bg="#b177ff", command=downloadBingWallpaper)
downloadBingButton.grid(row=0, column=0, sticky="ew", padx=5)

loadImageButton = tk.Button(button_frame, text="Load Image", fg="white", bg="#6699ff", command=loadImage)
loadImageButton.grid(row=0, column=1, sticky="ew", padx=5)

setWallpaperButton = tk.Button(button_frame, text="Set as Wallpaper", fg="white", bg="#ff7f50", command=setWallpaper)
setWallpaperButton.grid(row=0, column=2, sticky="ew", padx=5)

# Bind resize event
root.bind("<Configure>", onResize)

# Run the application
root.mainloop()
