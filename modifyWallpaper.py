import tkinter as tk
from tkinter import filedialog
import os
import ctypes
import shutil
from PIL import Image, ImageTk
import requests
from datetime import datetime

# Initialize the main window
root = tk.Tk()
root.title("Wallpaper Manager")
root.geometry("600x500")
root.minsize(400, 300)

SPI_SETDESKWALLPAPER = 20  # Constant for setting wallpaper
selected_image = None  # Store selected image path
selected_image_filename = ""  # Store the selected image filename

# Define the folder for Bing Wallpapers
wp_image_folder = os.path.expandvars(r"%userprofile%\AppData\Local\Microsoft\BingWallpaperApp\WPImages")
theme_wallpaper_path = os.path.expandvars(r"%userprofile%\AppData\Roaming\Microsoft\Windows\Themes\TranscodedWallpaper")

# Create a Frame for Image Preview
image_frame = tk.Frame(root, bg="gray")
image_frame.pack(side="top", fill="both", expand=True)

# Create a Label inside the Frame for displaying the image
image_label = tk.Label(image_frame, text="No image loaded", bg="gray")
image_label.pack(fill="both", expand=True)

# Create Labels for displaying image information
info_label = tk.Label(root, text="No information available", bg="white", font=("Helvetica", 10))
info_label.pack(side="bottom", fill="x")

# Create a Frame for Buttons
button_frame = tk.Frame(root)
button_frame.pack(side="bottom", fill="x", pady=10)

def downloadBingWallpaper():
    global selected_image_filename  # Declare selected_image_filename as global to use it within the function
    """Downloads the Bing Wallpaper of the day and saves it with the start date as the filename."""
    # Fetch the wallpaper of the day from Bing API
    bing_url = "https://www.bing.com/HPImageArchive.aspx?format=js&idx=0&n=1"  # API for Bing wallpaper
    response = requests.get(bing_url)
    data = response.json()
    
    # Extract the image URL and metadata from the API response
    image_url = "https://www.bing.com" + data["images"][0]["url"]
    copyright_info = data["images"][0].get("copyright", "No copyright information available")
    start_date = data["images"][0].get("startdate", "No start date available")
    
    # Create filename based on start date (e.g., 20250310.jpg)
    filename = f"{start_date}.jpg"
    file_path = os.path.join(wp_image_folder, filename)

    # Download the wallpaper image
    img_data = requests.get(image_url).content
    with open(file_path, 'wb') as f:
        f.write(img_data)
    
    print(f"Downloaded Bing Wallpaper as: {file_path}")
    
    # Store the selected image filename (before renaming)
    selected_image_filename = filename
    
    # After downloading, fetch the image resolution and file size
    print(f"Image Resolution: {getImageResolution(file_path)}")
    print(f"File Size: {getFileSize(file_path)} bytes")
    
    # Proceed to preview the image
    previewImage(file_path)  # Preview the downloaded image
    global selected_image
    selected_image = file_path  # Store the selected image path for later use
    filename_label.config(text=f"{selected_image_filename}")  # Display only the filename
    
    # Update the info_label to only show the copyright information
    info_label.config(text=f"Copyright: {copyright_info}")

def saveAsTranscodedWallpaper(image_path):
    """Copies the selected image and renames it to TranscodedWallpaper in the Themes directory."""
    try:
        shutil.copy(image_path, theme_wallpaper_path)
        print(f"Image saved as TranscodedWallpaper at: {theme_wallpaper_path}")
    except Exception as e:
        print(f"Error copying file: {e}")

def previewImage(image_path):
    """Loads and displays the selected image in the app, resizing it to fit."""
    img = Image.open(image_path)
    
    # Get available space (without overlapping buttons)
    window_width = image_frame.winfo_width()
    window_height = image_frame.winfo_height()
    
    # Maintain aspect ratio
    img.thumbnail((window_width, window_height))
    img = ImageTk.PhotoImage(img)
    
    image_label.config(image=img, text="")  # Update label with image
    image_label.image = img  # Keep reference

def setWallpaper():
    """Sets the renamed image as the wallpaper."""
    if selected_image:
        saveAsTranscodedWallpaper(selected_image)  # Save as TranscodedWallpaper when setting the wallpaper
        ctypes.windll.user32.SystemParametersInfoW(SPI_SETDESKWALLPAPER, 0, theme_wallpaper_path, 3)
        print(f"Wallpaper set to: {theme_wallpaper_path}")
    else:
        print("No image selected to set as wallpaper.")

def onResize(event):
    """Adjust image preview when window is resized."""
    if selected_image:
        previewImage(selected_image)

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
        selected_image_filename = os.path.basename(filepath)  # Extract the filename from the selected path
        print(f"Selected file: {selected_image}")
        previewImage(filepath)  # Preview the selected image
        filename_label.config(text=f"{selected_image_filename}")  # Display only the filename
        print(f"Image Resolution: {getImageResolution(filepath)}")
        print(f"File Size: {getFileSize(filepath)} bytes")
        info_label.config(text="No information available")  # Reset information when manually loading image

def getImageResolution(image_path):
    """Returns the resolution of the image in width x height format."""
    with Image.open(image_path) as img:
        return f"{img.width}x{img.height}"

def getFileSize(image_path):
    """Returns the file size of the image in bytes."""
    return os.path.getsize(image_path)

# Create a label to display the selected image filename
filename_label = tk.Label(root, text="No image loaded", font=("Helvetica", 12))
filename_label.pack(side="top", fill="x", pady=5)

# Buttons with darker colors
downloadBingButton = tk.Button(button_frame, text="Download Bing Wallpaper", padx=10, pady=5, fg="white", bg="#b177ff", command=downloadBingWallpaper)
downloadBingButton.pack(side="left", expand=True, padx=5)

loadImageButton = tk.Button(button_frame, text="Load Image", padx=10, pady=5, fg="white", bg="#6699ff", command=loadImage)
loadImageButton.pack(side="left", expand=True, padx=5)

setWallpaperButton = tk.Button(button_frame, text="Set as Wallpaper", padx=10, pady=5, fg="white", bg="#ff7f50", command=setWallpaper)
setWallpaperButton.pack(side="right", expand=True, padx=5)

# Bind resize event
root.bind("<Configure>", onResize)

# Run the application
root.mainloop()
