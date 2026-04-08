import tkinter as tk
from tkinter import filedialog
import os
import ctypes
import shutil
from PIL import Image, ImageTk
import requests
import re

# ------------------ Main Window ------------------
root = tk.Tk()
root.title("Wallpaper Manager")
root.geometry("800x600")
root.minsize(500, 400)

SPI_SETDESKWALLPAPER = 20

selected_image = None
selected_image_filename = ""

wp_image_folder = os.path.expandvars(
    r"D:\users\NG_Boh_Wai\OneDrive - MSI Global Private Limited\Pictures\wallpaper"
)
theme_wallpaper_path = os.path.expandvars(
    r"%userprofile%\AppData\Roaming\Microsoft\Windows\Themes\TranscodedWallpaper"
)

# ------------------ Layout ------------------
root.rowconfigure(0, weight=1)
root.rowconfigure(1, weight=0)
root.rowconfigure(2, weight=0)
root.rowconfigure(3, weight=0)
root.columnconfigure(0, weight=1)

# ------------------ Canvas (Image Renderer) ------------------
canvas = tk.Canvas(root, bg="gray", highlightthickness=0)
canvas.grid(row=0, column=0, sticky="nsew")

# Image state
original_img = None
canvas_img = None
canvas_img_id = None

# ------------------ Labels ------------------
filename_label = tk.Label(root, text="No image loaded", font=("Helvetica", 12))
filename_label.grid(row=1, column=0, sticky="ew", pady=5)

info_label = tk.Label(root, text="No information available", bg="white", font=("Helvetica", 10))
info_label.grid(row=2, column=0, sticky="ew")

# ------------------ Buttons ------------------
button_frame = tk.Frame(root)
button_frame.grid(row=3, column=0, sticky="ew", pady=10)
button_frame.columnconfigure([0, 1, 2], weight=1)

# ------------------ Canvas Rendering ------------------
def redraw_canvas(event=None):
    """Resize & redraw image live while dragging."""
    global canvas_img, canvas_img_id

    if not original_img:
        return

    cw = canvas.winfo_width()
    ch = canvas.winfo_height()

    if cw < 20 or ch < 20:
        return

    img_ratio = original_img.width / original_img.height
    canvas_ratio = cw / ch

    if img_ratio > canvas_ratio:
        new_w = cw
        new_h = int(cw / img_ratio)
    else:
        new_h = ch
        new_w = int(ch * img_ratio)

    resized = original_img.resize((new_w, new_h), Image.BILINEAR)
    canvas_img = ImageTk.PhotoImage(resized)

    cx, cy = cw // 2, ch // 2

    if canvas_img_id:
        canvas.itemconfig(canvas_img_id, image=canvas_img)
        canvas.coords(canvas_img_id, cx, cy)
    else:
        canvas_img_id = canvas.create_image(cx, cy, image=canvas_img)

# Redraw during live resize
canvas.bind("<Configure>", redraw_canvas)

# ------------------ Image Load ------------------
def load_image(path):
    global original_img
    original_img = Image.open(path)
    redraw_canvas()

# ------------------ Bing Download ------------------
def downloadBingWallpaper():
    global selected_image, selected_image_filename

    bing_url = "https://www.bing.com/HPImageArchive.aspx?format=js&idx=0&n=1"
    data = requests.get(bing_url).json()

    image_url = "https://www.bing.com" + data["images"][0]["url"]
    image_url_4k = re.sub(r"_\d+x\d+\.jpg", "_UHD.jpg", image_url)

    filename = f"{data['images'][0].get('startdate','bing')}.jpg"
    path = os.path.join(wp_image_folder, filename)

    with open(path, "wb") as f:
        f.write(requests.get(image_url_4k).content)

    selected_image = path
    selected_image_filename = filename

    load_image(path)
    filename_label.config(text=filename)
    info_label.config(text=data['images'][0].get("copyright",""))

# ------------------ Manual Load ------------------
def loadImage():
    global selected_image, selected_image_filename

    path = filedialog.askopenfilename(
        initialdir=wp_image_folder,
        title="Select Image",
        filetypes=[("Images", "*.jpg;*.jpeg;*.png"), ("All", "*.*")]
    )
    if path:
        selected_image = path
        selected_image_filename = os.path.basename(path)
        load_image(path)
        filename_label.config(text=selected_image_filename)
        info_label.config(text="No information available")

# ------------------ Wallpaper Set ------------------
def saveAsTranscodedWallpaper(image_path):
    shutil.copy(image_path, theme_wallpaper_path)

def setWallpaper():
    if selected_image:
        saveAsTranscodedWallpaper(selected_image)
        ctypes.windll.user32.SystemParametersInfoW(
            SPI_SETDESKWALLPAPER, 0, theme_wallpaper_path, 3
        )

# ------------------ Buttons ------------------
tk.Button(
    button_frame, text="Download Bing Wallpaper",
    bg="#b177ff", fg="white", command=downloadBingWallpaper
).grid(row=0, column=0, sticky="ew", padx=5)

tk.Button(
    button_frame, text="Load Image",
    bg="#6699ff", fg="white", command=loadImage
).grid(row=0, column=1, sticky="ew", padx=5)

tk.Button(
    button_frame, text="Set as Wallpaper",
    bg="#ff7f50", fg="white", command=setWallpaper
).grid(row=0, column=2, sticky="ew", padx=5)

# ------------------ Run ------------------
root.mainloop()