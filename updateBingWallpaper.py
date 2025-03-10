import os
import shutil
import requests
from datetime import datetime
import ctypes
from PIL import Image
import re
import json

# Constants for setting wallpaper
SPI_SETDESKWALLPAPER = 20  # Constant for setting wallpaper
# wp_image_folder = os.path.expandvars(r"%userprofile%\AppData\Local\Microsoft\BingWallpaperApp\WPImages")
wp_image_folder = os.path.expandvars(r"%userprofile%\Pictures")
theme_wallpaper_path = os.path.expandvars(r"%userprofile%\AppData\Roaming\Microsoft\Windows\Themes\TranscodedWallpaper")

def download_bing_wallpaper():
    """Downloads the Bing Wallpaper of the day and saves it with the start date as the filename."""
    # Fetch the wallpaper of the day from Bing API
    bing_url = "https://www.bing.com/HPImageArchive.aspx?format=js&idx=0&n=1"  # API for Bing wallpaper
    response = requests.get(bing_url)
    data = response.json()
    print(f"Response Data:\n{json.dumps(data, indent=4)}")

    # Extract the image URL and metadata from the API response
    image_url = "https://www.bing.com" + data["images"][0]["url"]
    print(f"Image url: {image_url}")
    image_url_4k = re.sub(r"_\d+x\d+\.jpg", "_UHD.jpg", image_url)
    start_date = data["images"][0].get("startdate", "No start date available")
    
    # Create filename based on start date (e.g., 20250310.jpg)
    filename = f"{start_date}.jpg"
    file_path = os.path.join(wp_image_folder, filename)

    # Download the wallpaper image
    img_data = requests.get(image_url_4k).content
    with open(file_path, 'wb') as f:
        f.write(img_data)
    
    print(f"Downloaded Bing Wallpaper as: {file_path}")
    return file_path

def save_as_transcoded_wallpaper(image_path):
    """Copies the selected image and renames it to TranscodedWallpaper in the Themes directory."""
    try:
        shutil.copy(image_path, theme_wallpaper_path)
        print(f"Image saved as TranscodedWallpaper at: {theme_wallpaper_path}")
    except Exception as e:
        print(f"Error copying file: {e}")

def set_wallpaper(image_path):
    """Sets the renamed image as the wallpaper."""
    save_as_transcoded_wallpaper(image_path)  # Save as TranscodedWallpaper when setting the wallpaper
    ctypes.windll.user32.SystemParametersInfoW(SPI_SETDESKWALLPAPER, 0, theme_wallpaper_path, 3)
    print(f"Wallpaper set to: {theme_wallpaper_path}")

def main():
    """Main function to download, save, and set the Bing wallpaper."""
    # Download the Bing wallpaper and get the file path
    downloaded_image = download_bing_wallpaper()
    
    # Set the downloaded image as the wallpaper
    set_wallpaper(downloaded_image)

if __name__ == "__main__":
    main()
