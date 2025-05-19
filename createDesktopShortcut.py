import os
import subprocess

# Check and install winshell if needed
try:
    import winshell
    print("winshell is installed.")
except ImportError:
    print("winshell is NOT installed.")
    subprocess.check_call([os.sys.executable, "-m", "pip", "install", "winshell"])
    import winshell  # Re-import after installation
    
# Check and install pywin32 if needed
try:
    import win32com.client
    print("pywin32 is installed.")
except ImportError:
    print("pywin32 is NOT installed.")
    subprocess.check_call([os.sys.executable, "-m", "pip", "install", "pywin32"])
    import win32com.client  # Re-import after installation

from win32com.client import Dispatch

def create_shortcut_to_current_script(icon_idx = 221):
    # Get the current script's path
    current_script = os.path.abspath(__file__)
    
    # Define the shortcut name
    shortcut_name = os.path.basename(current_script).replace(".py", "")  # Removing .py extension for the shortcut name
    
    # Create the shortcut path
    desktop = winshell.desktop()
    shortcut_path = os.path.join(desktop, f"{shortcut_name}.lnk")
    
    # If the shortcut doesn't exist, create it
    if not os.path.exists(shortcut_path):
        # Icon location (using the provided icon index from SHELL32.dll)
        shell32_path = os.path.join(os.environ["SystemRoot"], "System32", "SHELL32.dll")
        icon_location = f"{shell32_path}, {icon_idx}"

        try:
            # Create the shortcut
            shell = Dispatch("WScript.Shell")
            shortcut = shell.CreateShortCut(shortcut_path)
            shortcut.TargetPath = current_script
            shortcut.WorkingDirectory = os.path.dirname(current_script)
            shortcut.IconLocation = icon_location
            shortcut.save()

            print(f"Shortcut created successfully with icon index {icon_idx}.")
        
        except Exception as e:
            print(f"Error creating shortcut: {e}")
    else:
        print(f"The shortcut {shortcut_name}.lnk already exists.")

# Example usage: Call the function with a specific icon index
create_shortcut_to_current_script(141)  # You can change this index
