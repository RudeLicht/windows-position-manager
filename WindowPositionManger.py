import tkinter as tk
from tkinter import IntVar, messagebox
import pygetwindow as gw
import psutil
import win32process
import json
import os
import sys
import winreg as reg
import threading
from PIL import Image
import pystray

# ========== AppData Paths ==========
APP_NAME = "WindowPositionManager"
appdata_dir = os.path.join(os.getenv('LOCALAPPDATA'), APP_NAME)
os.makedirs(appdata_dir, exist_ok=True)

preset_path = os.path.join(appdata_dir, "presets.json")
settings_path = os.path.join(appdata_dir, "settings.json")

# ========== Tkinter Window Setup ==========
root = tk.Tk()
root.title(APP_NAME)
root.geometry("400x350")
root.minsize(400, 350)
root.config(bg="#f4f4f9")

# Set the window icon
try:
    root.iconbitmap('image.ico')
except Exception as e:
    print(f"Failed to load window icon: {e}")

# Fonts
font_large = ("Segoe UI", 11, "bold")
font_medium = ("Segoe UI", 9)

# Load presets
presets = {}
if os.path.exists(preset_path):
    with open(preset_path, 'r') as f:
        presets = json.load(f)

# Load settings
settings = {"start_on_boot": False}
if os.path.exists(settings_path):
    with open(settings_path, 'r') as f:
        settings.update(json.load(f))

# Global app list
app_list = []

# ========== Tray Icon Setup ==========
def create_image():
    """ Load your custom ico for the tray icon """
    try:
        image = Image.open("image.ico")
        return image
    except Exception as e:
        print(f"Failed to load tray icon: {e}")
        # fallback to simple generated image
        image = Image.new('RGB', (64, 64), color=(0, 102, 204))
        return image

def on_quit(icon, item):
    icon.stop()
    root.destroy()

def on_restore(icon, item):
    icon.stop()
    root.after(0, root.deiconify)

def setup_tray():
    image = create_image()
    menu = pystray.Menu(
        pystray.MenuItem('Restore', on_restore),
        pystray.MenuItem('Exit', on_quit)
    )
    icon = pystray.Icon(APP_NAME, image, menu=menu)
    icon.run()

def minimize_to_tray():
    root.withdraw()  # Hide the window
    threading.Thread(target=setup_tray, daemon=True).start()

def on_closing():
    minimize_to_tray()

root.protocol("WM_DELETE_WINDOW", on_closing)

# ========== Functions ==========

def get_open_windows():
    global app_list
    windows = gw.getAllWindows()
    app_list = []
    for window in windows:
        if window.title and window.visible:
            try:
                _, pid = win32process.GetWindowThreadProcessId(window._hWnd)
                proc = psutil.Process(pid)
                app_list.append((window, proc.name()))
            except Exception:
                pass
    return app_list

def update_window_listbox():
    current_selection = window_listbox.curselection()
    selected_index = current_selection[0] if current_selection else None
    yview = window_listbox.yview()

    window_listbox.delete(0, tk.END)
    get_open_windows()

    for idx, (window, proc_name) in enumerate(app_list):
        window_listbox.insert(tk.END, f"{proc_name} - {window.title}")

    if selected_index is not None and selected_index < window_listbox.size():
        window_listbox.select_set(selected_index)

    window_listbox.yview_moveto(yview[0])
    root.after(2000, update_window_listbox)

def getCurrentWindowPosition():
    try:
        choice = int(window_listbox.curselection()[0])
        selected_window, selected_proc = app_list[choice]

        presets[selected_proc] = {
            "x": selected_window.left,
            "y": selected_window.top,
            "width": selected_window.width,
            "height": selected_window.height
        }
        with open(preset_path, 'w') as f:
            json.dump(presets, f, indent=4)
        messagebox.showinfo("Success", f"Saved preset for {selected_proc}!")
    except IndexError:
        messagebox.showerror("Error", "Please select a valid window.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def moveAllWindows():
    try:
        for proc_name, pos in presets.items():
            for window in gw.getAllWindows():
                if window.title and window.visible:
                    try:
                        _, pid = win32process.GetWindowThreadProcessId(window._hWnd)
                        proc = psutil.Process(pid)
                        if proc_name.lower() == proc.name().lower():
                            x, y = pos['x'], pos['y']
                            width, height = pos['width'], pos['height']

                            if window.isMinimized:
                                window.restore()

                            window.moveTo(x, y)
                            window.resizeTo(width, height)
                    except Exception:
                        pass
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while moving windows: {e}")
    else:
        messagebox.showinfo("Success", "Successfully moved windows.")

def on_window_select(event):
    try:
        choice = window_listbox.curselection()
        if choice:
            idx = choice[0]
            selected_window, selected_proc = app_list[idx]
            selected_application.config(text=f"Selected: {selected_proc}")
    except Exception:
        selected_application.config(text="Selected: None")

def manage_presets_window():
    manage_window = tk.Toplevel(root)
    manage_window.title("Manage Presets")
    manage_window.minsize(300, 350)
    manage_window.geometry("300x350")
    manage_window.config(bg="#f4f4f9")

    preset_listbox = tk.Listbox(manage_window, width=40, height=15, font=font_medium, bd=2, relief="solid",
                                selectmode=tk.SINGLE, bg="#ffffff", fg="#333333")
    preset_listbox.pack(pady=10)

    def refresh_preset_list():
        preset_listbox.delete(0, tk.END)
        for proc_name in presets.keys():
            preset_listbox.insert(tk.END, proc_name)

    def delete_selected_preset():
        try:
            choice = preset_listbox.curselection()
            if not choice:
                messagebox.showwarning("Warning", "Please select a preset to delete.")
                return
            idx = choice[0]
            selected_proc = preset_listbox.get(idx)

            confirm = messagebox.askyesno("Confirm Delete", f"Delete preset for {selected_proc}?")
            if confirm:
                presets.pop(selected_proc, None)
                with open(preset_path, 'w') as f:
                    json.dump(presets, f, indent=4)
                refresh_preset_list()
                messagebox.showinfo("Success", f"Deleted preset for {selected_proc}.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

    refresh_preset_list()

    delete_button = tk.Button(manage_window, text="Delete Preset", command=delete_selected_preset,
                              font=("Helvetica", 10, "bold"), bg="#f44336", fg="white", width=15, bd=1, relief="solid")
    delete_button.pack(pady=5)

    close_button = tk.Button(manage_window, text="Close", command=manage_window.destroy,
                             font=("Helvetica", 10, "bold"), bg="#2196F3", fg="white", width=15, bd=1, relief="solid")
    close_button.pack(pady=5)

def startOnBootSettings():
    settings["start_on_boot"] = bool(var1.get())
    with open(settings_path, 'w') as f:
        json.dump(settings, f, indent=4)

    if var1.get() == 1:
        addAppToStartOnBoot()
    else:
        removeAppFromStartOnBoot()

def addAppToStartOnBoot():
    try:
        script_path = os.path.realpath(sys.argv[0])
        exe_path = f'"{sys.executable}" "{script_path}"' if script_path.endswith(".py") else f'"{script_path}"'

        app_name = APP_NAME
        key = reg.HKEY_CURRENT_USER
        key_value = r"Software\Microsoft\Windows\CurrentVersion\Run"

        open_key = reg.OpenKey(key, key_value, 0, reg.KEY_ALL_ACCESS)
        reg.SetValueEx(open_key, app_name, 0, reg.REG_SZ, exe_path)
        reg.CloseKey(open_key)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to add to startup: {e}")

def removeAppFromStartOnBoot():
    try:
        app_name = APP_NAME
        key = reg.HKEY_CURRENT_USER
        key_value = r"Software\Microsoft\Windows\CurrentVersion\Run"

        open_key = reg.OpenKey(key, key_value, 0, reg.KEY_ALL_ACCESS)
        reg.DeleteValue(open_key, app_name)
        reg.CloseKey(open_key)
    except FileNotFoundError:
        pass
    except Exception as e:
        messagebox.showerror("Error", f"Failed to remove from startup: {e}")

# ========== GUI Layout ==========

window_listbox = tk.Listbox(root, width=58, height=8, font=font_medium, bd=2, relief="solid",
                            selectmode=tk.SINGLE, bg="#ffffff", fg="#333333")
window_listbox.pack(pady=(15, 5))

selected_application = tk.Label(
    root, text="Selected: None", font=font_large, bg="#f4f4f9", fg="#333333")
selected_application.pack(pady=(5, 10))

# Buttons Frame
button_frame = tk.Frame(root, bg="#f4f4f9")
button_frame.pack(pady=5)

save_button = tk.Button(button_frame, text="Save Preset", command=getCurrentWindowPosition,
                        font=("Helvetica", 10, "bold"), bg="#4CAF50", fg="white", width=18, bd=0.5, relief="solid")
save_button.grid(row=0, column=0, padx=5, pady=5)

move_all_button = tk.Button(button_frame, text="Move Windows", command=moveAllWindows,
                            font=("Helvetica", 10, "bold"), bg="#2196F3", fg="white", width=18, bd=0.5, relief="solid")
move_all_button.grid(row=0, column=1, padx=5, pady=5)

manage_presets_button = tk.Button(button_frame, text="Manage Presets", command=manage_presets_window,
                                  font=("Helvetica", 10, "bold"), bg="#009688", fg="white", width=38, bd=0.5, relief="solid")
manage_presets_button.grid(row=1, column=0, columnspan=2, pady=5)

# Bottom Frame
bottom_frame = tk.Frame(root, bg="#f4f4f9")
bottom_frame.pack(pady=10)

var1 = IntVar()
checkbox = tk.Checkbutton(bottom_frame, text="Start on boot", variable=var1,
                          onvalue=1, offvalue=0, command=startOnBootSettings,
                          bg="#f4f4f9", font=("Helvetica", 10, "bold"), activebackground="grey", bd=0.5, relief="solid", cursor="tcross")
checkbox.grid(row=0, column=0, padx=5)

exit_button = tk.Button(bottom_frame, text="Exit", command=root.quit,
                        font=("Helvetica", 10, "bold"), bg="#f44336", fg="white", width=15, bd=1, relief="solid")
exit_button.grid(row=0, column=1, padx=5)

if settings.get('start_on_boot', False):
    checkbox.select()
else:
    checkbox.deselect()

window_listbox.bind("<<ListboxSelect>>", on_window_select)
update_window_listbox()

root.mainloop()
