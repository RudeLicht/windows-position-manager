# WindowPositionManager

**WindowPositionManager** is a lightweight Windows utility built with Python and `tkinter` that lets you save and restore window positions and sizes for your running applications. It includes a system tray icon, preset manager, and optional auto-start on boot.

---

## ✨ Features

- View all open application windows
- Save position and size presets for specific applications
- Restore all windows to saved positions with one click
- Automatically minimize to tray instead of exiting
- Manage saved presets in a dedicated window
- Optionally launch on system startup

---

## 🛠 Requirements

- Python 3.7 or newer
- Windows OS

Install dependencies:

```bash
pip install -r requirements.txt
```

## 🚀 How to Run
Clone the repository (or download the files).

Make sure Python is installed and available in your PATH.

Install dependencies:

```bash
pip install -r requirements.txt
```

Run the app
```bash
python main.py
```

App data is stored in:
```bash
%LOCALAPPDATA%\WindowPositionManager\
```

## 🔧 Startup on Boot
Enable "Start on boot" from the checkbox in the app. It will register the app under:
```bash
HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run
```

## 📌 Notes
- Only visible windows with titles are listed.
- Presets are saved per application (by process name).
- Some system or protected windows may not respond to position changes.

## ✅ License
MIT License
