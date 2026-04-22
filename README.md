# Cuderm Print Watcher

A simple and reliable application that watches a folder for text files, cleans them, and sends them to a printer.

Built with love after hours — named after my family:  
**C**hantelle, **U**llyndyss (junior), **D**aniel, **E**li and **R**aniyah Maloy ❤️

---

## What it does

- Monitors a chosen folder for new `.txt` files
- Automatically cleans unnecessary blank lines from the files
- Adds proper spacing before the signature/footer section
- Sends the cleaned file to the selected printer
- Moves the processed file to a "Printed" folder
- Runs silently in the background with a system tray icon

---

## How to Use

1. Run `cuderm-logis-printer.exe`
2. In the setup window, select:
   - **Watch Folder** – where new text files will appear
   - **Printed Folder** – where processed files will be moved
   - **Log File** – location for the activity log
   - **Printer** – your local printer
3. Click **"Start Cuderm Logis Printer"**
4. The application will minimize to the system tray and start working

Simply copy any `.txt` file into the Watch Folder — it will be automatically processed and printed.

---

## Features

- Remembers your settings between updates
- Smart update detection (warns if a newer version is launched while the old one is running)
- Clean system tray icon with status and exit option
- Detailed log file for troubleshooting

---

## Requirements

- Windows 10 or 11
- A working local printer installed on the computer

---

## Files

- `cuderm-logis-printer.ps1` – PowerShell source code
- `cuderm-logis-printer.exe` – Compiled executable (recommended for end users)

---

## Open Source

This project is open source under the **MIT License**.  
You are free to use, modify, and distribute it.

The full source code is available in this repository.

---

## Disclaimer

This tool was developed on a best-effort basis. No formal technical support is provided. Use at your own risk.

---

## Author

Developed by Ullyndyss with love for his family.

---

## License

[MIT License](LICENSE) — see the `LICENSE` file for details.
