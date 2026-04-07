# Serial Checker

**Repository:** [github.com/houssamX911/SerialChecker](https://github.com/houssamX911/SerialChecker)

A Windows desktop tool that collects system, BIOS, board, CPU, disk, GPU, TPM, network, and monitor data in one run, shows it in a dark **Tkinter** UI, and can **export** a full snapshot or **compare** the last export to the live machine.

**Author:** houssamX  
Copyright © houssamX. All rights reserved.

---

## Features

- **Gather** hardware and identity-related fields via a single **PowerShell / CIM** script (WMI-style data).
- **Scrollable** three-column GUI with sections for system info, disks (paged), monitors (paged), GPU, TPM, and network.
- **Export** writes:
  - `serial_report.txt` in the **current working directory** when you run `python app.py`, or **next to `SerialChecker.exe`** when you use the built `.exe`
  - `SERIAL_DUMP_<timestamp>.html` and matching `.json` on your **Desktop**
  - Updates `serial_checker_state.json` next to the script / `.exe` so **Check** knows the latest export
- **Check** loads the last export JSON, gathers fresh data, and writes `SERIAL_COMPARISON_<timestamp>.html` on the Desktop (change detection between snapshots).
- HTML exports use an embedded credit line: **© houssamX**

---

## Requirements

- **Windows** (PowerShell + CIM are used for gathering).
- **Python 3.10+** (3.11+ recommended).
- **Tkinter** — included with the standard Windows Python installer (*tcl/tk*).
- Optional: **`dxdiag`** on PATH for the most accurate DirectX GPU device identifier (the app falls back if it is unavailable).

No third-party pip packages are required to **run** the app. Building an `.exe` needs **PyInstaller** (see below).

---

## How to run

Clone the repo (or download the ZIP from GitHub):

```bash
git clone https://github.com/houssamX911/SerialChecker.git
cd SerialChecker
python app.py
```

If you already have the project folder, from the directory that contains `app.py`:

```bash
python app.py
```

---

## Building a Windows `.exe` (PyInstaller)

This packs `app.py` into a single `SerialChecker.exe` with **no console window** (`--windowed`) and **Tkinter** included.

### One-command build (Windows)

1. Open a terminal **in this project folder** (where `app.py` lives).
2. Double-click **`build_exe.bat`** *or* run the same commands manually:

   ```bash
   pip install -r requirements-build.txt
   pyinstaller --noconfirm --clean --onefile --windowed --name SerialChecker app.py
   ```

3. Find the program here: **`dist\SerialChecker.exe`**

You can copy **only** that `.exe` to another PC or folder. On first run, Windows may start slowly while the one-file bundle unpacks to a temp directory; later starts are usually faster.

**Companion files** (created automatically when you use the app):

| File | Location when using `.exe` |
|------|----------------------------|
| `serial_report.txt` | Same folder as `SerialChecker.exe` |
| `serial_checker_state.json` | Same folder as `SerialChecker.exe` |
| `SERIAL_DUMP_*` / `SERIAL_COMPARISON_*` | Still on your **Desktop** (HTML + JSON dumps) |

Antivirus tools sometimes flag PyInstaller executables as *false positives*. If that happens, sign the binary or submit it as a false positive to your AV vendor.

---

## Usage

1. **Export** — Saves TXT + HTML + JSON; note the paths in the dialog (Desktop + current folder for TXT).
2. **Check** — Run **Export** at least once first. Then **Check** builds an HTML diff report on the Desktop and refreshes the UI.

Do not share exports publicly if they contain sensitive serials, keys, or internal network details.

---

## Repository layout

| File | Purpose |
|------|---------|
| `app.py` | Application entry point and all logic |
| `serial_checker_state.json` | Created locally; stores paths to the latest export (not required in git) |
| `serial_report.txt` | Created next to `app.py`’s cwd (dev) or next to `SerialChecker.exe` (frozen) when you export |
| `build_exe.bat` | Optional script to build `dist\SerialChecker.exe` |
| `requirements-build.txt` | PyInstaller (only for building the `.exe`) |

---

## License / copyright

Copyright © **houssamX**. All rights reserved.  
This project is provided as-is for personal use unless the author publishes a separate license.

---

## Contributing / pushing updates (maintainers)

After changes, from your local clone with `origin` set to this repo:

```bash
git add -A
git commit -m "Describe your change"
git push origin main
```

If you forked instead, open a **Pull Request** to [houssamX/SerialChecker](https://github.com/houssamX911/SerialChecker).
