# Stronghold File Editor

A desktop utility built with PyQt6 for curating batches of documents before merging or exporting them. Drop in PDF, TIFF, or JPG files, reorder and rename them inline, then merge everything into a single PDF or export every page/frame as JPG.

## Prerequisites
- Python 3.9+ (Python 3.11 is recommended for PyQt6)
- Git (if you plan to clone the repository)
- Windows, macOS, or Linux capable of running PyQt6 applications

## 1. Create and Activate a Virtual Environment
```bash
python -m venv .venv
```

- **Windows PowerShell**
  ```powershell
  .\.venv\Scripts\Activate.ps1
  ```
- **macOS/Linux (bash/zsh)**
  ```bash
  source .venv/bin/activate
  ```

Upgrade pip (optional but recommended):
```bash
python -m pip install --upgrade pip
```

## 2. Install Runtime Dependencies
```bash
pip install -r requirements.txt
```
This pulls in PyQt6 for the UI plus libraries used for PDF merging and conversions (PyPDF2, Pillow, PyMuPDF).

## 3. Launch the Application
From the project root:
```bash
python main.py
```
The Stronghold File Editor window should open. Use the **Add Files** button or drag-and-drop to populate the list, then reorganise, rename, merge into a single PDF, or export as needed.

Exporting converts PDF pages and TIFF frames to JPG images and copies any existing JPG inputs into the destination folder.

## 4. Optional: Build a Stand-alone Windows Executable
PyInstaller is used to bundle the app and images into a single executable.

1. Install PyInstaller inside the virtual environment (one-time step):
   ```bash
   pip install pyinstaller
   ```
2. Run the provided build specification:
   ```bash
   pyinstaller StrongholdFileEditor.spec
   ```
3. The packaged app appears under `dist/StrongholdFileEditor/StrongholdFileEditor.exe`. Copy the `dist/StrongholdFileEditor` directory wherever you need and launch the `.exe`.

If you alter resources (icons, images), update `StrongholdFileEditor.spec` accordingly before rebuilding.

## 5. Deactivating the Virtual Environment
When you finish:
```bash
deactivate
```

## Troubleshooting Tips
- **Missing DLLs on Windows:** Ensure Visual C++ redistributables for your Python version are installed (Python.org installers ship with them).
- **PyMuPDF errors:** Some PDF files require additional fonts. Installing system fonts can resolve rendering issues.
- **PyInstaller antivirus flags:** Signing the executable or adding the build folder to antivirus exclusions typically clears false positives.

Enjoy managing your document batches with Stronghold File Editor!
