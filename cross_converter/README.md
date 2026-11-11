# Cross Reference Converter

A desktop utility (powered by `pywebview`) to convert raw cross-reference Excel files into two standardized output formats:
- `<filename>_type1.xlsx`
- `<filename>_type2.xlsx`

---

## 📥 How to Download the Windows Executable (.exe)

You don't need to install Python or build the code manually. Every time you push to this repository, **GitHub Actions automatically compiles the standalone Windows executable (`CrossConverter.exe`)**:

1. Go to your repository on GitHub.
2. Click the **Actions** tab at the top.
3. Click the most recent workflow run (e.g. `Build Windows Executable`).
4. Scroll down to the **Artifacts** section at the bottom of the page.
5. Click **`CrossConverter-Windows-exe`** to download the zip file.
6. Unzip it and double-click `CrossConverter.exe` to use it!

---

## 🚀 Running from Source

If you want to run or test the Python script directly:

### 1. Requirements
- Python 3.10+
- Install dependencies:
  ```bash
  pip install -r requirements.txt
  ```

### 2. Launch GUI
```bash
python cross_converter.py
```

### 3. CLI Mode (Headless / Batch)
You can also run conversions directly from the terminal without launching the GUI:
```bash
python cross_converter.py "path/to/input.xlsx" ["path/to/output_dir"]
```

---

## 📋 Expected Input File Columns

The input Excel spreadsheet should include the following columns (column matching is case-insensitive):

- `PartNumber`
- `CompanyName`
- `PartId`
- `CrossPartNumber`
- `CrossCompanyName`
- `CrossPartID`
- `SourceName`
- `Grade`
- `Lifecycle`
- `GroupDescription`
- `FamilyName`
- `TempGradeCode`
- `TempGrade`
- `RiskScore`
- `CrossLifecycle`
- `CrossGroupDescription`
- `CrossFamilyName`
- `CrossTempGradeCode`
- `CrossTempGrade`
- `CrossRohsStatus`
- `CrossRiskScore`

---

## 🛠️ Building the .exe Manually (Local Windows Machine)

If you have a local Windows machine and want to package the executable locally:
```bash
pip install -r requirements.txt
pyinstaller cross_converter.spec --clean --noconfirm
```
The compiled executable will be located in the `dist/` directory:
`dist/CrossConverter.exe`
