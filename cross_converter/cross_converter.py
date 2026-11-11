"""
Cross Reference Converter  (pywebview UI)

Install:   pip install pywebview pandas openpyxl
Run:       python cross_converter.py

Reads an Excel file with the columns:
ProductName, PartId, PartNumber, CompanyName, Lifecycle, GroupDescription, FamilyName,
TempGradeCode, TempGrade, OrgPartRisk, RiskScore, CrossPartID, CrossPartNumber,
CrossCompanyName, CrossLifecycle, CrossGroupDescription, CrossFamilyName,
CrossTempGradeCode, CrossTempGrade, CrossRohsStatus, CrossRiskScore, Grade, DisplayOrder,
FroecedCrossGrade, ManualRank, SourceName, Type, GroupId, IsApproved, Action, actions

and writes two Excel files:
  <name>_type1.xlsx  and  <name>_type2.xlsx
"""

import os
import sys
import subprocess
import traceback

import pandas as pd

try:
    import webview
except ImportError:
    webview = None

# ---------------------------------------------------------------------------
# MAPPINGS:  output header  ->  source column in the input file
# Use None for a header that has no source column (it is written empty).
# Edit these if you want a different mapping.
# ---------------------------------------------------------------------------
TYPE1_MAP = {
    "PartNumberC": "PartNumber",
    "CompanyNameC": "CompanyName",
    "LatestCompanyC": "CompanyName",
    "PartIDc": "PartId",
    "PartNumberX": "CrossPartNumber",
    "CompanyNamex": "CrossCompanyName",
    "LatestCompanyX": "CrossCompanyName",
    "PartIDx": "CrossPartID",
    "FeatureName": None,
    "SourceName": "SourceName",
}

TYPE2_MAP = {
    "ParIDC": "PartId",
    "PartNumberC": "PartNumber",
    "CompanyNameC": "CompanyName",
    "LetestCompanyName": "CompanyName",
    "PLc": None,
    "ParIDX": "CrossPartID",
    "PartNumberX": "CrossPartNumber",
    "CompanyNamex": "CrossCompanyName",
    "LatestCompanyX": "CrossCompanyName",
    "PLx": None,
    "CrossGrade": "Grade",
    "SourceName": "SourceName",
    "Lifecycle": "Lifecycle",
    "GroupDescription": "GroupDescription",
    "FamilyName": "FamilyName",
    "TempGradeCode": "TempGradeCode",
    "TempGrade": "TempGrade",
    "RohsStatus": None,
    "RiskScore": "RiskScore",
    "CrossLifecycle": "CrossLifecycle",
    "CrossGroupDescription": "CrossGroupDescription",
    "CrossFamilyName": "CrossFamilyName",
    "CrossTempGradeCode": "CrossTempGradeCode",
    "CrossTempGrade": "CrossTempGrade",
    "CrossRohsStatus": "CrossRohsStatus",
    "CrossRiskScore": "CrossRiskScore",
}


def resolve_columns(df_columns, needed_columns):
    """
    Match needed columns against actual DataFrame columns.
    Matches exact column names first, then case-insensitively (trimmed).
    Returns (resolved_dict, missing_list).
    """
    actual_cols = {str(c).strip(): c for c in df_columns}
    lower_map = {str(c).strip().lower(): c for c in df_columns}

    resolved = {}
    missing = []
    for req in needed_columns:
        if req in actual_cols:
            resolved[req] = actual_cols[req]
        elif req.lower() in lower_map:
            resolved[req] = lower_map[req.lower()]
        else:
            missing.append(req)
    return resolved, sorted(missing)


def build_frame(df, mapping, resolved=None):
    out = pd.DataFrame(index=df.index)
    for target, source in mapping.items():
        if source:
            actual_col = resolved.get(source, source) if resolved else source
            out[target] = df[actual_col]
        else:
            out[target] = ""
    return out


def convert(input_path, output_dir):
    if not os.path.isfile(input_path):
        raise FileNotFoundError(f"Input file not found: {input_path}")

    # Read excel (preserves raw strings and IDs without converting to float)
    df = pd.read_excel(input_path, dtype=object)

    needed = {s for s in {**TYPE1_MAP, **TYPE2_MAP}.values() if s}
    resolved, missing = resolve_columns(df.columns, needed)
    if missing:
        raise ValueError("Missing columns in input file: " + ", ".join(missing))

    base = os.path.splitext(os.path.basename(input_path))[0]
    os.makedirs(output_dir, exist_ok=True)
    p1 = os.path.join(output_dir, f"{base}_type1.xlsx")
    p2 = os.path.join(output_dir, f"{base}_type2.xlsx")

    try:
        build_frame(df, TYPE1_MAP, resolved).to_excel(p1, index=False)
        build_frame(df, TYPE2_MAP, resolved).to_excel(p2, index=False)
    except PermissionError as e:
        raise PermissionError(
            f"Permission denied writing output files: {e}. "
            f"If '{os.path.basename(p1)}' or '{os.path.basename(p2)}' is open in Excel, please close it and try again."
        ) from e

    return len(df), p1, p2


# ---------------------------------------------------------------------------
# Non-blocking Native File & Folder Dialogs (Avoids WebView2 COM Deadlock)
# ---------------------------------------------------------------------------
def pick_file_tkinter(title="Select Excel File"):
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    root.focus_force()
    path = filedialog.askopenfilename(
        parent=root,
        title=title,
        filetypes=[
            ("Excel files", "*.xlsx;*.xlsm;*.xls"),
            ("All files", "*.*"),
        ],
    )
    root.destroy()
    return (path or "").strip()


def pick_folder_tkinter(title="Select Output Folder"):
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    root.focus_force()
    path = filedialog.askdirectory(parent=root, title=title)
    root.destroy()
    return (path or "").strip()


def pick_file_powershell(title="Select Excel File"):
    try:
        ps_cmd = (
            "[System.Reflection.Assembly]::LoadWithPartialName('System.windows.forms') | Out-Null; "
            "$f = New-Object System.Windows.Forms.OpenFileDialog; "
            "$f.Filter = 'Excel Files (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls|All Files (*.*)|*.*'; "
            f"$f.Title = '{title}'; "
            "if ($f.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { Write-Output $f.FileName }"
        )
        res = subprocess.run(
            ["powershell", "-NoProfile", "-Command", ps_cmd],
            capture_output=True,
            text=True,
            timeout=120,
        )
        return res.stdout.strip()
    except Exception:
        return ""


def pick_folder_powershell(title="Select Output Folder"):
    try:
        ps_cmd = (
            "[System.Reflection.Assembly]::LoadWithPartialName('System.windows.forms') | Out-Null; "
            "$f = New-Object System.Windows.Forms.FolderBrowserDialog; "
            f"$f.Description = '{title}'; "
            "if ($f.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { Write-Output $f.SelectedPath }"
        )
        res = subprocess.run(
            ["powershell", "-NoProfile", "-Command", ps_cmd],
            capture_output=True,
            text=True,
            timeout=120,
        )
        return res.stdout.strip()
    except Exception:
        return ""


def select_file():
    # 1. Try tkinter (clean native dialog, no WebView2 conflict)
    try:
        p = pick_file_tkinter()
        if p:
            return p
    except Exception:
        pass

    # 2. Try PowerShell fallback on Windows
    if sys.platform == "win32" or os.name == "nt":
        try:
            p = pick_file_powershell()
            if p:
                return p
        except Exception:
            pass

    return ""


def select_folder():
    # 1. Try tkinter
    try:
        p = pick_folder_tkinter()
        if p:
            return p
    except Exception:
        pass

    # 2. Try PowerShell fallback on Windows
    if sys.platform == "win32" or os.name == "nt":
        try:
            p = pick_folder_powershell()
            if p:
                return p
        except Exception:
            pass

    return ""


class Api:
    def __init__(self):
        self.window = None

    def pick_input(self):
        try:
            return select_file()
        except Exception as e:
            traceback.print_exc()
            return ""

    def pick_output_dir(self):
        try:
            return select_folder()
        except Exception as e:
            traceback.print_exc()
            return ""

    def run(self, input_path, output_dir):
        try:
            input_path = (input_path or "").strip().strip('"').strip("'")
            output_dir = (output_dir or "").strip().strip('"').strip("'")

            if not input_path or not os.path.isfile(input_path):
                return {"ok": False, "msg": f"Please choose a valid input Excel file: {input_path}"}
            if not output_dir:
                output_dir = os.path.dirname(input_path) or os.getcwd()

            rows, p1, p2 = convert(input_path, output_dir)
            return {"ok": True, "msg": f"Success! {rows} rows processed.", "files": [p1, p2]}
        except Exception as e:
            traceback.print_exc()
            return {"ok": False, "msg": str(e)}


HTML = """<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Cross Reference Converter</title>
<style>
  :root {
    --bg: #f4f6fa;
    --card: #ffffff;
    --text: #1d2433;
    --muted: #6b7385;
    --accent: #2563eb;
    --accent-hover: #1d4ed8;
    --ok: #15803d;
    --ok-bg: #dcfce7;
    --err: #b91c1c;
    --err-bg: #fee2e2;
    --border: #dde2ec;
  }
  @media (prefers-color-scheme: dark) {
    :root {
      --bg: #0f172a;
      --card: #1e293b;
      --text: #f1f5f9;
      --muted: #94a3b8;
      --accent: #3b82f6;
      --accent-hover: #60a5fa;
      --ok: #4ade80;
      --ok-bg: #064e3b;
      --err: #f87171;
      --err-bg: #7f1d1d;
      --border: #334155;
    }
  }
  * { box-sizing: border-box; }
  body {
    margin: 0;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
    background: var(--bg);
    color: var(--text);
    padding: 32px 20px;
    display: flex;
    justify-content: center;
  }
  .card {
    width: 100%;
    max-width: 680px;
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 12px;
    box-shadow: 0 4px 12px rgba(0,0,0,0.05);
    padding: 32px;
  }
  h1 { margin: 0 0 6px; font-size: 22px; font-weight: 700; }
  p.sub { margin: 0 0 24px; color: var(--muted); font-size: 14px; }
  label { display: block; font-size: 13px; font-weight: 600; margin: 18px 0 6px; }
  .row { display: flex; gap: 8px; }
  input {
    flex: 1;
    padding: 10px 14px;
    border: 1px solid var(--border);
    border-radius: 8px;
    background: var(--bg);
    color: var(--text);
    font-size: 13px;
  }
  input:focus { outline: 2px solid var(--accent); }
  button {
    padding: 10px 18px;
    border: 0;
    border-radius: 8px;
    background: var(--accent);
    color: #fff;
    font-size: 13px;
    font-weight: 600;
    cursor: pointer;
    transition: background 0.15s ease-in-out;
    white-space: nowrap;
  }
  button:hover { background: var(--accent-hover); }
  button.secondary {
    background: transparent;
    color: var(--accent);
    border: 1px solid var(--accent);
  }
  button.secondary:hover {
    background: rgba(37, 99, 235, 0.08);
  }
  button.go {
    width: 100%;
    margin-top: 26px;
    padding: 13px;
    font-size: 15px;
  }
  button:disabled { opacity: 0.55; cursor: not-allowed; }
  #status {
    margin-top: 20px;
    padding: 14px;
    border-radius: 8px;
    font-size: 13px;
    line-height: 1.5;
    white-space: pre-wrap;
    word-break: break-all;
    display: none;
  }
  #status.ok {
    display: block;
    color: var(--ok);
    background: var(--ok-bg);
    border: 1px solid var(--ok);
  }
  #status.err {
    display: block;
    color: var(--err);
    background: var(--err-bg);
    border: 1px solid var(--err);
  }
</style>
</head>
<body>
<div class="card">
  <h1>Cross Reference Converter</h1>
  <p class="sub">Convert raw cross reference Excel spreadsheets into Type 1 and Type 2 formats.</p>

  <label for="inp">Input Excel File (*.xlsx, *.xlsm, *.xls)</label>
  <div class="row">
    <input id="inp" placeholder="Click Browse or paste Excel file path...">
    <button class="secondary" type="button" id="btnInp" onclick="pickIn()">Browse</button>
  </div>

  <label for="out">Output Folder (optional, defaults to same folder as input)</label>
  <div class="row">
    <input id="out" placeholder="Same folder as input (or choose custom folder)...">
    <button class="secondary" type="button" id="btnOut" onclick="pickOut()">Browse</button>
  </div>

  <button class="go" id="go" type="button" onclick="run()">Convert</button>
  <div id="status"></div>
</div>

<script>
async function pickIn() {
  const btn = document.getElementById('btnInp');
  btn.disabled = true;
  btn.textContent = 'Opening...';
  try {
    const p = await pywebview.api.pick_input();
    if (p) {
      document.getElementById('inp').value = p;
      const out = document.getElementById('out');
      if (!out.value) {
        const sep = p.includes('\\\\') ? '\\\\' : '/';
        const idx = p.lastIndexOf(sep);
        if (idx > 0) out.placeholder = p.substring(0, idx);
      }
    }
  } catch (err) {
    console.error(err);
  } finally {
    btn.disabled = false;
    btn.textContent = 'Browse';
  }
}

async function pickOut() {
  const btn = document.getElementById('btnOut');
  btn.disabled = true;
  btn.textContent = 'Opening...';
  try {
    const p = await pywebview.api.pick_output_dir();
    if (p) document.getElementById('out').value = p;
  } catch (err) {
    console.error(err);
  } finally {
    btn.disabled = false;
    btn.textContent = 'Browse';
  }
}

async function run() {
  const btn = document.getElementById('go');
  const st = document.getElementById('status');
  btn.disabled = true;
  st.style.display = 'block';
  st.className = '';
  st.textContent = 'Processing, please wait...';

  try {
    const inp = document.getElementById('inp').value;
    const out = document.getElementById('out').value;
    const r = await pywebview.api.run(inp, out);

    if (r && r.ok) {
      st.className = 'ok';
      let html = r.msg + '\\n\\nGenerated Files:\\n';
      r.files.forEach(f => { html += '• ' + f + '\\n'; });
      st.textContent = html;
    } else {
      st.className = 'err';
      st.textContent = 'Error: ' + ((r && r.msg) ? r.msg : 'Unknown conversion error.');
    }
  } catch (err) {
    st.className = 'err';
    st.textContent = 'Unexpected application error: ' + err;
  } finally {
    btn.disabled = false;
  }
}
</script>
</body>
</html>
"""


def main():
    # CLI support: python cross_converter.py <input_excel> [output_dir]
    if len(sys.argv) > 1 and not sys.argv[1].startswith("-"):
        in_file = sys.argv[1]
        out_dir = sys.argv[2] if len(sys.argv) > 2 else os.path.dirname(os.path.abspath(in_file))
        print(f"Running conversion on: {in_file}")
        try:
            rows, p1, p2 = convert(in_file, out_dir)
            print(f"Successfully processed {rows} rows!")
            print(f"  Type 1: {p1}")
            print(f"  Type 2: {p2}")
            sys.exit(0)
        except Exception as e:
            print(f"Conversion failed: {e}", file=sys.stderr)
            sys.exit(1)

    # GUI mode
    if webview is None:
        print("pywebview is not installed. Please run: pip install pywebview", file=sys.stderr)
        sys.exit(1)

    api = Api()
    api.window = webview.create_window(
        title="Cross Reference Converter",
        html=HTML,
        js_api=api,
        width=780,
        height=580,
        resizable=True,
    )
    webview.start()


if __name__ == "__main__":
    main()
