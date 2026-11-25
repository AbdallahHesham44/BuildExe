import os
import traceback

import pandas as pd
import webview


# ------------------------------------------------------------
# CORE LOGIC (same as your script, but returns stats instead of
# print() - a --noconsole exe has no stdout, so print() can crash)
# ------------------------------------------------------------
def filter_and_add_qa_comments(first_file, second_file, output_file):
    df_first = pd.read_excel(first_file, dtype=object)
    df_second = pd.read_excel(second_file, dtype=object)

    # Remove rows from second file where String ID is blank
    df_second = df_second[
        df_second["String ID"].notna()
        & df_second["String ID"].astype(str).str.strip().ne("")
    ].copy()
    df_second.reset_index(drop=True, inplace=True)

    def normalize(value):
        if pd.isna(value):
            return ""
        return str(value).strip()

    def unique_sources(series):
        return list(dict.fromkeys(normalize(v) for v in series if normalize(v)))

    # STEP 1: ORIGINAL PART ID CHECK
    df_first["_part_id"] = df_first["PartId"].apply(normalize)

    z2_df = df_first[df_first["SourceName"].apply(normalize) == "Z2 Cross"]
    z2_part_lookup = set(z2_df["_part_id"])

    all_part_lookup = (
        df_first.groupby("_part_id")["SourceName"].agg(unique_sources).to_dict()
    )

    def check_original_part(part_id):
        part_id = normalize(part_id)
        if not part_id:
            return "Not Found Original PN"
        if part_id in z2_part_lookup:
            return "Found at Z2 Cross"
        sources = all_part_lookup.get(part_id, [])
        if sources:
            return f"Found at {', '.join(sources)}"
        return "Not Found Original PN"

    df_second["QA Comment"] = df_second["Original PartID"].apply(check_original_part)

    # STEP 2: FAMILY KEYS
    df_first["_family_key"] = list(zip(
        df_first["PartId"].apply(normalize),
        df_first["CrossFamilyName"].apply(normalize),
    ))
    df_second["_family_key"] = list(zip(
        df_second["Original PartID"].apply(normalize),
        df_second["CrossFamily Name"].apply(normalize),
    ))

    # STEP 3: FAMILY LOOKUP
    z2_family_df = df_first[df_first["SourceName"].apply(normalize) == "Z2 Cross"]
    z2_family_lookup = set(z2_family_df["_family_key"])

    all_family_lookup = (
        df_first.groupby("_family_key")["SourceName"].agg(unique_sources).to_dict()
    )

    # STEP 4: FAMILY QA CHECK
    def check_family(row):
        if row["QA Comment"] != "Found at Z2 Cross":
            return row["QA Comment"]
        key = row["_family_key"]
        if key in z2_family_lookup:
            return "Found Family at Z2 Cross"
        sources = all_family_lookup.get(key, [])
        if sources:
            return f"Found Family at {', '.join(sources)}"
        return "Not Found Family"

    df_second["QA Comment"] = df_second.apply(check_family, axis=1)

    # STEP 5: CLEAN TEMP COLUMNS
    df_first.drop(columns=["_part_id", "_family_key"], inplace=True)
    df_second.drop(columns=["_family_key"], inplace=True)

    # STEP 6: SAVE
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df_first.to_excel(writer, sheet_name="First File", index=False)
        df_second.to_excel(writer, sheet_name="Second File With QA", index=False)

    # STEP 7: STATS
    return {
        "first_rows": int(len(df_first)),
        "second_rows": int(len(df_second)),
        "stats": {str(k): int(v) for k, v in df_second["QA Comment"].value_counts().items()},
        "output": output_file,
    }


# ------------------------------------------------------------
# JS <-> PYTHON BRIDGE
# ------------------------------------------------------------
def _dialog(kind):
    """Support both new (FileDialog enum) and old pywebview versions."""
    fd = getattr(webview, "FileDialog", None)
    if fd is not None:
        return getattr(fd, kind)
    return getattr(webview, kind + "_DIALOG")


class Api:
    def pick_file(self):
        result = webview.windows[0].create_file_dialog(
            _dialog("OPEN"),
            allow_multiple=False,
            file_types=("Excel Files (*.xlsx;*.xls)", "All files (*.*)"),
        )
        return result[0] if result else None

    def pick_output(self):
        result = webview.windows[0].create_file_dialog(
            _dialog("SAVE"),
            save_filename="Final_Result.xlsx",
            file_types=("Excel Files (*.xlsx)",),
        )
        if not result:
            return None
        path = result if isinstance(result, str) else result[0]
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"
        return path

    def run(self, first_file, second_file, output_file):
        try:
            if not (first_file and second_file and output_file):
                return {"ok": False, "error": "Please select all three paths."}
            data = filter_and_add_qa_comments(first_file, second_file, output_file)
            return {"ok": True, **data}
        except KeyError as e:
            return {"ok": False, "error": f"Missing expected column: {e}"}
        except PermissionError:
            return {"ok": False, "error": "Cannot write output. Is it open in Excel?"}
        except Exception as e:
            return {"ok": False, "error": f"{e}\n\n{traceback.format_exc()}"}

    def open_output_folder(self, path):
        if path and os.path.exists(path):
            os.startfile(os.path.dirname(path))  # Windows


# ------------------------------------------------------------
# UI
# ------------------------------------------------------------
HTML = """
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  body { font-family: Segoe UI, Arial, sans-serif; background:#f4f6f9; margin:0; padding:24px; color:#222; }
  h2 { margin-top:0; }
  .card { background:#fff; border-radius:10px; padding:20px; box-shadow:0 1px 4px rgba(0,0,0,.1); max-width:720px; margin:auto; }
  .row { display:flex; gap:8px; margin-bottom:12px; align-items:center; }
  .row label { width:130px; font-weight:600; font-size:14px; }
  .row input { flex:1; padding:8px; border:1px solid #ccc; border-radius:6px; font-size:13px; background:#fafafa; }
  button { padding:8px 14px; border:0; border-radius:6px; cursor:pointer; background:#2563eb; color:#fff; font-size:13px; }
  button.secondary { background:#6b7280; }
  button:disabled { opacity:.5; cursor:default; }
  #run { width:100%; padding:12px; font-size:15px; margin-top:8px; }
  #result { margin-top:18px; font-size:14px; white-space:pre-wrap; }
  table { border-collapse:collapse; width:100%; margin-top:8px; }
  td, th { border:1px solid #ddd; padding:6px 10px; text-align:left; font-size:13px; }
  th { background:#f0f3f8; }
  .err { color:#b91c1c; }
  .ok { color:#15803d; font-weight:600; }
</style>
</head>
<body>
<div class="card">
  <h2>QA Comments Tool</h2>
  <div class="row"><label>First file</label><input id="first" readonly><button onclick="pick('first')">Browse</button></div>
  <div class="row"><label>Second file</label><input id="second" readonly><button onclick="pick('second')">Browse</button></div>
  <div class="row"><label>Output file</label><input id="out" readonly><button onclick="pickOut()">Save as</button></div>
  <button id="run" onclick="runQA()">Run</button>
  <div id="result"></div>
</div>

<script>
async function pick(id) {
  const p = await pywebview.api.pick_file();
  if (p) document.getElementById(id).value = p;
}
async function pickOut() {
  const p = await pywebview.api.pick_output();
  if (p) document.getElementById('out').value = p;
}
async function runQA() {
  const btn = document.getElementById('run');
  const res = document.getElementById('result');
  btn.disabled = true; btn.textContent = 'Processing...';
  res.innerHTML = '';
  const r = await pywebview.api.run(
    document.getElementById('first').value,
    document.getElementById('second').value,
    document.getElementById('out').value
  );
  btn.disabled = false; btn.textContent = 'Run';
  if (!r.ok) { res.innerHTML = '<span class="err"></span>'; res.firstChild.textContent = r.error; return; }
  let rows = Object.entries(r.stats).map(([k,v]) => `<tr><td></td><td>${v}</td></tr>`);
  res.innerHTML = `<div class="ok">Processing completed successfully!</div>
    <div>First file rows: ${r.first_rows}<br>Second file rows: ${r.second_rows}</div>
    <table><tr><th>QA Comment</th><th>Count</th></tr>${rows.join('')}</table>
    <div style="margin-top:10px"><button class="secondary" onclick="pywebview.api.open_output_folder(document.getElementById('out').value)">Open folder</button></div>`;
  // set labels safely as text
  const cells = res.querySelectorAll('table tr:not(:first-child) td:first-child');
  Object.keys(r.stats).forEach((k, i) => cells[i].textContent = k);
}
</script>
</body>
</html>
"""

if __name__ == "__main__":
    webview.create_window(
        "QA Comments Tool", html=HTML, js_api=Api(), width=800, height=640
    )
    webview.start()
