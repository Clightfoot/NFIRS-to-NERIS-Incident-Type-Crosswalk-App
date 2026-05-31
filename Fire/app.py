from __future__ import annotations

import cgi
import html
import io
import json
import mimetypes
import re
import subprocess
import sys
import traceback
import uuid
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from tempfile import TemporaryDirectory
from urllib.parse import unquote, urlparse

import pandas as pd


BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
OUTPUT_DIR = BASE_DIR / "outputs"
UPLOAD_DIR = BASE_DIR / "uploads"
MAPPING_CSV = DATA_DIR / "incident_type_crosswalk.csv"

REFERENCE_WORKBOOKS = [
    BASE_DIR / "Copy of Oregon State Fire Marshal_NFIRS to NERIS Crosswalk_Dec 2025_3.0.xlsx",
    Path(r"C:\Users\jcbro\OneDrive\Documents\NFA\R0387 ATDM\NFRIS to NERIS Conversion\Oregon State Fire Marshal_NFIRS to NERIS Crosswalk_Dec 2025.xlsx"),
    Path(r"C:\Users\jcbro\OneDrive\Documents\NFA\R0387 ATDM\NFRIS to NERIS Conversion\NFIRS to NERIS Coding Crosswalk Using Slicers 5.0.xlsx"),
]

CODE_COLUMN_HINTS = [
    "incident type 1",
    "in type 1",
    "in_type_1",
    "type 1",
    "incident type",
    "nfirs incident type",
    "nfirs incident type code",
    "incident_type",
    "incidenttype",
    "inc type",
    "type code",
    "nfirs code",
]


def clean_label(value: object) -> str:
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).replace("\n", " ")).strip()


def normalize_code(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if not text:
        return ""
    if re.fullmatch(r"\d+\.0+", text):
        text = text.split(".", 1)[0]
    match = re.search(r"\d{3}", text)
    if match:
        return match.group(0)
    if text.isdigit():
        if len(text) == 1:
            return text
        return text.zfill(3)[-3:]
    return ""


def normalized_header(value: object) -> str:
    return re.sub(r"[^a-z0-9]+", " ", str(value).lower()).strip()


def ensure_directories() -> None:
    DATA_DIR.mkdir(exist_ok=True)
    OUTPUT_DIR.mkdir(exist_ok=True)
    UPLOAD_DIR.mkdir(exist_ok=True)


def build_mapping_from_workbook() -> pd.DataFrame:
    for workbook_path in REFERENCE_WORKBOOKS:
        if not workbook_path.exists():
            continue

        excel = pd.ExcelFile(workbook_path)
        sheet_name = None
        for candidate in ("Incident Type", "NFIRS Incident Types"):
            if candidate in excel.sheet_names:
                sheet_name = candidate
                break
        if not sheet_name:
            continue

        df = pd.read_excel(workbook_path, sheet_name=sheet_name, dtype=object)
        df.columns = [clean_label(col) for col in df.columns]
        required = [
            "NFIRS Code",
            "NFIRS Incident Type Code Description",
            "NERIS Category",
            "NERIS Subcategory",
            "NERIS Description",
            "NERIS Definition",
        ]
        missing = [col for col in required if col not in df.columns]
        if missing:
            continue

        df = df[required].copy()
        df["NFIRS Code"] = df["NFIRS Code"].map(normalize_code)
        for col in required[1:]:
            df[col] = df[col].map(clean_label)
        df = df[df["NFIRS Code"] != ""].drop_duplicates().reset_index(drop=True)
        df.to_csv(MAPPING_CSV, index=False)
        return df

    raise RuntimeError(
        "No incident type crosswalk workbook was found. Place the Oregon or NFIRS crosswalk workbook next to app.py."
    )


def load_mapping() -> pd.DataFrame:
    ensure_directories()
    if MAPPING_CSV.exists():
        return pd.read_csv(MAPPING_CSV, dtype=str).fillna("")
    return build_mapping_from_workbook()


def make_lookup(mapping: pd.DataFrame) -> dict[str, dict[str, str]]:
    lookup: dict[str, dict[str, str]] = {}
    grouped_frames: dict[str, pd.DataFrame] = {
        str(code): group for code, group in mapping.groupby("NFIRS Code", sort=True)
    }
    for digit, group in mapping.assign(_group=mapping["NFIRS Code"].str[0]).groupby("_group", sort=True):
        grouped_frames[str(digit)] = group

    for code, group in grouped_frames.items():
        lookup[str(code)] = {
            "nfirs_description": "; ".join(sorted(set(group["NFIRS Incident Type Code Description"]) - {""})),
            "neris_category": "; ".join(sorted(set(group["NERIS Category"]) - {""})),
            "neris_subcategory": "; ".join(sorted(set(group["NERIS Subcategory"]) - {""})),
            "neris_description": "; ".join(sorted(set(group["NERIS Description"]) - {""})),
            "neris_definition": "; ".join(sorted(set(group["NERIS Definition"]) - {""})),
            "match_count": str(len(group)),
        }
    return lookup


def pick_code_column(df: pd.DataFrame, requested_column: str | None = None) -> str:
    columns = [str(col) for col in df.columns]
    if requested_column and requested_column in columns:
        return requested_column

    normalized = {col: normalized_header(col) for col in columns}
    for hint in CODE_COLUMN_HINTS:
        hint_norm = normalized_header(hint)
        for col, col_norm in normalized.items():
            if col_norm == hint_norm or hint_norm in col_norm:
                return col

    scored: list[tuple[int, str]] = []
    for col in columns:
        sample = df[col].dropna().head(200).map(normalize_code)
        valid_count = int((sample != "").sum())
        if valid_count:
            scored.append((valid_count, col))
    if scored:
        return sorted(scored, reverse=True)[0][1]

    raise ValueError("I could not find a column that looks like an NFIRS incident type code.")


def read_upload(filename: str, content: bytes) -> pd.DataFrame:
    suffix = Path(filename).suffix.lower()
    if suffix == ".csv":
        return pd.read_csv(io.BytesIO(content), dtype=object)
    if suffix in {".xlsx", ".xlsm"}:
        return read_excel_best_sheet(io.BytesIO(content))
    if suffix == ".xls":
        try:
            return read_excel_best_sheet(io.BytesIO(content))
        except Exception as exc:
            html_table = read_html_disguised_as_xls(content)
            if html_table is not None:
                return html_table
            return read_xls_via_excel_conversion(filename, content, exc)
    raise ValueError("Upload a .csv, .xls, or .xlsx file.")


def read_excel_best_sheet(source: object) -> pd.DataFrame:
    sheets = pd.read_excel(source, sheet_name=None, dtype=object)
    best_name = ""
    best_score = -1
    best_df = pd.DataFrame()
    for sheet_name, sheet_df in sheets.items():
        candidate = sheet_df.dropna(how="all").dropna(axis=1, how="all")
        if candidate.empty:
            continue
        headers = [normalized_header(col) for col in candidate.columns]
        header_match = any(
            normalized_header(hint) in header or header in normalized_header(hint)
            for header in headers
            for hint in CODE_COLUMN_HINTS
            if header
        )
        score = (len(candidate) * max(len(candidate.columns), 1)) + (100000 if header_match else 0)
        if score > best_score:
            best_name = str(sheet_name)
            best_score = score
            best_df = candidate

    if best_df.empty:
        return best_df
    best_df.attrs["source_sheet"] = best_name
    return best_df


def read_html_disguised_as_xls(content: bytes) -> pd.DataFrame | None:
    prefix = content[:500].lower()
    if b"<html" not in prefix and b"<table" not in prefix:
        return None
    tables = pd.read_html(io.BytesIO(content))
    if not tables:
        return None
    tables = [table.dropna(how="all").dropna(axis=1, how="all") for table in tables]
    tables = [table for table in tables if not table.empty]
    if not tables:
        return None
    return sorted(tables, key=lambda table: len(table) * max(len(table.columns), 1), reverse=True)[0]


def read_xls_via_excel_conversion(filename: str, content: bytes, original_error: Exception) -> pd.DataFrame:
    with TemporaryDirectory(dir=UPLOAD_DIR) as temp_dir:
        temp_path = Path(temp_dir)
        source = temp_path / (re.sub(r"[^A-Za-z0-9_.-]+", "_", Path(filename).name) or "upload.xls")
        converted = temp_path / "converted.xlsx"
        source.write_bytes(content)

        ps_script = f"""
$ErrorActionPreference = 'Stop'
$excel = $null
$workbook = $null
try {{
  $excel = New-Object -ComObject Excel.Application
  $excel.DisplayAlerts = $false
  $workbook = $excel.Workbooks.Open('{str(source).replace("'", "''")}')
  $workbook.SaveAs('{str(converted).replace("'", "''")}', 51)
}} finally {{
  if ($workbook -ne $null) {{ $workbook.Close($false) | Out-Null }}
  if ($excel -ne $null) {{
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
  }}
}}
"""
        result = subprocess.run(
            ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_script],
            capture_output=True,
            text=True,
            timeout=120,
        )
        if result.returncode != 0 or not converted.exists():
            detail = clean_label(result.stderr or result.stdout or str(original_error))
            raise ValueError(
                "This older .xls file could not be read directly, and automatic Excel conversion was not available. "
                "Open the file in Excel, save it as .xlsx or .csv, and upload that copy. "
                f"Details: {detail[:240]}"
            ) from original_error
        return read_excel_best_sheet(converted)


def add_crosswalk_columns(df: pd.DataFrame, code_column: str, lookup: dict[str, dict[str, str]]) -> pd.DataFrame:
    output = df.copy()
    output["NFIRS Incident Type Code Clean"] = output[code_column].map(normalize_code)

    def lookup_field(code: str, field: str) -> str:
        return lookup.get(code, {}).get(field, "")

    output["NERIS Incident Type Category"] = output["NFIRS Incident Type Code Clean"].map(
        lambda code: lookup_field(code, "neris_category")
    )
    output["NERIS Incident Type Subcategory"] = output["NFIRS Incident Type Code Clean"].map(
        lambda code: lookup_field(code, "neris_subcategory")
    )
    output["NERIS Incident Type Description"] = output["NFIRS Incident Type Code Clean"].map(
        lambda code: lookup_field(code, "neris_description")
    )
    output["NERIS Definition"] = output["NFIRS Incident Type Code Clean"].map(
        lambda code: lookup_field(code, "neris_definition")
    )
    output["Crosswalk Match Count"] = output["NFIRS Incident Type Code Clean"].map(
        lambda code: lookup_field(code, "match_count")
    )
    output["Crosswalk Status"] = output["NFIRS Incident Type Code Clean"].map(
        lambda code: "Blank code" if not code else "Matched" if code in lookup else "No crosswalk match"
    )
    return output


def summary_tables(enriched: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    total = len(enriched)
    overview = pd.DataFrame(
        [
            ("Rows analyzed", total),
            ("Matched rows", int((enriched["Crosswalk Status"] == "Matched").sum())),
            ("Unmatched rows", int((enriched["Crosswalk Status"] == "No crosswalk match").sum())),
            ("Blank incident type rows", int((enriched["Crosswalk Status"] == "Blank code").sum())),
            ("Unique NFIRS incident type codes", int(enriched["NFIRS Incident Type Code Clean"].replace("", pd.NA).nunique())),
        ],
        columns=["Metric", "Value"],
    )

    category = (
        enriched.groupby(["NERIS Incident Type Category", "Crosswalk Status"], dropna=False)
        .size()
        .reset_index(name="Row Count")
        .sort_values("Row Count", ascending=False)
    )
    category["NERIS Incident Type Category"] = category["NERIS Incident Type Category"].replace("", "Unmapped")
    category["Percent of Rows"] = category["Row Count"].map(lambda count: round((count / total * 100), 2) if total else 0)

    by_code = (
        enriched.groupby(
            [
                "NFIRS Incident Type Code Clean",
                "NERIS Incident Type Category",
                "NERIS Incident Type Subcategory",
                "Crosswalk Status",
            ],
            dropna=False,
        )
        .size()
        .reset_index(name="Row Count")
        .sort_values(["Row Count", "NFIRS Incident Type Code Clean"], ascending=[False, True])
    )
    return overview, category, by_code


def write_outputs(enriched: pd.DataFrame, original_filename: str) -> dict[str, str]:
    run_id = datetime.now().strftime("%Y%m%d_%H%M%S") + "_" + uuid.uuid4().hex[:8]
    run_dir = OUTPUT_DIR / run_id
    run_dir.mkdir(parents=True, exist_ok=True)

    stem = re.sub(r"[^A-Za-z0-9_-]+", "_", Path(original_filename).stem).strip("_") or "crosswalk"
    csv_path = run_dir / f"{stem}_neris_crosswalk_rows.csv"
    xlsx_path = run_dir / f"{stem}_neris_crosswalk_summary.xlsx"

    overview, category, by_code = summary_tables(enriched)
    enriched.to_csv(csv_path, index=False)

    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        overview.to_excel(writer, index=False, sheet_name="Summary")
        category.to_excel(writer, index=False, sheet_name="Category Summary")
        by_code.to_excel(writer, index=False, sheet_name="Code Summary")
        enriched.to_excel(writer, index=False, sheet_name="Crosswalk Rows")
        for sheet in writer.book.worksheets:
            sheet.freeze_panes = "A2"
            for column_cells in sheet.columns:
                max_length = max(len(str(cell.value or "")) for cell in column_cells[:200])
                sheet.column_dimensions[column_cells[0].column_letter].width = min(max(max_length + 2, 12), 55)

    return {
        "csv": f"/download/{run_id}/{csv_path.name}",
        "xlsx": f"/download/{run_id}/{xlsx_path.name}",
    }


def analyze_file(filename: str, content: bytes, requested_column: str | None = None) -> dict[str, object]:
    mapping = load_mapping()
    lookup = make_lookup(mapping)
    df = read_upload(filename, content)
    if df.empty:
        raise ValueError("The uploaded file does not contain any rows.")
    source_sheet = df.attrs.get("source_sheet")
    code_column = pick_code_column(df, requested_column)
    enriched = add_crosswalk_columns(df, code_column, lookup)
    overview, category, by_code = summary_tables(enriched)
    downloads = write_outputs(enriched, filename)
    return {
        "fileName": filename,
        "codeColumn": code_column,
        "sourceSheet": source_sheet,
        "downloads": downloads,
        "overview": overview.to_dict(orient="records"),
        "category": category.head(25).to_dict(orient="records"),
        "byCode": by_code.head(50).to_dict(orient="records"),
        "columns": [str(col) for col in df.columns],
    }


INDEX_HTML = """<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>NFIRS to NERIS Incident Type Crosswalk</title>
  <style>
    :root {
      color-scheme: light;
      --ink: #18212f;
      --muted: #5f6b7a;
      --line: #d7dde5;
      --panel: #ffffff;
      --page: #f4f7f9;
      --accent: #0b6f6b;
      --accent-dark: #084f4c;
      --warn: #9a4f00;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      background: var(--page);
      color: var(--ink);
      font-family: "Segoe UI", system-ui, -apple-system, sans-serif;
    }
    header {
      background: #ffffff;
      border-bottom: 1px solid var(--line);
      padding: 22px 28px;
    }
    h1 { margin: 0; font-size: 24px; font-weight: 700; letter-spacing: 0; }
    main {
      width: min(1180px, calc(100vw - 32px));
      margin: 24px auto 48px;
      display: grid;
      gap: 18px;
    }
    section {
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 18px;
    }
    .upload-grid {
      display: grid;
      grid-template-columns: minmax(260px, 1fr) minmax(220px, 320px) auto;
      gap: 12px;
      align-items: end;
    }
    label { display: grid; gap: 6px; color: var(--muted); font-size: 13px; font-weight: 600; }
    input, select {
      width: 100%;
      border: 1px solid var(--line);
      border-radius: 6px;
      padding: 10px 11px;
      background: #fff;
      color: var(--ink);
      font: inherit;
    }
    button, .download {
      border: 0;
      border-radius: 6px;
      background: var(--accent);
      color: white;
      padding: 11px 15px;
      font-weight: 700;
      cursor: pointer;
      text-decoration: none;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      min-height: 42px;
    }
    button:hover, .download:hover { background: var(--accent-dark); }
    button:disabled { background: #8aa3a1; cursor: wait; }
    .status { color: var(--muted); min-height: 22px; }
    .error { color: #8a1f17; font-weight: 650; }
    .metrics {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
      gap: 12px;
    }
    .metric {
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 14px;
      background: #fbfcfd;
    }
    .metric strong { display: block; font-size: 26px; margin-bottom: 4px; }
    .metric span { color: var(--muted); font-size: 13px; }
    .downloads { display: flex; flex-wrap: wrap; gap: 10px; }
    .table-wrap { overflow: auto; border: 1px solid var(--line); border-radius: 8px; }
    table { border-collapse: collapse; width: 100%; min-width: 720px; background: white; }
    th, td { border-bottom: 1px solid var(--line); padding: 9px 10px; text-align: left; vertical-align: top; }
    th { background: #eef3f6; font-size: 13px; position: sticky; top: 0; }
    td { font-size: 13px; }
    .hidden { display: none; }
    .note { color: var(--warn); font-size: 13px; margin-top: 8px; }
    @media (max-width: 780px) {
      header { padding: 18px 16px; }
      .upload-grid { grid-template-columns: 1fr; }
      section { padding: 14px; }
    }
  </style>
</head>
<body>
  <header>
    <h1>NFIRS to NERIS Incident Type Crosswalk</h1>
  </header>
  <main>
    <section>
      <form id="uploadForm" class="upload-grid">
        <label>
          Data file
          <input id="fileInput" name="file" type="file" accept=".csv,.xls,.xlsx,.xlsm" required>
        </label>
        <label>
          Incident type column
          <select id="columnSelect" name="code_column">
            <option value="">Auto-detect</option>
          </select>
        </label>
        <button id="runButton" type="submit">Analyze</button>
      </form>
      <div id="status" class="status"></div>
      <div class="note">For older binary .xls files, saving a copy as .xlsx or .csv gives the most reliable upload.</div>
    </section>

    <section id="results" class="hidden">
      <h2>Summary</h2>
      <p id="columnUsed" class="status"></p>
      <div id="metrics" class="metrics"></div>
      <h2>Downloads</h2>
      <div id="downloads" class="downloads"></div>
    </section>

    <section id="categorySection" class="hidden">
      <h2>NERIS Category Summary</h2>
      <div class="table-wrap"><table id="categoryTable"></table></div>
    </section>

    <section id="codeSection" class="hidden">
      <h2>NFIRS Code Summary</h2>
      <div class="table-wrap"><table id="codeTable"></table></div>
    </section>
  </main>

  <script>
    const form = document.getElementById("uploadForm");
    const fileInput = document.getElementById("fileInput");
    const runButton = document.getElementById("runButton");
    const statusBox = document.getElementById("status");
    const results = document.getElementById("results");
    const categorySection = document.getElementById("categorySection");
    const codeSection = document.getElementById("codeSection");

    function escapeHtml(value) {
      return String(value ?? "").replace(/[&<>"']/g, c => ({
        "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#039;"
      }[c]));
    }

    function renderTable(table, rows) {
      if (!rows.length) {
        table.innerHTML = "<tbody><tr><td>No rows to show.</td></tr></tbody>";
        return;
      }
      const headers = Object.keys(rows[0]);
      table.innerHTML = "<thead><tr>" + headers.map(h => `<th>${escapeHtml(h)}</th>`).join("") +
        "</tr></thead><tbody>" +
        rows.map(row => "<tr>" + headers.map(h => `<td>${escapeHtml(row[h])}</td>`).join("") + "</tr>").join("") +
        "</tbody>";
    }

    form.addEventListener("submit", async event => {
      event.preventDefault();
      const file = fileInput.files[0];
      if (!file) return;

      runButton.disabled = true;
      statusBox.className = "status";
      statusBox.textContent = "Analyzing uploaded file...";
      results.classList.add("hidden");
      categorySection.classList.add("hidden");
      codeSection.classList.add("hidden");

      const body = new FormData(form);
      try {
        const response = await fetch("/analyze", { method: "POST", body });
        const payload = await response.json();
        if (!response.ok) throw new Error(payload.error || "Analysis failed.");

        document.getElementById("columnUsed").textContent =
          `Incident type column used: ${payload.codeColumn}` +
          (payload.sourceSheet ? ` | Excel sheet used: ${payload.sourceSheet}` : "");
        document.getElementById("metrics").innerHTML = payload.overview.map(row =>
          `<div class="metric"><strong>${escapeHtml(row.Value)}</strong><span>${escapeHtml(row.Metric)}</span></div>`
        ).join("");
        document.getElementById("downloads").innerHTML = `
          <a class="download" href="${payload.downloads.xlsx}">Download summary workbook</a>
          <a class="download" href="${payload.downloads.csv}">Download crosswalk CSV</a>
        `;
        renderTable(document.getElementById("categoryTable"), payload.category);
        renderTable(document.getElementById("codeTable"), payload.byCode);

        results.classList.remove("hidden");
        categorySection.classList.remove("hidden");
        codeSection.classList.remove("hidden");
        statusBox.textContent = `Finished ${payload.fileName}.`;
      } catch (error) {
        statusBox.className = "status error";
        statusBox.textContent = error.message;
      } finally {
        runButton.disabled = false;
      }
    });
  </script>
</body>
</html>
"""


class CrosswalkHandler(BaseHTTPRequestHandler):
    def send_json(self, status: int, payload: dict[str, object]) -> None:
        body = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/":
            body = INDEX_HTML.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return

        if parsed.path.startswith("/download/"):
            relative = unquote(parsed.path.removeprefix("/download/"))
            file_path = (OUTPUT_DIR / relative).resolve()
            if not str(file_path).startswith(str(OUTPUT_DIR.resolve())) or not file_path.exists():
                self.send_error(404)
                return
            content_type = mimetypes.guess_type(file_path.name)[0] or "application/octet-stream"
            content = file_path.read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Disposition", f'attachment; filename="{file_path.name}"')
            self.send_header("Content-Length", str(len(content)))
            self.end_headers()
            self.wfile.write(content)
            return

        self.send_error(404)

    def do_POST(self) -> None:
        if urlparse(self.path).path != "/analyze":
            self.send_error(404)
            return
        try:
            form = cgi.FieldStorage(
                fp=self.rfile,
                headers=self.headers,
                environ={
                    "REQUEST_METHOD": "POST",
                    "CONTENT_TYPE": self.headers.get("Content-Type", ""),
                    "CONTENT_LENGTH": self.headers.get("Content-Length", "0"),
                },
            )
            upload = form["file"] if "file" in form else None
            if upload is None or not upload.filename:
                raise ValueError("Choose a data file to analyze.")
            requested_column = form.getfirst("code_column") or None
            content = upload.file.read()
            result = analyze_file(upload.filename, content, requested_column)
            self.send_json(200, result)
        except Exception as exc:
            traceback.print_exc()
            self.send_json(400, {"error": str(exc)})

    def log_message(self, format: str, *args: object) -> None:
        sys.stdout.write("%s - %s\n" % (self.address_string(), format % args))


def main() -> None:
    ensure_directories()
    load_mapping()
    port = 8765
    server = ThreadingHTTPServer(("127.0.0.1", port), CrosswalkHandler)
    print(f"NFIRS to NERIS Crosswalk app is running at http://127.0.0.1:{port}")
    print("Press Ctrl+C to stop.")
    server.serve_forever()


if __name__ == "__main__":
    main()
