import { useState } from "react";
import ExcelJS from "exceljs";

function buildSheetUrl(baseUrl, sheetName) {
  try {
    const parsed = new URL(baseUrl);
    const params = new URLSearchParams(parsed.search);
    if (params.has("id")) {
      const basePath = params.get("id").replace(/\/$/, "");
      const fullPath = basePath + "/" + sheetName;
      const encodedPath = encodeURIComponent(fullPath);
      let newSearch = "id=" + encodedPath;
      params.forEach((v, k) => {
        if (k !== "id") newSearch += "&" + encodeURIComponent(k) + "=" + encodeURIComponent(v);
      });
      return parsed.origin + parsed.pathname + "?" + newSearch;
    }
  } catch {}
  return baseUrl.replace(/\/$/, "") + "/" + sheetName;
}

async function processExcelClientSide(file, baseUrl) {
  const arrayBuffer = await file.arrayBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(arrayBuffer);

  const results = [];

  workbook.eachSheet((worksheet) => {
    const sheetName = worksheet.name;
    let linked = false;

    worksheet.eachRow((row, rowNumber) => {
      if (linked) return;
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        if (linked) return;
        const raw = cell.value;
        const text =
          raw === null || raw === undefined
            ? ""
            : typeof raw === "object" && raw.text
            ? raw.text
            : String(raw);

        if (text.trim().toLowerCase() === "test report path") {
          const targetRow = rowNumber;
          const targetCol = colNumber + 1;
          const folderUrl = buildSheetUrl(baseUrl, sheetName);

          // Unmerge the target cell if it is part of a merged region
          const targetCell = worksheet.getCell(targetRow, targetCol);
          if (targetCell.isMerged) {
            try {
              const masterAddr = targetCell.master.address;
              const mergeEntry = worksheet._merges && worksheet._merges[masterAddr];
              if (mergeEntry) {
                const m = mergeEntry.model;
                worksheet.unMergeCells(m.top, m.left, m.bottom, m.right);
              }
            } catch (_) {
              // ignore unmerge errors
            }
          }

          const tc = worksheet.getCell(targetRow, targetCol);
          tc.value = { text: sheetName, hyperlink: folderUrl };
          tc.font = { color: { argb: "FF0563C1" }, underline: true };

          results.push({ sheet: sheetName, url: folderUrl, status: "linked" });
          linked = true;
        }
      });
    });

    if (!linked) {
      results.push({ sheet: sheetName, url: null, status: "Test Report path label not found" });
    }
  });

  const buffer = await workbook.xlsx.writeBuffer();
  return { buffer, results };
}

export default function App() {
  const [file, setFile] = useState(null);
  const [baseUrl, setBaseUrl] = useState("");
  const [loading, setLoading] = useState(false);
  const [status, setStatus] = useState(null);
  const [results, setResults] = useState(null);
  const [dragOver, setDragOver] = useState(false);

  const handleFile = (f) => {
    if (!f) return;
    if (!f.name.endsWith(".xlsx")) {
      setStatus({ type: "error", msg: "Please upload an .xlsx file." });
      return;
    }
    setFile(f);
    setStatus(null);
    setResults(null);
  };

  const process = async () => {
    if (!file) {
      setStatus({ type: "error", msg: "Please upload your Excel file first." });
      return;
    }
    if (!baseUrl.trim()) {
      setStatus({ type: "error", msg: "Please enter the base SharePoint folder URL." });
      return;
    }

    setLoading(true);
    setResults(null);
    setStatus({ type: "info", msg: "Processing your Excel file…" });

    try {
      const { buffer, results: res } = await processExcelClientSide(file, baseUrl.trim());

      const blob = new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = file.name.replace(".xlsx", "_linked.xlsx");
      a.click();
      URL.revokeObjectURL(url);

      setResults(res);
      setStatus({
        type: "success",
        msg: "Done! Your updated Excel file has been downloaded. Upload it back to SharePoint to replace the original.",
      });
    } catch (e) {
      setStatus({ type: "error", msg: e.message });
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="container">
      <header>
        <h1>SharePoint Test Report Linker</h1>
        <p className="subtitle">
          Inserts SharePoint folder hyperlinks into <strong>Test Report path</strong> cells across all sheets.
        </p>
      </header>

      {/* Step 1 – Upload */}
      <section className="card">
        <div className="step-label">Step 1 — Upload Excel File</div>
        <p className="hint">
          Download your Excel file from SharePoint first, then upload it here.
        </p>
        <div
          className={`drop-zone ${dragOver ? "drag-over" : ""} ${file ? "has-file" : ""}`}
          onClick={() => document.getElementById("fileInput").click()}
          onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
          onDragLeave={() => setDragOver(false)}
          onDrop={(e) => { e.preventDefault(); setDragOver(false); handleFile(e.dataTransfer.files[0]); }}
        >
          <input
            id="fileInput"
            type="file"
            accept=".xlsx"
            style={{ display: "none" }}
            onChange={(e) => handleFile(e.target.files[0])}
          />
          {file ? (
            <div className="file-info">
              <span className="file-icon">📄</span>
              <span className="file-name">{file.name}</span>
              <span className="file-size">{(file.size / 1024).toFixed(1)} KB</span>
            </div>
          ) : (
            <div className="drop-hint">
              <span className="drop-icon">⬆</span>
              <span>Click to upload or drag & drop</span>
              <span className="drop-sub">.xlsx files only</span>
            </div>
          )}
        </div>
        {file && (
          <button className="btn btn-ghost" style={{ marginTop: 8, fontSize: 13 }}
            onClick={() => { setFile(null); setResults(null); setStatus(null); }}>
            ✕ Remove file
          </button>
        )}
      </section>

      {/* Step 2 – Base URL */}
      <section className="card">
        <div className="step-label">Step 2 — SharePoint Base Folder URL</div>
        <p className="hint">
          In SharePoint, open the folder that contains your test result subfolders
          (e.g. <code>IMDS_001</code>, <code>IMDS_002</code>).
          Copy the URL from the <strong>browser address bar</strong> and paste it here.
        </p>
        <input
          className="input"
          value={baseUrl}
          onChange={(e) => setBaseUrl(e.target.value)}
          placeholder="https://polarisgrid-my.sharepoint.com/personal/rnd_.../Documents/Test Results"
        />
        <p className="hint" style={{ marginTop: 6 }}>
          Each sheet name will be appended: <code>…/Test Results/IMDS_001</code>
        </p>
      </section>

      {/* Step 3 – Process */}
      <section className="card">
        <div className="step-label">Step 3 — Process & Download</div>
        <button
          className={`btn btn-success ${!file || loading ? "disabled" : ""}`}
          onClick={process}
          disabled={!file || loading}
        >
          {loading ? <><Spinner /> Processing…</> : "⚡ Insert Hyperlinks & Download"}
        </button>
      </section>

      {/* Step 4 hint */}
      <section className="card" style={{ background: "#f8fafc" }}>
        <div className="step-label">Step 4 — Upload back to SharePoint</div>
        <p className="hint" style={{ margin: 0 }}>
          After downloading the updated file, go to SharePoint → navigate to where the original file is →
          upload the new file to replace it.
        </p>
      </section>

      {/* Status */}
      {status && (
        <div className={`alert alert-${status.type}`}>{status.msg}</div>
      )}

      {/* Results */}
      {results && (
        <section className="card">
          <h2 className="results-title">Results</h2>
          <table className="results-table">
            <thead>
              <tr>
                <th>Sheet</th>
                <th>Status</th>
                <th>Folder URL</th>
              </tr>
            </thead>
            <tbody>
              {results.map((r, i) => (
                <tr key={i}>
                  <td><code>{r.sheet}</code></td>
                  <td>
                    {r.url
                      ? <span className="badge-success">✓ Linked</span>
                      : <span className="badge-error">✗ {r.status}</span>}
                  </td>
                  <td>
                    {r.url
                      ? <a href={r.url} target="_blank" rel="noopener noreferrer">{r.url}</a>
                      : <span className="muted">—</span>}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </section>
      )}
    </div>
  );
}

function Spinner() {
  return <span className="spinner" />;
}
