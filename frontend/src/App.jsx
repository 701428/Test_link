import { useState } from "react";

const BACKEND_URL = import.meta.env.VITE_BACKEND_URL || "/api";

export default function App() {
  const [file, setFile] = useState(null);
  const [baseUrl, setBaseUrl] = useState("");
  const [loading, setLoading] = useState(false);
  const [status, setStatus] = useState(null);
  const [results, setResults] = useState(null);
  const [dragOver, setDragOver] = useState(false);

  const MAX_FILE_SIZE = 4 * 1024 * 1024; // 4 MB (Vercel serverless limit is 4.5 MB)

  const handleFile = (f) => {
    if (!f) return;
    if (!f.name.endsWith(".xlsx")) {
      setStatus({ type: "error", msg: "Please upload an .xlsx file." });
      return;
    }
    if (f.size > MAX_FILE_SIZE) {
      setStatus({
        type: "error",
        msg: `File too large (${(f.size / 1024 / 1024).toFixed(1)} MB). Maximum allowed size is 4 MB due to server limits.`,
      });
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
      const formData = new FormData();
      formData.append("file", file);
      formData.append("base_url", baseUrl.trim());

      let resp;
      try {
        resp = await fetch(`${BACKEND_URL}/process`, {
          method: "POST",
          body: formData,
        });
      } catch (networkErr) {
        throw new Error(`Cannot reach backend at ${BACKEND_URL}. Make sure the backend server is running. (${networkErr.message})`);
      }

      if (!resp.ok) {
        if (resp.status === 413) {
          throw new Error(
            "File too large (HTTP 413). The server rejected the upload because the file exceeds the 4.5 MB limit. Please reduce the file size and try again."
          );
        }
        let detail = `Server error ${resp.status}`;
        try { const err = await resp.json(); detail = err.detail || detail; } catch {}
        throw new Error(detail);
      }

      const data = await resp.json();

      // Trigger download from base64
      const byteChars = atob(data.file);
      const byteArr = new Uint8Array(byteChars.length);
      for (let i = 0; i < byteChars.length; i++) byteArr[i] = byteChars.charCodeAt(i);
      const blob = new Blob([byteArr], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = data.filename;
      a.click();
      URL.revokeObjectURL(url);

      if (data.results) setResults(data.results);

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
              <span className="drop-sub">.xlsx files only · max 4 MB</span>
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
