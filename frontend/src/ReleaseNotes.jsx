import { useState } from "react";
import { downloadPdf } from "./pdfUtils";

const SITE_URL = "https://test-link-navy.vercel.app";
const REPO_URL = "https://github.com/701428/Test_link";

export default function ReleaseNotes({ setPage }) {
  const [exporting, setExporting] = useState(false);

  const handlePdf = async () => {
    setExporting(true);
    await downloadPdf("release-notes-content", "SharePoint_Test_Report_Linker_Release_Notes.pdf");
    setExporting(false);
  };

  return (
    <div className="container doc-page">
      <div className="pdf-toolbar">
        <button className="btn btn-pdf" onClick={handlePdf} disabled={exporting}>
          {exporting ? "Generating PDF…" : "⬇ Download PDF"}
        </button>
      </div>

      <div id="release-notes-content">
      <header>
        <div className="doc-badge">Release Notes</div>
        <h1>SharePoint Test Report Linker</h1>
        <p className="subtitle">
          Version history and changelog for the SharePoint Test Report Linker tool.
        </p>
        <p className="hint" style={{ marginTop: 8 }}>
          Live tool:{" "}
          <a href={SITE_URL} target="_blank" rel="noopener noreferrer">{SITE_URL}</a>
          {" · "}
          Source:{" "}
          <a href={REPO_URL} target="_blank" rel="noopener noreferrer">GitHub</a>
        </p>
      </header>

      {/* v1.2.0 */}
      <section className="card release-card">
        <div className="release-header">
          <span className="release-version">v1.2.0</span>
          <span className="release-date">20 March 2026</span>
          <span className="release-tag tag-latest">Latest</span>
        </div>
        <p className="release-title">Client-Side Processing — No More 413 Error</p>
        <div className="release-section">
          <div className="release-section-label">What's New</div>
          <ul className="doc-list">
            <li>
              <strong>Fully browser-based processing:</strong> The Excel file is now read, modified,
              and downloaded entirely inside the browser using{" "}
              <a href="https://github.com/exceljs/exceljs" target="_blank" rel="noopener noreferrer">ExcelJS</a>.
              No file is sent to the server.
            </li>
            <li>
              <strong>No file size limit:</strong> The previous Vercel serverless function had a
              hard 4.5 MB request body limit (HTTP 413). Since there is no longer any upload, files
              of any size are supported.
            </li>
            <li>
              <strong>Faster processing:</strong> No network round-trip means results are available
              immediately, limited only by browser speed.
            </li>
            <li>
              <strong>Improved privacy:</strong> Your Excel data never leaves your device.
            </li>
          </ul>
        </div>
        <div className="release-section">
          <div className="release-section-label">Bug Fixes</div>
          <ul className="doc-list">
            <li>
              Fixed <code>vercel.json</code> conflict where <code>functions</code> and{" "}
              <code>builds</code> blocks were both present, causing GitHub Actions deployment to
              fail silently and the site to stay on the old version.
            </li>
            <li>
              Removed incorrect 4 MB client-side file size gate (now unnecessary since processing
              is local).
            </li>
          </ul>
        </div>
        <div className="release-section">
          <div className="release-section-label">Technical Changes</div>
          <ul className="doc-list">
            <li>Added <code>exceljs</code> as a frontend dependency.</li>
            <li>
              Removed <code>fetch()</code> call to <code>/api/process</code> from the main
              processing path.
            </li>
            <li>
              Added <code>buildSheetUrl()</code> helper in JavaScript to replicate the Python URL
              encoding logic exactly (uses <code>encodeURIComponent</code> matching Python's{" "}
              <code>quote(path, safe="")</code>).
            </li>
            <li>Merged-cell detection and unmerging now handled via ExcelJS <code>_merges</code> API.</li>
          </ul>
        </div>
      </section>

      {/* v1.1.0 */}
      <section className="card release-card">
        <div className="release-header">
          <span className="release-version">v1.1.0</span>
          <span className="release-date">19 March 2026</span>
        </div>
        <p className="release-title">Vercel Routing Fix & 413 Awareness</p>
        <div className="release-section">
          <div className="release-section-label">Changes</div>
          <ul className="doc-list">
            <li>
              Fixed API routing on Vercel — added <code>/api</code> prefix to all backend routes
              so that <code>/api/process</code> and <code>/api/health</code> resolve correctly in
              production.
            </li>
            <li>
              Added client-side 4 MB file size check to surface a user-friendly error before the
              upload reaches the Vercel 4.5 MB limit.
            </li>
            <li>
              Added explicit HTTP 413 error message on the frontend instead of a generic
              "Server error 413".
            </li>
            <li>
              Fixed Vercel build by adding a root <code>package.json</code> to drive the frontend
              static build correctly.
            </li>
            <li>
              Added GitHub Actions workflow (<code>.github/workflows/deploy.yml</code>) for
              automatic Vercel deployment on every push to <code>main</code>.
            </li>
          </ul>
        </div>
      </section>

      {/* v1.0.0 */}
      <section className="card release-card">
        <div className="release-header">
          <span className="release-version">v1.0.0</span>
          <span className="release-date">19 March 2026</span>
          <span className="release-tag tag-initial">Initial Release</span>
        </div>
        <p className="release-title">Initial Release</p>
        <div className="release-section">
          <div className="release-section-label">Features</div>
          <ul className="doc-list">
            <li>
              React + Vite frontend with drag-and-drop Excel file upload.
            </li>
            <li>
              FastAPI Python backend (<code>@vercel/python</code>) that reads <code>.xlsx</code>{" "}
              files using <code>openpyxl</code>.
            </li>
            <li>
              Automatically finds <strong>"Test Report Path"</strong> cells across all sheets and
              inserts a blue-underline hyperlink to the corresponding SharePoint subfolder.
            </li>
            <li>
              Handles SharePoint <code>?id=</code> URL format with correct percent-encoding.
            </li>
            <li>
              Automatically unmerges cells that overlap the target hyperlink cell.
            </li>
            <li>
              Returns a per-sheet results table showing which sheets were linked and which were
              skipped.
            </li>
            <li>
              Deployed to Vercel with SPA routing support.
            </li>
          </ul>
        </div>
        <div className="release-section">
          <div className="release-section-label">Known Limitations at Release</div>
          <ul className="doc-list">
            <li>Backend upload limited to 4.5 MB due to Vercel serverless function constraints (resolved in v1.2.0).</li>
            <li>Only <code>.xlsx</code> format supported; <code>.xls</code> and <code>.xlsm</code> are not.</li>
          </ul>
        </div>
      </section>

      </div>{/* end #release-notes-content */}

      <div style={{ textAlign: "center", marginTop: 32 }}>
        <button className="btn btn-primary" onClick={() => setPage("home")}>
          ← Back to Tool
        </button>
      </div>
    </div>
  );
}
