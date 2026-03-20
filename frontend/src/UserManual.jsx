const SITE_URL = "https://test-link-navy.vercel.app";

export default function UserManual({ setPage }) {
  return (
    <div className="container doc-page">
      <header>
        <div className="doc-badge">User Manual</div>
        <h1>SharePoint Test Report Linker</h1>
        <p className="subtitle">
          A browser-based tool that automatically inserts clickable SharePoint folder hyperlinks into
          your Excel test report tracker — no installation required.
        </p>
        <p className="hint" style={{ marginTop: 8 }}>
          Live tool:{" "}
          <a href={SITE_URL} target="_blank" rel="noopener noreferrer">
            {SITE_URL}
          </a>
        </p>
      </header>

      {/* Overview */}
      <section className="card">
        <div className="step-label">Overview</div>
        <p className="hint">
          This tool reads your Excel file, finds every sheet that contains a cell labelled
          <strong> "Test Report Path"</strong>, and places a clickable hyperlink in the cell
          immediately to the right of it. The hyperlink points to the corresponding SharePoint
          subfolder for that sheet (e.g. sheet <code>IMDS_001</code> → folder <code>…/IMDS_001</code>).
        </p>
        <p className="hint" style={{ marginTop: 8 }}>
          Everything runs <strong>inside your browser</strong> — the file is never uploaded to any
          server, so there are no file size restrictions and no data privacy concerns.
        </p>
      </section>

      {/* Prerequisites */}
      <section className="card">
        <div className="step-label">Prerequisites</div>
        <ul className="doc-list">
          <li>
            <strong>Excel file format:</strong> The file must be <code>.xlsx</code> (Excel 2007+).
            Older <code>.xls</code> files are not supported.
          </li>
          <li>
            <strong>"Test Report Path" label:</strong> Each sheet that needs a hyperlink must
            contain exactly the text <code>Test Report Path</code> (case-insensitive) in one cell.
            The hyperlink will be written to the cell immediately to its right.
          </li>
          <li>
            <strong>SharePoint folder URL:</strong> You need the browser address-bar URL of the
            parent folder in SharePoint that contains one subfolder per sheet name
            (e.g. <code>IMDS_001</code>, <code>IMDS_002</code>).
          </li>
          <li>
            <strong>Modern browser:</strong> Chrome, Edge, Firefox, or Safari (latest versions).
          </li>
        </ul>
      </section>

      {/* Step by step */}
      <section className="card">
        <div className="step-label">Step-by-Step Instructions</div>

        <div className="doc-step">
          <div className="doc-step-num">1</div>
          <div className="doc-step-body">
            <strong>Download the Excel file from SharePoint</strong>
            <p className="hint">
              Open the file in SharePoint Online → click the three-dot menu → select
              <em> Download</em>. Save the <code>.xlsx</code> file to your computer.
            </p>
            <div className="callout callout-warning">
              Do <strong>not</strong> open the file in Excel Online and export — that can change
              formatting. Use the Download option directly.
            </div>
          </div>
        </div>

        <div className="doc-step">
          <div className="doc-step-num">2</div>
          <div className="doc-step-body">
            <strong>Open the tool</strong>
            <p className="hint">
              Go to{" "}
              <a href={SITE_URL} target="_blank" rel="noopener noreferrer">
                {SITE_URL}
              </a>{" "}
              in your browser.
            </p>
          </div>
        </div>

        <div className="doc-step">
          <div className="doc-step-num">3</div>
          <div className="doc-step-body">
            <strong>Upload the Excel file (Step 1 on the tool)</strong>
            <p className="hint">
              Click the drop zone or drag-and-drop your <code>.xlsx</code> file. The filename and
              size will appear once loaded.
            </p>
          </div>
        </div>

        <div className="doc-step">
          <div className="doc-step-num">4</div>
          <div className="doc-step-body">
            <strong>Paste the SharePoint Base Folder URL (Step 2 on the tool)</strong>
            <p className="hint">
              In SharePoint, navigate to the folder that contains all the test-result subfolders
              (e.g. the folder that has <code>IMDS_001</code>, <code>IMDS_002</code> inside it).
              Copy the full URL from the browser address bar and paste it into the input field.
            </p>
            <div className="callout callout-info">
              <strong>Example URL format:</strong>
              <br />
              <code style={{ fontSize: 12, wordBreak: "break-all" }}>
                https://yourcompany-my.sharepoint.com/personal/user/Documents/Test%20Results
              </code>
              <br />
              The tool will append each sheet name to this URL automatically.
            </div>
          </div>
        </div>

        <div className="doc-step">
          <div className="doc-step-num">5</div>
          <div className="doc-step-body">
            <strong>Click "Insert Hyperlinks & Download" (Step 3 on the tool)</strong>
            <p className="hint">
              The file is processed locally in your browser. A modified file named
              <code> originalname_linked.xlsx</code> is automatically downloaded to your computer.
              The Results table below shows which sheets were linked and which were not found.
            </p>
          </div>
        </div>

        <div className="doc-step">
          <div className="doc-step-num">6</div>
          <div className="doc-step-body">
            <strong>Upload the modified file back to SharePoint (Step 4)</strong>
            <p className="hint">
              In SharePoint, navigate to the location of the original file → click
              <em> Upload</em> → select the <code>_linked.xlsx</code> file. When prompted, choose
              to <em>Replace</em> the existing file.
            </p>
            <div className="callout callout-warning">
              Make sure you replace the <strong>correct file</strong> in SharePoint. Uploading to
              the wrong folder will not update the original tracker.
            </div>
          </div>
        </div>
      </section>

      {/* Important notes */}
      <section className="card">
        <div className="step-label">Important Notes & Limitations</div>
        <ul className="doc-list">
          <li>
            <strong>One label per sheet:</strong> The tool finds the <em>first</em> occurrence of
            "Test Report Path" per sheet and adds a link to the cell on its right. Duplicate labels
            in the same sheet are ignored.
          </li>
          <li>
            <strong>Merged cells:</strong> If the target cell (to the right of the label) is part
            of a merged region, the tool automatically unmerges it before writing the hyperlink.
            Visual formatting of the surrounding cells is preserved.
          </li>
          <li>
            <strong>Sheet names must match folder names exactly:</strong> The tool appends the
            sheet name verbatim to the base URL. If your SharePoint folder is named{" "}
            <code>IMDS 001</code> (with a space) but the sheet is <code>IMDS_001</code> (with an
            underscore), the link will not resolve correctly.
          </li>
          <li>
            <strong>SharePoint URL encoding:</strong> The tool correctly percent-encodes the full
            path including slashes (<code>%2F</code>) to match SharePoint's <code>?id=</code> URL
            format. Do not manually edit the pasted URL.
          </li>
          <li>
            <strong>Browser processing only:</strong> The file is processed entirely in your
            browser. No data is sent to any server. Refreshing the page clears everything.
          </li>
          <li>
            <strong>File must be closed in Excel:</strong> If you have the file open in desktop
            Excel while processing, the download will succeed but uploading back to SharePoint may
            cause a conflict. Close the file in Excel first.
          </li>
        </ul>
      </section>

      {/* Troubleshooting */}
      <section className="card">
        <div className="step-label">Troubleshooting</div>
        <table className="results-table">
          <thead>
            <tr>
              <th>Problem</th>
              <th>Likely Cause</th>
              <th>Fix</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td>Sheet shows <em>"Test Report path label not found"</em></td>
              <td>The cell text doesn't match exactly, or the sheet has no such label</td>
              <td>
                Open the file in Excel and confirm a cell contains exactly the text{" "}
                <code>Test Report Path</code> (any capitalisation is fine)
              </td>
            </tr>
            <tr>
              <td>Hyperlink opens the wrong SharePoint folder</td>
              <td>Sheet name doesn't match the folder name in SharePoint</td>
              <td>Rename the sheet or the SharePoint folder so they match exactly</td>
            </tr>
            <tr>
              <td>Downloaded file appears corrupted in Excel</td>
              <td>The original file had unsupported features (macros, embedded objects)</td>
              <td>
                Remove macros / embedded objects, save as plain <code>.xlsx</code>, then re-process
              </td>
            </tr>
            <tr>
              <td>Base URL field shows a validation error</td>
              <td>URL is missing or malformed</td>
              <td>Copy the URL directly from the SharePoint browser address bar</td>
            </tr>
            <tr>
              <td>Nothing happens after clicking the button</td>
              <td>Browser blocked the automatic download</td>
              <td>Allow pop-ups / downloads for this site in your browser settings</td>
            </tr>
          </tbody>
        </table>
      </section>

      <div style={{ textAlign: "center", marginTop: 32 }}>
        <button className="btn btn-primary" onClick={() => setPage("home")}>
          ← Back to Tool
        </button>
      </div>
    </div>
  );
}
