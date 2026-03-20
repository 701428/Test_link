export default function Nav({ page, setPage }) {
  return (
    <nav className="topnav">
      <div className="topnav-inner">
        <span className="topnav-brand" onClick={() => setPage("home")} style={{ cursor: "pointer" }}>
          SharePoint Test Report Linker
        </span>
        <div className="topnav-links">
          <button
            className={`topnav-link ${page === "home" ? "active" : ""}`}
            onClick={() => setPage("home")}
          >
            Tool
          </button>
          <button
            className={`topnav-link ${page === "manual" ? "active" : ""}`}
            onClick={() => setPage("manual")}
          >
            User Manual
          </button>
          <button
            className={`topnav-link ${page === "release" ? "active" : ""}`}
            onClick={() => setPage("release")}
          >
            Release Notes
          </button>
        </div>
      </div>
    </nav>
  );
}
