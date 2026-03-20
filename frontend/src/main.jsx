import React, { useState } from "react";
import ReactDOM from "react-dom/client";
import App from "./App";
import UserManual from "./UserManual";
import ReleaseNotes from "./ReleaseNotes";
import Nav from "./Nav";
import "./App.css";

function Root() {
  const [page, setPage] = useState("home");

  return (
    <>
      <Nav page={page} setPage={setPage} />
      {page === "home" && <App />}
      {page === "manual" && <UserManual setPage={setPage} />}
      {page === "release" && <ReleaseNotes setPage={setPage} />}
    </>
  );
}

ReactDOM.createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <Root />
  </React.StrictMode>
);
