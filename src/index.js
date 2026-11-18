import React from "react";
import ReactDOM from "react-dom/client";
import "./index.css";

// 1. Update the import path and component name:
import EdmontonInfillDashboard from "./EdmontonInfillDashboard";

const root = ReactDOM.createRoot(document.getElementById("root"));
root.render(
  <React.StrictMode>
    {/* 2. Update the rendered tag here: */}
    <EdmontonInfillDashboard />
  </React.StrictMode>,
);
