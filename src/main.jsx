import React from "react";
import { createRoot } from "react-dom/client";
import StoreTransferApp from "../store_transfer_system.jsx";
import "./styles.css";

createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <StoreTransferApp />
  </React.StrictMode>,
);
