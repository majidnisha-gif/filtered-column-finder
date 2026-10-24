/* global Office, document */
import { findLastFilteredColumn } from "../taskpane/filterFinder";

async function runJump() {
  try {
    const result = await findLastFilteredColumn();
    const status = document.getElementById("status");
    if (status) {
      status.className = result ? "success" : "error";
      status.textContent = result
        ? `Jumped to ${result.address}. Detected: ${result.detected.join(" | ")}`
        : "No active filters found on this sheet.";
    }
  } catch (error) {
    const status = document.getElementById("status");
    if (status) {
      status.className = "error";
      status.textContent = `Could not find the filtered column: ${error.message || error}`;
    }
  }
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    Office.actions.associate("jumpFromRibbon", (event) => {
      runJump().finally(() => event.completed());
    });
    Office.actions.associate("jumpFromKeyboard", runJump);
  }
});
