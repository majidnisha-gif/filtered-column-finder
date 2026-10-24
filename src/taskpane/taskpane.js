/* global Office, document */
import { findLastFilteredColumn } from "./filterFinder";

const button = document.getElementById("find-button");
const status = document.getElementById("status");

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    button.addEventListener("click", runJump);
  }
});

export async function runJump() {
  button.disabled = true;
  status.className = "";
  status.textContent = "Checking filters on this sheet…";

  try {
    const result = await findLastFilteredColumn();
    if (!result) {
      status.className = "error";
      status.textContent = "No active filters found on this sheet.";
      return;
    }

    status.className = "success";
    status.textContent = `Jumped to ${result.address}. Detected: ${result.detected.join(" | ")}`;
  } catch (error) {
    status.className = "error";
    status.textContent = `Could not find the filtered column: ${error.message || error}`;
  } finally {
    button.disabled = false;
  }
}
