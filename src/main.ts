import "./styles.css";
import {
  addIfError,
  applyNumberFormat,
  applyPreset,
  autocolorSelection,
  clearFormats,
  fastFill,
  inspectSelection,
  scaleSelection,
  type NumberFormatName,
  type PresetName,
} from "./excel";

const getElement = <T extends HTMLElement>(id: string): T => {
  const element = document.getElementById(id);
  if (!element) throw new Error(`Missing element #${id}`);
  return element as T;
};

const connectionStatus = getElement<HTMLSpanElement>("connection-status");
const toast = getElement<HTMLDivElement>("toast");
const actionButtons = Array.from(
  document.querySelectorAll<HTMLButtonElement>("[data-action]"),
);
let toastTimer: number | undefined;

function showToast(message: string, kind: "success" | "error" = "success"): void {
  window.clearTimeout(toastTimer);
  toast.textContent = message;
  toast.className = `toast visible ${kind}`;
  toastTimer = window.setTimeout(() => {
    toast.className = "toast";
  }, 3200);
}

function setBusy(busy: boolean): void {
  for (const button of actionButtons) button.disabled = busy;
}

function errorMessage(error: unknown): string {
  if (error instanceof Error) return error.message;
  return "Excel could not complete that action.";
}

async function refreshSelection(): Promise<void> {
  try {
    const summary = await inspectSelection();
    getElement("selection-address").textContent = summary.address;
    getElement("metric-cells").textContent = summary.cells.toLocaleString();
    getElement("metric-formulas").textContent = summary.formulas.toLocaleString();
    getElement("metric-errors").textContent = summary.errors.toLocaleString();
    getElement("metric-blanks").textContent = summary.blanks.toLocaleString();
  } catch (error) {
    showToast(errorMessage(error), "error");
  }
}

async function runAction(action: string): Promise<void> {
  setBusy(true);
  try {
    if (action.startsWith("style-")) {
      await applyPreset(action.replace("style-", "") as PresetName);
    } else if (action.startsWith("number-")) {
      await applyNumberFormat(action.replace("number-", "") as NumberFormatName);
    } else {
      switch (action) {
        case "clear-formats":
          await clearFormats();
          break;
        case "fill-right":
          await fastFill("right");
          break;
        case "fill-down":
          await fastFill("down");
          break;
        case "if-error":
          await addIfError();
          break;
        case "autocolor":
          await autocolorSelection();
          break;
        case "divide-1000":
          await scaleSelection(0.001);
          break;
        case "multiply-1000":
          await scaleSelection(1000);
          break;
        default:
          throw new Error(`Unknown action: ${action}`);
      }
    }

    await refreshSelection();
    showToast("Selection updated");
  } catch (error) {
    showToast(errorMessage(error), "error");
  } finally {
    setBusy(false);
  }
}

Office.onReady(async ({ host }) => {
  if (host !== Office.HostType.Excel) {
    connectionStatus.textContent = "Excel required";
    connectionStatus.className = "connection error";
    return;
  }

  connectionStatus.textContent = "Excel connected";
  connectionStatus.className = "connection ready";

  for (const button of actionButtons) {
    button.addEventListener("click", () => {
      const action = button.dataset.action;
      if (action) void runAction(action);
    });
  }

  getElement<HTMLButtonElement>("refresh-selection").addEventListener(
    "click",
    () => void refreshSelection(),
  );

  Office.context.document.addHandlerAsync(
    Office.EventType.DocumentSelectionChanged,
    () => void refreshSelection(),
  );

  await refreshSelection();
});
