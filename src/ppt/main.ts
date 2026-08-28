// PowerPoint pane boot: connects to Office.js, gates on the PowerPoint host
// and PowerPointApi 1.5, then wires the shared toast, error reporting, tabs
// and version footer. Task 15 fills in the Links/Inbox/Settings views.

import "../styles.css";
import { installErrorReporting } from "../ui/report";
import { installTabs } from "../ui/tabs";
import { createToast } from "../ui/toast";
import { formatVersion } from "../ui/version";

const APP_VERSION = formatVersion(__APP_VERSION__);

const getElement = <T extends HTMLElement>(id: string): T => {
  const element = document.getElementById(id);
  if (!element) throw new Error(`Missing element #${id}`);
  return element as T;
};

const connectionStatus = getElement<HTMLSpanElement>("connection-status");
const toast = createToast(getElement("toast"));

getElement("app-version").textContent = APP_VERSION;

// Installed first so a throw during the rest of boot is still reported.
installErrorReporting(
  { host: "PowerPoint", version: APP_VERSION },
  (message, details) => toast.show(message, "error", details),
);
installTabs(getElement("tab-bar"));

Office.onReady(({ host }) => {
  if (host !== Office.HostType.PowerPoint) {
    connectionStatus.textContent = "PowerPoint required";
    connectionStatus.className = "connection error";
    return;
  }

  if (!Office.context.requirements.isSetSupported("PowerPointApi", "1.5")) {
    connectionStatus.textContent = "PowerPoint 2021 / Microsoft 365 required";
    connectionStatus.className = "connection error";
    return;
  }

  connectionStatus.textContent = "PowerPoint connected";
  connectionStatus.className = "connection ready";
});
