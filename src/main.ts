// Excel pane entry: the top-level render sequence, boot() in the order the
// tabs depend on (overlay restore before the Links tab, brand adoption, then
// the command table) and the host probe. Owns no tab state and no business
// logic: those live in src/pane/ and src/excel/.
import "./styles.css";
import { readSeparators, restorePersistedOverlay } from "./excel";
import { hostReady } from "./host-ready";
import { RelayClient, relayBaseUrl } from "./link/relay";
import { officeKeyStore } from "./link/workspace";
import {
  adoptWorkbookBrand,
  applyExcelSeparators,
  loadSettings,
  renderBrand,
  syncAutocolorOnEdit,
  wireBrand,
} from "./pane/brand-tab";
import { registerCommands } from "./pane/commands";
import { dispatch } from "./pane/dispatch";
import { renderFind, runFind } from "./pane/find-panel";
import { installLinksTab } from "./pane/links-tab";
import { copyReport, renderModelCheck } from "./pane/model-check-panel";
import { loadPaintSlots, renderPaintSlots } from "./pane/paint-slots";
import {
  actionButtons,
  APP_VERSION,
  guard,
  refreshSelection,
  renderActionState,
  setExcelReady,
  toast,
} from "./pane/shared";
import { renderShare } from "./pane/share-panel";
import {
  armStyles,
  disarmStyles,
  isStylesArmed,
  renderStyles,
} from "./pane/styles-panel";
import { renderAuditState, traceBack } from "./pane/trace-panel";
import {
  armDelete,
  deleteNames,
  disarmDelete,
  isDeleteArmed,
  refreshSheets,
  renderNames,
} from "./pane/workbook-tab";
import { getElement } from "./ui/dom";
import { installFirstRun } from "./ui/first-run";
import { installHelp } from "./ui/help";
import { installErrorReporting } from "./ui/report";

const connectionStatus = getElement<HTMLSpanElement>("connection-status");

getElement("app-version").textContent = APP_VERSION;

// Installed first so a throw during the rest of boot is still reported.
installErrorReporting(
  { host: "Excel", version: APP_VERSION },
  (message, details) => toast.show(message, "error", details),
);
loadSettings();
loadPaintSlots();
// The "?" on every section heading, added once the markup is in place.
installHelp(document);
// The Tools tab's "New here?" card; its Shortcut card button runs the same
// dispatch action as the Tools tab's own button.
installFirstRun(document, "plsfix.firstRun.v1", "first-run", () => {
  void guard(() => dispatch("shortcut-card"), "shortcut-card");
});
getElement<HTMLButtonElement>("tab-workbook").addEventListener(
  "click",
  // Sheets change without the pane hearing about it, so the explorer is
  // read when it comes into view rather than on every selection change.
  () => void refreshSheets(),
);
wireBrand();
renderBrand();
renderPaintSlots();
renderAuditState();
renderActionState();
renderNames(false);
renderFind(null);
renderStyles();
renderShare(null);
renderModelCheck();

// Both overlays saved what they covered inside the workbook, and both put those
// fills back at boot. They are restored one after the other, never side by
// side: two floating promises could interleave, and the one that lands second
// would write its snapshot - which is the other one's tint - over cells that
// had just been handed their originals back.
async function restoreOverlayFills(): Promise<void> {
  try {
    if (await restorePersistedOverlay()) {
      toast.show("Audit overlay fills from the last session were restored.");
    }
  } catch {
    return;
  }
}

async function boot(host: Office.HostType, degraded: boolean): Promise<void> {
  if (host === Office.HostType.PowerPoint) {
    location.replace("pptpane.html");
    return;
  }

  if (host !== Office.HostType.Excel) {
    connectionStatus.textContent = "Excel required";
    connectionStatus.className = "connection error";
    return;
  }

  // The manifest no longer guarantees ExcelApi 1.9 at the top level now that
  // PowerPoint is a second host, so the pane checks for itself.
  if (!Office.context.requirements.isSetSupported("ExcelApi", "1.9")) {
    connectionStatus.textContent = "Excel 2021 / Microsoft 365 required";
    connectionStatus.className = "connection error";
    return;
  }

  connectionStatus.textContent = "Excel connected";
  connectionStatus.className = "connection ready";
  setExcelReady(true);
  if (degraded) toast.show(DEGRADED_BOOT_MESSAGE);
  void readSeparators()
    .then(applyExcelSeparators)
    .catch(() => undefined);

  registerCommands();
  syncAutocolorOnEdit();
  void adoptWorkbookBrand();

  // Before the Links tab, whose own boot restores the linked-cell highlight:
  // the two snapshots cover overlapping cells, so the order they go back in
  // decides whose paint the modeller is left with.
  await restoreOverlayFills();

  const linksTab = installLinksTab({
    guard,
    toast,
    relay: new RelayClient(relayBaseUrl(document.baseURI)),
    keyStore: officeKeyStore(),
    root: document,
  });

  for (const button of actionButtons) {
    button.addEventListener("click", () => {
      const action = button.dataset.action;
      if (action) void guard(() => dispatch(action), action);
    });
  }

  getElement<HTMLButtonElement>("refresh-selection").addEventListener(
    "click",
    () => void refreshSelection(),
  );

  getElement<HTMLButtonElement>("trace-back").addEventListener(
    "click",
    () => void guard(traceBack),
  );

  getElement<HTMLButtonElement>("refresh-sheets").addEventListener(
    "click",
    () => void refreshSheets(),
  );

  // Enter is what a search box owes the hands already on the keyboard.
  getElement<HTMLInputElement>("find-query").addEventListener(
    "keydown",
    (event) => {
      if (event.key === "Enter") void guard(runFind, "find");
    },
  );

  // Wired here rather than through the dispatch table: the copy button stays
  // disabled until a report exists, and setBusy re-enables every [data-action]
  // button after each run.
  getElement<HTMLButtonElement>("copy-model-check").addEventListener(
    "click",
    () => void guard(copyReport, "copy-model-check"),
  );

  getElement<HTMLButtonElement>("delete-names").addEventListener(
    "click",
    () => {
      if (!isDeleteArmed()) {
        armDelete();
        return;
      }
      disarmDelete();
      void guard(deleteNames);
    },
  );

  getElement<HTMLButtonElement>("styles-delete").addEventListener(
    "click",
    () => {
      if (!isStylesArmed()) {
        armStyles();
        return;
      }
      disarmStyles();
      void guard(() => dispatch("styles-delete"), "styles-delete");
    },
  );

  // Debounced: dragging a selection fires the event continuously.
  let selectionTimer: number | undefined;
  Office.context.document.addHandlerAsync(
    Office.EventType.DocumentSelectionChanged,
    () => {
      window.clearTimeout(selectionTimer);
      selectionTimer = window.setTimeout(() => {
        void refreshSelection();
        void linksTab.sheetChanged();
      }, 150);
    },
  );

  await refreshSelection();
}

// Office.onReady is the ready signal. On Excel for the web the custom-functions
// runtime can fail to start ("session expired" at its init) and the promise
// then never settles although Excel.run answers, so from HOST_HEAD_START_MS on
// the host itself is asked (src/host-ready.ts) and the pane boots degraded.
const HOST_HEAD_START_MS = 4_000;
const HOST_PROBE_EVERY_MS = 1_000;
const HOST_GIVE_UP_MS = 30_000;
const DEGRADED_BOOT_MESSAGE =
  "Excel answered but never reported the add-in ready: =PLSFIX.ROUND may need the workbook reopened.";

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => window.setTimeout(resolve, ms));
}

async function probeExcelHost(): Promise<Office.HostType | null> {
  try {
    if (Office.context.host !== Office.HostType.Excel) return null;
    await Excel.run(async (context) => {
      context.workbook.load("name");
      await context.sync();
    });
    return Office.HostType.Excel;
  } catch {
    return null;
  }
}

void hostReady({
  onReady: Office.onReady().then(({ host }) => host),
  probe: probeExcelHost,
  headStartMs: HOST_HEAD_START_MS,
  probeEveryMs: HOST_PROBE_EVERY_MS,
  giveUpAfterMs: HOST_GIVE_UP_MS,
  sleep,
}).then(({ host, degraded }) => boot(host, degraded));
