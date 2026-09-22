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
import { dispatch, EXCEL_NOT_CONNECTED_MESSAGE } from "./pane/dispatch";
import { renderFind, runFind } from "./pane/find-panel";
import { installLinksTab } from "./pane/links-tab";
import { copyReport, renderModelCheck } from "./pane/model-check-panel";
import {
  loadPaintSlots,
  loadWorkbookSlots,
  renderPaintSlots,
} from "./pane/paint-slots";
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
import { installShortcutsPanel } from "./pane/shortcuts-panel";
import {
  armStyles,
  disarmStyles,
  isStylesArmed,
  renderStyles,
} from "./pane/styles-panel";
import { installToolSearch } from "./pane/tool-search";
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
// The Brand tab's shortcut manager; its own card button opens the same dialog.
void installShortcutsPanel(document, () => {
  void guard(() => dispatch("shortcut-card"), "shortcut-card");
});
getElement<HTMLButtonElement>("tab-workbook").addEventListener(
  "click",
  // Sheets change without the pane hearing about it, so the explorer is
  // read when it comes into view rather than on every selection change.
  () => void refreshSheets(),
);
// Office's shared runtime can fire a ribbon button or keyboard shortcut the
// instant the pane's script has loaded, well before hostReady() settles
// (worst on Excel for the web, where the custom-functions init is slow) -
// registering here, unconditionally, matches src/ppt/main.ts's own
// module-top-level call, so Office always finds an association for the
// FunctionName instead of a silent no-op. Each command's own promise chain
// (src/pane/commands.ts) reports whatever Excel.run refuses at that point.
registerCommands();
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
const OVERLAY_RESTORE_FAILED_MESSAGE =
  "Last session's audit overlay could not be put back: toggle it once to clear the stripes.";

async function restoreOverlayFills(): Promise<void> {
  try {
    if (await restorePersistedOverlay()) {
      toast.show("Audit overlay fills from the last session were restored.");
    }
  } catch {
    // A host refusing workbook settings, or a corrupt snapshot, leaves last
    // session's stripes on the sheet with no way back to plain fills except
    // the toggle - said once, rather than swallowed, so that is discoverable.
    toast.show(OVERLAY_RESTORE_FAILED_MESSAGE, "error");
  }
}

// Everything here needs a live, supported Excel: the workbook-scoped setup
// that only makes sense once boot() has confirmed the host, in the order the
// tabs depend on. Returns the Links tab handle so boot() can wire the
// selection-changed listener that feeds it.
async function connectExcel(
  degraded: boolean,
): Promise<ReturnType<typeof installLinksTab>> {
  connectionStatus.textContent = "Excel connected";
  connectionStatus.className = "connection ready";
  setExcelReady(true);
  // Slots saved with the workbook win over the machine-local fallback loaded
  // before Office was ready, so a shared model carries its formatting kit.
  void loadWorkbookSlots().catch(() => undefined);
  if (degraded) toast.show(DEGRADED_BOOT_MESSAGE);
  void readSeparators()
    .then(applyExcelSeparators)
    .catch(() => undefined);

  syncAutocolorOnEdit();
  void adoptWorkbookBrand();

  // Before the Links tab, whose own boot restores the linked-cell highlight:
  // the two snapshots cover overlapping cells, so the order they go back in
  // decides whose paint the modeller is left with.
  await restoreOverlayFills();

  return installLinksTab({
    guard,
    toast,
    relay: new RelayClient(relayBaseUrl(document.baseURI)),
    keyStore: officeKeyStore(),
    root: document,
  });
}

// Wired regardless of host state: dispatch() and every named handler below
// already refuse with one sentence (via isExcelReady()) when Excel is not
// connected, so a rejected or absent host still gets a pane that switches
// tabs, searches and answers every click instead of leaving buttons dead.
function wireControls(): void {
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
}

async function boot(host: Office.HostType, degraded: boolean): Promise<void> {
  if (host === Office.HostType.PowerPoint) {
    location.replace("pptpane.html");
    return;
  }

  // The manifest no longer guarantees ExcelApi 1.9 at the top level now that
  // PowerPoint is a second host, so the pane checks for itself. Whichever
  // way this comes out, wireControls() below still runs: a host that is not
  // Excel, or too old a one, must not leave every button and the search box
  // dead - each one already refuses gracefully without a connected workbook.
  const excelConnected =
    host === Office.HostType.Excel &&
    Office.context.requirements.isSetSupported("ExcelApi", "1.9");

  if (host !== Office.HostType.Excel) {
    connectionStatus.textContent = "Excel required";
    connectionStatus.className = "connection error";
  } else if (!excelConnected) {
    connectionStatus.textContent = "Excel 2021 / Microsoft 365 required";
    connectionStatus.className = "connection error";
  }

  const linksTab = excelConnected ? await connectExcel(degraded) : null;

  wireControls();

  // installLinksTab() lives inside connectExcel() and never runs without a
  // connected, supported Excel, so its 16 id-wired buttons (export, push,
  // project and link-key) would otherwise sit with no listener at all -
  // dead rather than refusing gracefully like every [data-action] button.
  if (!excelConnected) {
    for (const button of document.querySelectorAll<HTMLButtonElement>(
      "#view-links button",
    )) {
      button.addEventListener("click", () => {
        toast.show(EXCEL_NOT_CONNECTED_MESSAGE, "error");
      });
    }
  }

  if (excelConnected) {
    // Debounced: dragging a selection fires the event continuously.
    let selectionTimer: number | undefined;
    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      () => {
        window.clearTimeout(selectionTimer);
        selectionTimer = window.setTimeout(() => {
          void refreshSelection();
          void linksTab?.sheetChanged();
        }, 150);
      },
    );
  }

  // Last: the catalogue it builds needs every other tab already wired.
  installToolSearch(document);

  if (excelConnected) await refreshSelection();
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
    // The active worksheet's name, not the workbook's own: any real,
    // trivial round trip proves the host is genuinely there, and this one
    // is the pattern every other adapter already uses (unlike
    // Workbook.load, which nothing else in this codebase ever calls).
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");
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
