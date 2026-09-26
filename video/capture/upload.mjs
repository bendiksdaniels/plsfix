// upload.mjs: puts the demo workbook and the demo deck into a fresh run folder under
// "plsfix-video" in the signed-in OneDrive through SharePoint REST, run inside a page on that
// site (the rig profile's cookies), and answers each file's id and edit URL. A run gets its own
// folder because SharePoint keeps a file locked (423) for a while after its tab closes. Invariant:
// it writes only under plsfix-video, and removes only older run-* folders it made itself.
import { readFileSync } from "node:fs";
import { basename } from "node:path";

export const FOLDER = "plsfix-video";
// The add-in's registration on a web document (scripts/rig/README.md, step 6).
const ADDIN =
  "&wdaddindevserverport=3001&wdaddinmanifestfile=manifest.prod.xml" +
  "&wdaddinmanifestguid=FF1B34D8-DD7D-4B39-8FA9-6248CA09DB6E";

// site: https://<tenant>-my.sharepoint.com/personal/<user> (never written to the repo)
export async function upload(ctx, site, paths) {
  const page = await ctx.newPage();
  try {
    await page.goto(`${site}/_api/web/title`, { waitUntil: "domcontentloaded" });
    const files = paths.map((p) => ({ name: basename(p), b64: readFileSync(p).toString("base64") }));
    const rel = new URL(site).pathname;
    const out = await page.evaluate(
      async ({ site, rel, folder, files }) => {
        const json = { Accept: "application/json;odata=nometadata" };
        const info = await fetch(`${site}/_api/contextinfo`, { method: "POST", headers: json });
        if (!info.ok) throw new Error(`upload: contextinfo ${info.status}`);
        const digest = (await info.json()).FormDigestValue;
        const post = (url, body, extra = {}) =>
          fetch(url, { method: "POST", headers: { ...json, "X-RequestDigest": digest, ...extra }, body });
        const root = `${rel}/Documents/${folder}`;
        const run = `run-${new Date().toISOString().replace(/[-:]/g, "").slice(0, 15)}`;
        const dir = `${root}/${run}`;
        for (const d of [root, dir]) {
          const made = await post(`${site}/_api/web/folders/add(@u)?@u='${encodeURIComponent(d)}'`);
          if (!made.ok) throw new Error(`upload: folder ${made.status}`);
        }
        // older runs of this tool; a folder still locked by an open tab stays for the next run
        const list = await fetch(`${site}/_api/web/GetFolderByServerRelativeUrl(@d)/Folders?@d='${encodeURIComponent(root)}'`, { headers: json });
        const old = list.ok ? (await list.json()).value.map((f) => f.Name).filter((n) => /^run-\d{8}T\d{6}$/.test(n) && n !== run) : [];
        for (const n of old) {
          await post(`${site}/_api/web/GetFolderByServerRelativeUrl(@d)?@d='${encodeURIComponent(`${root}/${n}`)}'`, null, { "X-HTTP-Method": "DELETE", "IF-MATCH": "*" });
        }
        const answers = [];
        for (const f of files) {
          const bytes = Uint8Array.from(atob(f.b64), (c) => c.charCodeAt(0));
          const url =
            `${site}/_api/web/GetFolderByServerRelativeUrl(@d)/Files/add(url=@n,overwrite=true)` +
            `?@d='${encodeURIComponent(dir)}'&@n='${encodeURIComponent(f.name)}'`;
          const r = await post(url, bytes, { "Content-Type": "application/octet-stream" });
          if (!r.ok) throw new Error(`upload: ${f.name} ${r.status}`);
          const meta = await r.json();
          answers.push({ name: f.name, id: meta.UniqueId, bytes: bytes.length });
        }
        return answers;
      },
      { site, rel, folder: FOLDER, files },
    );
    return out.map((f) => ({ ...f, url: editUrl(site, f.id, f.name) }));
  } finally {
    await page.close();
  }
}

export function editUrl(site, id, name) {
  return (
    `${site}/_layouts/15/Doc.aspx?sourcedoc=%7B${id}%7D&file=${encodeURIComponent(name)}` +
    `&action=edit${ADDIN}`
  );
}
