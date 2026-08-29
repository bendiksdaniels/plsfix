// Pure XML renderer for Office add-in manifests: buildManifest(env, spec) turns
// manifest/spec.ts data into manifest text. Owns whitespace, attribute order,
// the V1_0/V1_1 VersionOverrides duplication, and escaping every interpolated
// value; has no I/O and no knowledge of dev vs prod beyond what it is given.
// One thing it deliberately cannot express: the XML manifest has no per-host
// <Requirements>, so with PowerPoint declared beside Excel the ExcelApi 1.9
// floor is enforced at runtime in the Excel pane boot (src/main.ts) and the
// custom functions are gated by their Workbook-only extension point instead.
import type {
  AddinSpec,
  ButtonSpec,
  GroupSpec,
  HostSpec,
  ManifestEnvironment,
} from "./spec";

const pad = (block: string, spaces: number): string =>
  block
    .split("\n")
    .map((line) => (line ? " ".repeat(spaces) + line : line))
    .join("\n");

// Escapes the five XML-significant characters for use in attribute values and
// element text content. Every interpolated spec/env value goes through this,
// except env.comment: that string is written inside a raw <!-- --> comment
// node, where XML does not parse entity references, so escaping it would
// misrender the comment instead of protecting anything.
const escapeXml = (text: string): string =>
  text
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");

function icons(spec: AddinSpec, indent: number): string {
  return pad(
    [
      "<Icon>",
      `  <bt:Image size="16" resid="${escapeXml(spec.iconResids[16])}"/>`,
      `  <bt:Image size="32" resid="${escapeXml(spec.iconResids[32])}"/>`,
      `  <bt:Image size="80" resid="${escapeXml(spec.iconResids[80])}"/>`,
      "</Icon>",
    ].join("\n"),
    indent,
  );
}

function control(spec: AddinSpec, host: HostSpec, button: ButtonSpec): string {
  const action =
    button.action.kind === "showPane"
      ? [
          `<Action xsi:type="ShowTaskpane">`,
          `  <TaskpaneId>${escapeXml(host.taskpaneId)}</TaskpaneId>`,
          `  <SourceLocation resid="${escapeXml(host.urlResid)}"/>`,
          `</Action>`,
        ]
      : [
          `<Action xsi:type="ExecuteFunction">`,
          `  <FunctionName>${escapeXml(button.action.name)}</FunctionName>`,
          `</Action>`,
        ];
  return [
    `<Control xsi:type="Button" id="SMT.Button.${escapeXml(button.id)}">`,
    `  <Label resid="SMT.${escapeXml(button.id)}.Label"/>`,
    `  <Supertip>`,
    `    <Title resid="SMT.${escapeXml(button.id)}.Label"/>`,
    `    <Description resid="SMT.${escapeXml(button.id)}.Tip"/>`,
    `  </Supertip>`,
    icons(spec, 2),
    ...action.map((line) => `  ${line}`),
    `</Control>`,
  ].join("\n");
}

// One ribbon group: its own label resource, the shared three-size icon, and
// its buttons in order. Local indent 0/2 so pad(groupBlock(...), 8) below
// lands each group at the same depth the single inline group used to sit at.
function groupBlock(spec: AddinSpec, host: HostSpec, group: GroupSpec): string {
  return [
    `<Group id="${escapeXml(group.id)}">`,
    `  <Label resid="${escapeXml(group.id)}.Label"/>`,
    pad(icons(spec, 0), 2),
    ...group.buttons.map((button) => pad(control(spec, host, button), 2)),
    `</Group>`,
  ].join("\n");
}

// The custom functions Office publishes for this host, or nothing when the
// host declares none. <Page> points at the host's own pane URL: the functions
// share the pane's long-lived runtime rather than starting a second one. Order
// inside <Host> is fixed by the schema: Runtimes, AllFormFactors, then
// DesktopFormFactor.
function customFunctions(host: HostSpec): string[] {
  const functions = host.customFunctions;
  if (!functions) return [];
  return [
    `<AllFormFactors>`,
    `  <ExtensionPoint xsi:type="CustomFunctions">`,
    `    <Script>`,
    `      <SourceLocation resid="${escapeXml(functions.scriptResid)}"/>`,
    `    </Script>`,
    `    <Page>`,
    `      <SourceLocation resid="${escapeXml(host.urlResid)}"/>`,
    `    </Page>`,
    `    <Metadata>`,
    `      <SourceLocation resid="${escapeXml(functions.metadataResid)}"/>`,
    `    </Metadata>`,
    `    <Namespace resid="${escapeXml(functions.namespaceResid)}"/>`,
    `  </ExtensionPoint>`,
    `</AllFormFactors>`,
  ];
}

function hostBlock(spec: AddinSpec, host: HostSpec): string {
  return [
    `<Host xsi:type="${escapeXml(host.name)}">`,
    `  <Runtimes>`,
    `    <Runtime resid="${escapeXml(host.urlResid)}" lifetime="long"/>`,
    `  </Runtimes>`,
    ...customFunctions(host).map((line) => `  ${line}`),
    `  <DesktopFormFactor>`,
    `    <FunctionFile resid="${escapeXml(host.urlResid)}"/>`,
    `    <ExtensionPoint xsi:type="PrimaryCommandSurface">`,
    `      <CustomTab id="SMT.Tab">`,
    ...host.groups.map((group) => pad(groupBlock(spec, host, group), 8)),
    `        <Label resid="SMT.Tab.Label"/>`,
    `      </CustomTab>`,
    `    </ExtensionPoint>`,
    `  </DesktopFormFactor>`,
    `</Host>`,
  ].join("\n");
}

const ICON_SIZES = [16, 32, 80] as const;

// One <Resources> block serves every host in its VersionOverrides, and the ids
// are built from button and group ids, so two hosts that share a button id (a
// second OpenPane, a Push button on both) would emit one id twice. Office
// rejects that manifest, and the generated-bytes gate cannot see it because
// both sides are equally wrong - so the generator refuses to emit it at all.
function assertUniqueResourceIds(lines: string[]): void {
  const seen = new Set<string>();
  for (const line of lines) {
    const id = /\bid="([^"]*)"/.exec(line)?.[1];
    if (id === undefined) continue;
    if (seen.has(id)) throw new Error(`manifest: duplicate resource id ${id}`);
    seen.add(id);
  }
}

function resources(env: ManifestEnvironment, spec: AddinSpec): string {
  const images = ICON_SIZES.map(
    (size) =>
      `<bt:Image id="${escapeXml(spec.iconResids[size])}" DefaultValue="${escapeXml(env.baseUrl)}assets/icon-${size}.png"/>`,
  );
  const urls = [
    ...spec.hosts.map(
      (host) =>
        `<bt:Url id="${escapeXml(host.urlResid)}" DefaultValue="${escapeXml(env.baseUrl)}${escapeXml(host.page)}"/>`,
    ),
    ...spec.hosts.flatMap((host) =>
      host.customFunctions
        ? [
            `<bt:Url id="${escapeXml(host.customFunctions.scriptResid)}" DefaultValue="${escapeXml(env.baseUrl)}${escapeXml(host.customFunctions.scriptFile)}"/>`,
            `<bt:Url id="${escapeXml(host.customFunctions.metadataResid)}" DefaultValue="${escapeXml(env.baseUrl)}${escapeXml(host.customFunctions.metadataFile)}"/>`,
          ]
        : [],
    ),
  ];
  const shorts = [
    `<bt:String id="SMT.Tab.Label" DefaultValue="${escapeXml(spec.tabLabel)}"/>`,
    ...spec.hosts.flatMap((host) =>
      host.groups.flatMap((group) => [
        `<bt:String id="${escapeXml(group.id)}.Label" DefaultValue="${escapeXml(group.label)}"/>`,
        ...group.buttons.map(
          (b) =>
            `<bt:String id="SMT.${escapeXml(b.id)}.Label" DefaultValue="${escapeXml(b.label)}"/>`,
        ),
      ]),
    ),
    ...spec.hosts.flatMap((host) =>
      host.customFunctions
        ? [
            `<bt:String id="${escapeXml(host.customFunctions.namespaceResid)}" DefaultValue="${escapeXml(host.customFunctions.namespace)}"/>`,
          ]
        : [],
    ),
  ];
  const longs = spec.hosts.flatMap((host) =>
    host.groups.flatMap((group) =>
      group.buttons.map(
        (b) =>
          `<bt:String id="SMT.${escapeXml(b.id)}.Tip" DefaultValue="${escapeXml(b.tip)}"/>`,
      ),
    ),
  );
  assertUniqueResourceIds([...images, ...urls, ...shorts, ...longs]);
  return [
    `<Resources>`,
    `  <bt:Images>`,
    ...images.map((line) => `    ${line}`),
    `  </bt:Images>`,
    `  <bt:Urls>`,
    ...urls.map((line) => `    ${line}`),
    `  </bt:Urls>`,
    `  <bt:ShortStrings>`,
    ...shorts.map((line) => `    ${line}`),
    `  </bt:ShortStrings>`,
    `  <bt:LongStrings>`,
    ...longs.map((line) => `    ${line}`),
    `  </bt:LongStrings>`,
    `</Resources>`,
  ].join("\n");
}

function overrides(
  env: ManifestEnvironment,
  spec: AddinSpec,
  nested: string | null,
): string {
  const xmlns =
    nested === null
      ? `xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides/1.1" xsi:type="VersionOverridesV1_1"`
      : `xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0"`;
  return [
    `<VersionOverrides ${xmlns}>`,
    `  <Requirements>`,
    `    <bt:Sets DefaultMinVersion="1.1">`,
    `      <bt:Set Name="SharedRuntime" MinVersion="1.1"/>`,
    `    </bt:Sets>`,
    `  </Requirements>`,
    `  <Hosts>`,
    ...spec.hosts.map((host) => pad(hostBlock(spec, host), 4)),
    `  </Hosts>`,
    pad(resources(env, spec), 2),
    ...(nested === null
      ? [`  <ExtendedOverrides Url="${escapeXml(env.baseUrl)}shortcuts.json"/>`]
      : [pad(nested, 2)]),
    `</VersionOverrides>`,
  ].join("\n");
}

// A top-level requirement applies to every host, and an ExcelApi set would
// hide the add-in in PowerPoint, so the block is emitted only while Excel is
// the sole host - which, with PowerPoint declared, means no ExcelApi floor is
// published anywhere. The floor did not go away, it moved: the Excel pane boot
// (src/main.ts) refuses below ExcelApi 1.9, and the custom functions are gated
// by their Workbook-only extension point.
function requirements(spec: AddinSpec): string[] {
  if (!spec.hosts.every((host) => host.name === "Workbook")) return [];
  return [
    `  <Requirements>`,
    `    <Sets DefaultMinVersion="1.9">`,
    `      <Set Name="ExcelApi" MinVersion="1.9"/>`,
    `    </Sets>`,
    `  </Requirements>`,
  ];
}

export function buildManifest(
  env: ManifestEnvironment,
  spec: AddinSpec,
): string {
  const inner = overrides(env, spec, null);
  const outer = overrides(env, spec, inner);
  const primary = spec.hosts[0]!;
  return [
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`,
    // env.comment is raw comment text, not escaped - see the escapeXml note above.
    `<!-- ${env.comment} -->`,
    `<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"`,
    `  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"`,
    `  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"`,
    `  xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"`,
    `  xsi:type="TaskPaneApp">`,
    `  <Id>${escapeXml(spec.id)}</Id>`,
    `  <Version>${escapeXml(spec.version)}</Version>`,
    `  <ProviderName>${escapeXml(spec.provider)}</ProviderName>`,
    `  <DefaultLocale>en-US</DefaultLocale>`,
    `  <DisplayName DefaultValue="${escapeXml(spec.displayName)}"/>`,
    `  <Description DefaultValue="${escapeXml(spec.description)}"/>`,
    `  <IconUrl DefaultValue="${escapeXml(env.baseUrl)}assets/icon-32.png"/>`,
    `  <SupportUrl DefaultValue="${escapeXml(spec.supportUrl)}"/>`,
    `  <AppDomains>`,
    `    <AppDomain>${escapeXml(spec.appDomain)}</AppDomain>`,
    `  </AppDomains>`,
    `  <Hosts>`,
    ...spec.hosts.map((host) => `    <Host Name="${escapeXml(host.name)}"/>`),
    `  </Hosts>`,
    ...requirements(spec),
    `  <DefaultSettings>`,
    `    <SourceLocation DefaultValue="${escapeXml(env.baseUrl)}${escapeXml(primary.page)}"/>`,
    `  </DefaultSettings>`,
    `  <Permissions>ReadWriteDocument</Permissions>`,
    pad(outer, 2),
    `</OfficeApp>`,
    ``,
  ].join("\n");
}
