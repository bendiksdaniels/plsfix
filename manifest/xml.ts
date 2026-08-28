// Pure XML renderer for Office add-in manifests: buildManifest(env, spec) turns
// manifest/spec.ts data into manifest text. Owns whitespace, attribute order,
// the V1_0/V1_1 VersionOverrides duplication, and escaping every interpolated
// value; has no I/O and no knowledge of dev vs prod beyond what it is given.
import type { AddinSpec, ButtonSpec, HostSpec, ManifestEnvironment } from "./spec";

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

function hostBlock(spec: AddinSpec, host: HostSpec): string {
  return [
    `<Host xsi:type="${escapeXml(host.name)}">`,
    `  <Runtimes>`,
    `    <Runtime resid="${escapeXml(host.urlResid)}" lifetime="long"/>`,
    `  </Runtimes>`,
    `  <DesktopFormFactor>`,
    `    <FunctionFile resid="${escapeXml(host.urlResid)}"/>`,
    `    <ExtensionPoint xsi:type="PrimaryCommandSurface">`,
    `      <CustomTab id="SMT.Tab">`,
    `        <Group id="${escapeXml(host.groupId)}">`,
    `          <Label resid="${escapeXml(host.groupId)}.Label"/>`,
    pad(icons(spec, 0), 10),
    ...host.buttons.map((button) => pad(control(spec, host, button), 10)),
    `        </Group>`,
    `        <Label resid="SMT.Tab.Label"/>`,
    `      </CustomTab>`,
    `    </ExtensionPoint>`,
    `  </DesktopFormFactor>`,
    `</Host>`,
  ].join("\n");
}

function resources(env: ManifestEnvironment, spec: AddinSpec): string {
  const shorts = [
    `<bt:String id="SMT.Tab.Label" DefaultValue="${escapeXml(spec.tabLabel)}"/>`,
    ...spec.hosts.flatMap((host) => [
      `<bt:String id="${escapeXml(host.groupId)}.Label" DefaultValue="${escapeXml(host.groupLabel)}"/>`,
      ...host.buttons.map(
        (b) => `<bt:String id="SMT.${escapeXml(b.id)}.Label" DefaultValue="${escapeXml(b.label)}"/>`,
      ),
    ]),
  ];
  const longs = spec.hosts.flatMap((host) =>
    host.buttons.map((b) => `<bt:String id="SMT.${escapeXml(b.id)}.Tip" DefaultValue="${escapeXml(b.tip)}"/>`),
  );
  return [
    `<Resources>`,
    `  <bt:Images>`,
    `    <bt:Image id="${escapeXml(spec.iconResids[16])}" DefaultValue="${escapeXml(env.baseUrl)}assets/icon-16.png"/>`,
    `    <bt:Image id="${escapeXml(spec.iconResids[32])}" DefaultValue="${escapeXml(env.baseUrl)}assets/icon-32.png"/>`,
    `    <bt:Image id="${escapeXml(spec.iconResids[80])}" DefaultValue="${escapeXml(env.baseUrl)}assets/icon-80.png"/>`,
    `  </bt:Images>`,
    `  <bt:Urls>`,
    ...spec.hosts.map(
      (host) => `    <bt:Url id="${escapeXml(host.urlResid)}" DefaultValue="${escapeXml(env.baseUrl)}${escapeXml(host.page)}"/>`,
    ),
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

function overrides(env: ManifestEnvironment, spec: AddinSpec, nested: string | null): string {
  const xmlns = nested === null
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

export function buildManifest(env: ManifestEnvironment, spec: AddinSpec): string {
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
    // A top-level requirement applies to every host; ExcelApi would hide the
    // add-in in PowerPoint, so it is declared only while Excel is the sole host.
    ...(spec.hosts.every((host) => host.name === "Workbook")
      ? [
          `  <Requirements>`,
          `    <Sets DefaultMinVersion="1.9">`,
          `      <Set Name="ExcelApi" MinVersion="1.9"/>`,
          `    </Sets>`,
          `  </Requirements>`,
        ]
      : []),
    `  <DefaultSettings>`,
    `    <SourceLocation DefaultValue="${escapeXml(env.baseUrl)}${escapeXml(primary.page)}"/>`,
    `  </DefaultSettings>`,
    `  <Permissions>ReadWriteDocument</Permissions>`,
    pad(outer, 2),
    `</OfficeApp>`,
    ``,
  ].join("\n");
}
