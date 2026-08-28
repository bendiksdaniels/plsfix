// Pure XML renderer for Office add-in manifests: buildManifest(env, spec) turns
// manifest/spec.ts data into manifest text. Owns whitespace, attribute order
// and the V1_0/V1_1 VersionOverrides duplication; has no I/O and no knowledge
// of dev vs prod beyond the ManifestEnvironment it is given.
import type { AddinSpec, ButtonSpec, HostSpec, ManifestEnvironment } from "./spec";

const pad = (block: string, spaces: number): string =>
  block
    .split("\n")
    .map((line) => (line ? " ".repeat(spaces) + line : line))
    .join("\n");

function icons(spec: AddinSpec, indent: number): string {
  return pad(
    [
      "<Icon>",
      `  <bt:Image size="16" resid="${spec.iconResids[16]}"/>`,
      `  <bt:Image size="32" resid="${spec.iconResids[32]}"/>`,
      `  <bt:Image size="80" resid="${spec.iconResids[80]}"/>`,
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
          `  <TaskpaneId>${host.taskpaneId}</TaskpaneId>`,
          `  <SourceLocation resid="${host.urlResid}"/>`,
          `</Action>`,
        ]
      : [
          `<Action xsi:type="ExecuteFunction">`,
          `  <FunctionName>${button.action.name}</FunctionName>`,
          `</Action>`,
        ];
  return [
    `<Control xsi:type="Button" id="SMT.Button.${button.id}">`,
    `  <Label resid="SMT.${button.id}.Label"/>`,
    `  <Supertip>`,
    `    <Title resid="SMT.${button.id}.Label"/>`,
    `    <Description resid="SMT.${button.id}.Tip"/>`,
    `  </Supertip>`,
    icons(spec, 2),
    ...action.map((line) => `  ${line}`),
    `</Control>`,
  ].join("\n");
}

function hostBlock(spec: AddinSpec, host: HostSpec): string {
  return [
    `<Host xsi:type="${host.name}">`,
    `  <Runtimes>`,
    `    <Runtime resid="${host.urlResid}" lifetime="long"/>`,
    `  </Runtimes>`,
    `  <DesktopFormFactor>`,
    `    <FunctionFile resid="${host.urlResid}"/>`,
    `    <ExtensionPoint xsi:type="PrimaryCommandSurface">`,
    `      <CustomTab id="SMT.Tab">`,
    `        <Group id="${host.groupId}">`,
    `          <Label resid="${host.groupId}.Label"/>`,
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
    `<bt:String id="SMT.Tab.Label" DefaultValue="${spec.tabLabel}"/>`,
    ...spec.hosts.flatMap((host) => [
      `<bt:String id="${host.groupId}.Label" DefaultValue="${host.groupLabel}"/>`,
      ...host.buttons.map((b) => `<bt:String id="SMT.${b.id}.Label" DefaultValue="${b.label}"/>`),
    ]),
  ];
  const longs = spec.hosts.flatMap((host) =>
    host.buttons.map((b) => `<bt:String id="SMT.${b.id}.Tip" DefaultValue="${b.tip}"/>`),
  );
  return [
    `<Resources>`,
    `  <bt:Images>`,
    `    <bt:Image id="${spec.iconResids[16]}" DefaultValue="${env.baseUrl}assets/icon-16.png"/>`,
    `    <bt:Image id="${spec.iconResids[32]}" DefaultValue="${env.baseUrl}assets/icon-32.png"/>`,
    `    <bt:Image id="${spec.iconResids[80]}" DefaultValue="${env.baseUrl}assets/icon-80.png"/>`,
    `  </bt:Images>`,
    `  <bt:Urls>`,
    ...spec.hosts.map((host) => `    <bt:Url id="${host.urlResid}" DefaultValue="${env.baseUrl}${host.page}"/>`),
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
      ? [`  <ExtendedOverrides Url="${env.baseUrl}shortcuts.json"/>`]
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
    `<!-- ${env.comment} -->`,
    `<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"`,
    `  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"`,
    `  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"`,
    `  xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"`,
    `  xsi:type="TaskPaneApp">`,
    `  <Id>${spec.id}</Id>`,
    `  <Version>${spec.version}</Version>`,
    `  <ProviderName>${spec.provider}</ProviderName>`,
    `  <DefaultLocale>en-US</DefaultLocale>`,
    `  <DisplayName DefaultValue="${spec.displayName}"/>`,
    `  <Description DefaultValue="${spec.description}"/>`,
    `  <IconUrl DefaultValue="${env.baseUrl}assets/icon-32.png"/>`,
    `  <SupportUrl DefaultValue="${spec.supportUrl}"/>`,
    `  <AppDomains>`,
    `    <AppDomain>${spec.appDomain}</AppDomain>`,
    `  </AppDomains>`,
    `  <Hosts>`,
    ...spec.hosts.map((host) => `    <Host Name="${host.name}"/>`),
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
    `    <SourceLocation DefaultValue="${env.baseUrl}${primary.page}"/>`,
    `  </DefaultSettings>`,
    `  <Permissions>ReadWriteDocument</Permissions>`,
    pad(outer, 2),
    `</OfficeApp>`,
    ``,
  ].join("\n");
}
