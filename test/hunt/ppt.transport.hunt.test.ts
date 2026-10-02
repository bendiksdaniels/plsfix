// Hunt pass 1: src/ppt/transport.ts's preBootTransport, the placeholder
// main.ts paints before Office.onReady opens the real one (no button ever
// reaches it, so its three methods had no test). Pins the contract: relay
// mode, unpaired, no store, and a setMode that resolves doing nothing.

import { describe, expect, it } from "vitest";
import type { RelayApi } from "../../src/link/relay";
import { preBootTransport } from "../../src/ppt/transport";

describe("preBootTransport", () => {
  it("answers relay mode, the given RelayApi, no workspace and no store", () => {
    const remote = {} as RelayApi;
    const transport = preBootTransport(remote);

    expect(transport.mode()).toBe("relay");
    expect(transport.relay()).toBe(remote);
    expect(transport.workspace()).toBeNull();
    expect(transport.store()).toBeNull();
  });

  it("setMode resolves without changing anything a button could observe", async () => {
    const transport = preBootTransport({} as RelayApi);

    await expect(transport.setMode("local")).resolves.toBeUndefined();

    expect(transport.mode()).toBe("relay");
    expect(transport.store()).toBeNull();
  });
});
