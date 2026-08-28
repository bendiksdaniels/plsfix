/// <reference types="vite/client" />

// Set by vite.config.ts / vitest.config.ts `define`, sourced from
// package.json's "version" - see src/ui/version.ts.
declare const __APP_VERSION__: string;

// @types/office-js 1.0.606 ships no OfficeRuntime declaration. The link
// workspace store needs only its key/value storage, so declare that much and
// probe it with `typeof OfficeRuntime` - the global is absent outside Office.
declare const OfficeRuntime:
  | {
      storage: {
        getItem(key: string): Promise<string | null>;
        setItem(key: string, value: string): Promise<void>;
        removeItem(key: string): Promise<void>;
      };
    }
  | undefined;
