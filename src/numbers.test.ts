import { describe, expect, it } from "vitest";

import {
  currencyFormat,
  formatAmount,
  isLanguage,
  localizeNumberText,
  separatorSample,
  separatorsMatch,
} from "./numbers";

describe("formatAmount", () => {
  it("groups thousands with a space in Latvian and Russian, a comma in English", () => {
    expect(formatAmount(1094417.5, "lv")).toBe("1 094 417.5");
    expect(formatAmount(1094417.5, "ru")).toBe("1 094 417.5");
    expect(formatAmount(1094417.5, "en")).toBe("1,094,417.5");
  });

  it("always writes the decimal as a point and drops trailing zeros", () => {
    expect(formatAmount(2436, "lv")).toBe("2 436");
    expect(formatAmount(2436.04, "lv")).toBe("2 436");
    expect(formatAmount(0.25, "lv", 2)).toBe("0.25");
    expect(formatAmount(999, "en")).toBe("999");
  });

  it("keeps the sign, but never prints -0", () => {
    expect(formatAmount(-1234.56, "en")).toBe("-1,234.6");
    expect(formatAmount(-0.01, "en")).toBe("0");
  });
});

describe("currencyFormat", () => {
  it("puts the symbol after the amount in Latvian and Russian, before it in English", () => {
    expect(currencyFormat("€", "lv", "#,##0")).toBe("#,##0 €");
    expect(currencyFormat("€", "ru", "#,##0.0")).toBe("#,##0.0 €");
    expect(currencyFormat("€", "en", "#,##0")).toBe("€ #,##0");
  });

  it("leaves the digits alone without a symbol", () => {
    expect(currencyFormat("", "lv", "#,##0")).toBe("#,##0");
  });
});

describe("separators", () => {
  it("samples what Excel will show and compares it to the house style", () => {
    expect(separatorSample(".", " ")).toBe("1 094 417.5");
    expect(separatorSample(",", ".")).toBe("1.094.417,5");
    expect(separatorsMatch(".", " ", "lv")).toBe(true);
    expect(separatorsMatch(".", ",", "lv")).toBe(false);
    expect(separatorsMatch(".", ",", "en")).toBe(true);
  });

  it("knows the three languages and nothing else", () => {
    expect(isLanguage("lv")).toBe(true);
    expect(isLanguage("de")).toBe(false);
    expect(isLanguage(3)).toBe(false);
  });
});

describe("localizeNumberText", () => {
  it("rewrites the invariant thousands comma to the application's separator", () => {
    expect(localizeNumberText("12,400", { decimal: ".", thousands: " " })).toBe(
      "12 400",
    );
  });

  it("rewrites a parenthesised negative the same way", () => {
    expect(
      localizeNumberText("(7,688)", { decimal: ".", thousands: " " }),
    ).toBe("(7 688)");
  });

  it("rewrites the invariant decimal point to a comma", () => {
    expect(localizeNumberText("8.0%", { decimal: ",", thousands: " " })).toBe(
      "8,0%",
    );
  });

  it("swaps both separators in one pass, never a two-step replace", () => {
    expect(
      localizeNumberText("1,234.56", { decimal: ",", thousands: "." }),
    ).toBe("1.234,56");
  });

  it("leaves a text cell untouched", () => {
    expect(
      localizeNumberText("Revenue", { decimal: ",", thousands: " " }),
    ).toBe("Revenue");
  });

  it("leaves a currency-prefixed cell untouched", () => {
    expect(
      localizeNumberText("EUR 12,400", { decimal: ",", thousands: " " }),
    ).toBe("EUR 12,400");
  });

  it("is the identity when the application already shows the invariant separators", () => {
    expect(
      localizeNumberText("12,400.5", { decimal: ".", thousands: "," }),
    ).toBe("12,400.5");
  });
});
