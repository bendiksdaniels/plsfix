// The typed source of public/functions.json, the registration file Office reads
// to publish =PLSFIX.ROUND, =PLSFIX.ROUNDSUM and =PLSFIX.CAGR (the PLSFIX.
// prefix comes from the manifest's <Namespace>, so the ids here stay
// unqualified). Every id must be associated in src/functions/index.ts: a
// registered function with nothing behind it returns #N/A in the cell.
//
// Written by scripts/build-functions-metadata.ts, checked by npm run check.

export interface FunctionParameterSpec {
  name: string;
  description: string;
  type: "number";
  dimensionality: "scalar" | "matrix";
}

export interface CustomFunctionSpec {
  id: string;
  name: string;
  description: string;
  /** Where the function's own help lives; the store validator wants one. */
  helpUrl: string;
  result: { type: "number"; dimensionality: "scalar" };
  parameters: FunctionParameterSpec[];
}

// One page for all three, anchored at the functions chapter.
const HELP_URL = "https://dbautomatizacijas.com/modelis/support.html#functions";

const RANGE: FunctionParameterSpec = {
  name: "range",
  description:
    "The whole group of numbers being rounded together, as a range reference.",
  type: "number",
  dimensionality: "matrix",
};

const DECIMALS: FunctionParameterSpec = {
  name: "decimals",
  description:
    "Decimal places, as in ROUND: 0 for whole numbers, negative for tens and thousands.",
  type: "number",
  dimensionality: "scalar",
};

const SCALAR_NUMBER = { type: "number", dimensionality: "scalar" } as const;

export const CUSTOM_FUNCTIONS: readonly CustomFunctionSpec[] = [
  {
    id: "ROUND",
    name: "ROUND",
    description:
      "One cell's share of a consistently rounded group: the rounded values add up to the rounded total.",
    helpUrl: HELP_URL,
    result: SCALAR_NUMBER,
    parameters: [
      RANGE,
      {
        name: "index",
        description: "Which cell of the range this is, counting from 1.",
        type: "number",
        dimensionality: "scalar",
      },
      DECIMALS,
    ],
  },
  {
    id: "ROUNDSUM",
    name: "ROUNDSUM",
    description:
      "The group's total, rounded so it equals the sum of PLSFIX.ROUND over the same range.",
    helpUrl: HELP_URL,
    result: SCALAR_NUMBER,
    parameters: [RANGE, DECIMALS],
  },
  {
    id: "CAGR",
    name: "CAGR",
    description:
      "The compound annual growth rate from a first value to a last one over a number of periods.",
    helpUrl: HELP_URL,
    result: SCALAR_NUMBER,
    parameters: [
      {
        name: "first",
        description: "The value at the start of the first period.",
        ...SCALAR_NUMBER,
      },
      {
        name: "last",
        description: "The value at the end of the last period.",
        ...SCALAR_NUMBER,
      },
      {
        name: "periods",
        description:
          "How many periods the growth ran over, usually years. At least one.",
        ...SCALAR_NUMBER,
      },
    ],
  },
];

// Office reads this file as-is; the trailing newline keeps it a well-behaved
// text file in git.
export function functionsMetadata(): string {
  return `${JSON.stringify({ functions: CUSTOM_FUNCTIONS }, null, 2)}\n`;
}
