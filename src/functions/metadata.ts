// The typed source of public/functions.json, the registration file Office reads
// to publish =SMT.ROUND and =SMT.ROUNDSUM (the SMT. prefix comes from the
// manifest's <Namespace>, so the ids here stay unqualified). Every id must be
// associated in src/functions/index.ts: a registered function with nothing
// behind it returns #N/A in the cell.
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
  result: { type: "number"; dimensionality: "scalar" };
  parameters: FunctionParameterSpec[];
}

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
      "The group's total, rounded so it equals the sum of SMT.ROUND over the same range.",
    result: SCALAR_NUMBER,
    parameters: [RANGE, DECIMALS],
  },
];

// Office reads this file as-is; the trailing newline keeps it a well-behaved
// text file in git.
export function functionsMetadata(): string {
  return `${JSON.stringify({ functions: CUSTOM_FUNCTIONS }, null, 2)}\n`;
}
