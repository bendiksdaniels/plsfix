// Ranks the pane's tools by what they do: typing part of a label or part of
// its help sentence should find the button, not just an exact name match.
// Pure so it is unit-tested without a DOM; the catalogue that turns the real
// markup into ToolEntry[] lives in src/pane/tool-search.ts.

/** One searchable tool: a button's identity, where it lives, what it does. */
export interface ToolEntry {
  /** The button's data-action, else its id - what running the tool clicks. */
  action: string;
  label: string;
  tab: string;
  sentence: string;
}

const LABEL_WEIGHT = 3;
const SENTENCE_WEIGHT = 1;
const WHOLE_LABEL_PREFIX_BONUS = 5;

type Hit = "label" | "sentence" | null;

function wordsOf(text: string): string[] {
  return text.split(/\s+/).filter((word) => word.length > 0);
}

// The four-tier priority from the spec, collapsed to which side of the card
// a token counts against: a prefix of a word in the label and a plain
// substring of the label both read as a label hit, the same two checks
// against the sentence both read as a sentence hit, and a label hit always
// beats a sentence hit for the same token because it is tried first.
function hitOf(token: string, label: string, sentence: string): Hit {
  if (wordsOf(label).some((word) => word.startsWith(token))) return "label";
  if (label.includes(token)) return "label";
  if (wordsOf(sentence).some((word) => word.startsWith(token)))
    return "sentence";
  if (sentence.includes(token)) return "sentence";
  return null;
}

/**
 * Every tool whose label or sentence carries all of the query's words, best
 * match first. A tool missing even one word is left out entirely - this is
 * an AND filter, not a fuzzy score over however many words happened to hit.
 * Ties keep the order `tools` was given in (Array#sort is stable). An empty
 * or whitespace-only query matches nothing.
 */
export function rankTools(
  query: string,
  tools: ToolEntry[],
  limit = 8,
): ToolEntry[] {
  const tokens = query
    .toLowerCase()
    .split(/\s+/)
    .filter((token) => token.length > 0);
  if (tokens.length === 0) return [];
  const wholeQuery = tokens.join(" ");

  const ranked: { tool: ToolEntry; score: number }[] = [];
  for (const tool of tools) {
    const label = tool.label.toLowerCase();
    const sentence = tool.sentence.toLowerCase();
    let score = 0;
    let matchedEveryToken = true;
    for (const token of tokens) {
      const hit = hitOf(token, label, sentence);
      if (hit === null) {
        matchedEveryToken = false;
        break;
      }
      score += hit === "label" ? LABEL_WEIGHT : SENTENCE_WEIGHT;
    }
    if (!matchedEveryToken) continue;
    if (label.startsWith(wholeQuery)) score += WHOLE_LABEL_PREFIX_BONUS;
    ranked.push({ tool, score });
  }

  ranked.sort((a, b) => b.score - a.score);
  return ranked.slice(0, limit).map((entry) => entry.tool);
}
