/* global OfficeRuntime, window */

export type ReviewComment = {
  quote?: string;
  comment: string;
  severity?: "info" | "suggestion" | "warning";
};

export type ReviewResult = {
  comments: ReviewComment[];
  summary?: string;
};

export type ApiSettings = {
  apiKey: string;
  baseUrl: string;
  model: string;
};

const STORAGE_KEYS = {
  apiKey: "RM_OPENAI_API_KEY",
  baseUrl: "RM_OPENAI_BASE_URL",
  model: "RM_OPENAI_MODEL",
};

export const DEFAULT_SETTINGS: ApiSettings = {
  apiKey: "",
  baseUrl: "https://api.openai.com/v1",
  model: "gpt-4o-mini",
};

export async function saveApiSettings(settings: Partial<ApiSettings>): Promise<void> {
  const merged = { ...(await loadApiSettings()), ...settings } as ApiSettings;
  try {
    if (typeof OfficeRuntime !== "undefined" && (OfficeRuntime as any)?.storage) {
      await (OfficeRuntime as any).storage.setItem(STORAGE_KEYS.apiKey, merged.apiKey || "");
      await (OfficeRuntime as any).storage.setItem(STORAGE_KEYS.baseUrl, merged.baseUrl || DEFAULT_SETTINGS.baseUrl);
      await (OfficeRuntime as any).storage.setItem(STORAGE_KEYS.model, merged.model || DEFAULT_SETTINGS.model);
      return;
    }
  } catch {
    // ignore and fallback
  }
  try {
    const g: any = (typeof window !== "undefined" ? (window as any) : undefined) as any;
    if (g && g.sessionStorage) {
      g.sessionStorage.setItem(STORAGE_KEYS.apiKey, merged.apiKey || "");
      g.sessionStorage.setItem(STORAGE_KEYS.baseUrl, merged.baseUrl || DEFAULT_SETTINGS.baseUrl);
      g.sessionStorage.setItem(STORAGE_KEYS.model, merged.model || DEFAULT_SETTINGS.model);
    }
  } catch {
    // ignore
  }
}

export async function loadApiSettings(): Promise<ApiSettings> {
  try {
    const storage = (typeof OfficeRuntime !== "undefined" ? (OfficeRuntime as any).storage : undefined) as
      | { getItem: (k: string) => Promise<string | null> }
      | undefined;
    if (storage) {
      const [apiKey, baseUrl, model] = await Promise.all([
        storage.getItem(STORAGE_KEYS.apiKey),
        storage.getItem(STORAGE_KEYS.baseUrl),
        storage.getItem(STORAGE_KEYS.model),
      ]);
      return {
        apiKey: apiKey || "",
        baseUrl: baseUrl || DEFAULT_SETTINGS.baseUrl,
        model: model || DEFAULT_SETTINGS.model,
      };
    }
  } catch {
    // ignore
  }
  const g: any = (typeof window !== "undefined" ? (window as any) : undefined) as any;
  return {
    apiKey: (g && g.sessionStorage ? (g.sessionStorage.getItem(STORAGE_KEYS.apiKey) as string) : "") || "",
    baseUrl:
      (g && g.sessionStorage ? (g.sessionStorage.getItem(STORAGE_KEYS.baseUrl) as string) : DEFAULT_SETTINGS.baseUrl) ||
      DEFAULT_SETTINGS.baseUrl,
    model:
      (g && g.sessionStorage ? (g.sessionStorage.getItem(STORAGE_KEYS.model) as string) : DEFAULT_SETTINGS.model) ||
      DEFAULT_SETTINGS.model,
  };
}

function buildPrompt(input: string, focuses: string[], custom: string): string {
  const focusText = focuses.length ? focuses.join(", ") : "general academic quality";
  return [
    "You are an expert academic peer reviewer. Provide high-quality, constructive, concise feedback.",
    `Focus areas: ${focusText}.`,
    custom ? `Additional reviewer instructions: ${custom}` : "",
    "Return your output as strict JSON only with the following structure:",
    `{
  "comments": [
    { "quote": "<optional short quote>", "comment": "<actionable feedback>", "severity": "info|suggestion|warning" }
  ],
  "summary": "<optional concise overall review summary>"
}`,
    "Constraints:",
    "- No Markdown code fences.",
    "- No extra commentary outside the JSON.",
    "- Keep each comment short and actionable.",
    "- Use at most 10 comments.",
    "Text to review:",
    input,
  ].join("\n");
}

function extractJson(text: string): any {
  try {
    return JSON.parse(text);
  } catch {
    const first = text.indexOf("{");
    const last = text.lastIndexOf("}");
    if (first >= 0 && last > first) {
      const candidate = text.slice(first, last + 1);
      try {
        return JSON.parse(candidate);
      } catch {
        const cleaned = candidate.replace(/```json|```/g, "");
        return JSON.parse(cleaned);
      }
    }
    throw new Error("LLM did not return JSON.");
  }
}

export async function generateReviewComments(
  input: string,
  focuses: string[],
  custom: string,
  settings?: Partial<ApiSettings>
): Promise<ReviewResult> {
  const cfg = { ...(await loadApiSettings()), ...(settings || {}) } as ApiSettings;
  if (!cfg.apiKey) throw new Error("Missing API key. Please set it in Settings.");

  const prompt = buildPrompt(input, focuses, custom);

  const res = await window.fetch(`${cfg.baseUrl.replace(/\/+$/, "")}/chat/completions`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${cfg.apiKey}`,
    },
    body: JSON.stringify({
      model: cfg.model,
      messages: [
        { role: "system", content: "You are an expert academic peer reviewer." },
        { role: "user", content: prompt },
      ],
      temperature: 0.2,
    }),
  });

  if (!res.ok) {
    const t = await res.text();
    throw new Error(`LLM error: ${res.status} ${res.statusText} - ${t}`);
  }

  const data = await res.json();
  const content = data?.choices?.[0]?.message?.content || "";
  const parsed = extractJson(content);

  const comments: ReviewComment[] = Array.isArray(parsed?.comments) ? parsed.comments : [];
  const summary: string | undefined = typeof parsed?.summary === "string" ? parsed.summary : undefined;

  const norm = comments
    .filter((c) => c && typeof c.comment === "string")
    .map((c) => {
      const sev: ReviewComment["severity"] =
        c.severity === "warning" || c.severity === "suggestion" ? c.severity : "info";
      return {
        quote: typeof c.quote === "string" ? c.quote : undefined,
        comment: c.comment,
        severity: sev,
      };
    });

  return { comments: norm.slice(0, 10), summary };
}

async function callSimpleCompletion(prompt: string, settings?: Partial<ApiSettings>): Promise<string> {
  const cfg = { ...(await loadApiSettings()), ...(settings || {}) } as ApiSettings;
  if (!cfg.apiKey) throw new Error("Missing API key. Please set it in Settings.");
  const res = await window.fetch(`${cfg.baseUrl.replace(/\/+$/, "")}/chat/completions`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Authorization: `Bearer ${cfg.apiKey}` },
    body: JSON.stringify({
      model: cfg.model,
      messages: [
        { role: "system", content: "You are a precise writing assistant." },
        { role: "user", content: prompt },
      ],
      temperature: 0.4,
    }),
  });
  if (!res.ok) {
    const t = await res.text();
    throw new Error(`LLM error: ${res.status} ${res.statusText} - ${t}`);
  }
  const data = await res.json();
  const content = data?.choices?.[0]?.message?.content || "";
  return content.trim();
}

export async function generateTrimmedText(
  input: string,
  maxChars: number = 0,
  style: string = "default",
  settings?: Partial<ApiSettings>
): Promise<string> {
  const target = maxChars && maxChars > 50 ? `Aim for about ${maxChars} characters (not strict).` : "Keep concise.";
  const tone =
    style === "simpler"
      ? "Use simpler vocabulary while retaining accuracy."
      : style === "formal"
        ? "Maintain a formal, scholarly tone."
        : style === "concise"
          ? "Prefer brevity and remove redundancy."
          : "Maintain clarity and neutral professional tone.";
  const prompt = [
    "Rewrite the following text into a significantly shorter version while preserving all essential meaning and intent.",
    tone,
    target,
    "Return only the rewritten text. No preface or explanation.",
    input,
  ].join("\n");
  return callSimpleCompletion(prompt, settings);
}

export async function generateParaphrasedText(
  input: string,
  variants: number = 1,
  style: string = "default",
  settings?: Partial<ApiSettings>
): Promise<string[]> {
  const n = Math.min(Math.max(variants, 1), 3);
  const prompt = [
    `Paraphrase the following text into ${n} distinct high-quality alternative version(s).`,
    "Each version must retain the original meaning and vary BOTH wording and sentence structure.",
    style === "simpler"
      ? "Use simpler vocabulary while preserving meaning."
      : style === "formal"
        ? "Adopt a formal scholarly tone."
        : style === "concise"
          ? "Prefer concise, tight phrasing."
          : "Maintain a neutral professional tone.",
    "Output format:",
    "Variant 1: <first paraphrase on same line>",
    "Variant 2: <second paraphrase on same line> (etc)",
    "Do NOT add explanations or any other text before or after the variants.",
    "If you cannot paraphrase, repeat the original text as Variant 1.",
    input,
  ].join("\n");

  const raw = await callSimpleCompletion(prompt, settings);
  const text = raw.trim();
  if (!text) return [];

  const lines = text
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter((l) => !!l);

  // Attempt structured multi-line aggregation: capture header lines and following continuation lines until next header
  const variantsOut: string[] = [];
  let current: { index: number; content: string } | null = null;
  const headerRegex = /^Variant\s+(\d+)\s*:?\s*(.*)$/i;
  for (const ln of lines) {
    const header = ln.match(headerRegex);
    if (header) {
      // Push previous
      if (current) variantsOut.push(current.content.trim());
      const idx = parseInt(header[1], 10);
      const rest = header[2] || "";
      current = { index: idx, content: rest };
    } else if (current) {
      // continuation line
      current.content += (current.content ? " " : "") + ln;
    }
  }
  if (current) variantsOut.push(current.content.trim());

  // Fallback: if nothing matched, treat entire raw as single variant
  const cleaned = variantsOut.map((v) => v.trim()).filter((v) => v.length > 0);
  if (!cleaned.length) return [text];
  return cleaned.slice(0, n);
}
