// Supabase Edge Function: breakdown-goal
//
// Two modes:
//   - default (task mode): takes a single goal/task and breaks it into 3-5 concrete,
//     tiny, physically-doable next actions — for someone with ADHD who is stuck and
//     needs the very first move spelled out, not a project plan.
//   - mode:"quarter": takes a quarter-long goal for one category and breaks it into
//     one milestone per month plus one concrete task per already-tracked week, so a
//     big goal cascades down into the Plan tab's Quarter -> Week -> Day structure.
//
// This function exists so the Anthropic API key never has to live in the browser —
// it stays a secret on Supabase's servers, set via `supabase secrets set`.
//
// Deploy with: supabase functions deploy breakdown-goal

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type, x-planner-secret",
};

function jsonResponse(body: unknown, status = 200) {
  return new Response(JSON.stringify(body), { status, headers: { ...corsHeaders, "Content-Type": "application/json" } });
}

async function callClaude(ANTHROPIC_API_KEY: string, prompt: string, maxTokens: number) {
  const resp = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "content-type": "application/json",
      "x-api-key": ANTHROPIC_API_KEY,
      "anthropic-version": "2023-06-01",
    },
    body: JSON.stringify({
      model: "claude-haiku-4-5-20251001",
      max_tokens: maxTokens,
      messages: [{ role: "user", content: prompt }],
    }),
  });
  if (!resp.ok) {
    const errText = await resp.text();
    throw new Error("Anthropic API error: " + errText);
  }
  const data = await resp.json();
  return (data?.content?.[0]?.text ?? "") as string;
}

Deno.serve(async (req: Request) => {
  if (req.method === "OPTIONS") {
    return new Response(null, { headers: corsHeaders });
  }

  try {
    const SHARED_SECRET = Deno.env.get("PLANNER_SHARED_SECRET");
    const ANTHROPIC_API_KEY = Deno.env.get("ANTHROPIC_API_KEY");

    const incomingSecret = req.headers.get("x-planner-secret");
    if (!SHARED_SECRET || incomingSecret !== SHARED_SECRET) {
      return jsonResponse({ error: "Unauthorized" }, 401);
    }
    if (!ANTHROPIC_API_KEY) {
      return jsonResponse({ error: "Server is missing ANTHROPIC_API_KEY — set it with `supabase secrets set`." }, 500);
    }

    const body = await req.json();
    const goal = typeof body.goal === "string" ? body.goal.trim() : "";
    if (!goal) return jsonResponse({ error: "Missing goal text." }, 400);

    // ---------- Quarter mode: goal -> month milestones + per-week tasks ----------
    if (body.mode === "quarter") {
      const months: string[] = Array.isArray(body.months) ? body.months.filter((m: unknown) => typeof m === "string") : [];
      const weeks: { start: string; end: string }[] = Array.isArray(body.weeks)
        ? body.weeks.filter((w: unknown) => w && typeof (w as any).start === "string" && typeof (w as any).end === "string")
        : [];
      if (months.length !== 3) return jsonResponse({ error: "Quarter mode needs exactly 3 month names." }, 400);

      const weeksList = weeks.length
        ? weeks.map((w) => `  - key "${w.start}" (this week runs ${w.start} through ${w.end})`).join("\n")
        : "  (no weeks are being tracked yet for this quarter)";

      const backlog = typeof body.backlogNotes === "string" ? body.backlogNotes.trim() : "";
      const backlogBlock = backlog
        ? `\nIMPORTANT — real status update: the following did NOT get finished in earlier months and is still outstanding: "${backlog}". Do NOT assume earlier months went as planned. The remaining months/weeks below need to absorb this leftover work FIRST, in addition to their own new goals — do not write the remaining months as if they're just a light final wrap-up when there's real unfinished backlog to carry.\n`
        : "";

      const prompt = `You help someone with ADHD plan a quarter-long goal by breaking it into a month-by-month milestone and week-by-week tasks, so a big goal turns into small, doable pieces spread across the quarter instead of one overwhelming block.

Goal for this category: "${goal}"

This goal spans three months, in order: ${months[0]}, ${months[1]}, ${months[2]}.
${backlogBlock}
These specific weeks are already being tracked. Each one below shows the EXACT key you must use for it in your JSON output:
${weeksList}

For each of the 3 months, write ONE short milestone sentence describing what should realistically be true by the end of that month if this goal is on track — factoring in any leftover backlog noted above. Milestones should build on each other across the 3 months.

For each week listed above, write ONE short, concrete task for that specific week that makes real progress toward the goal for whichever month that week falls in. Do NOT just repeat the month's milestone verbatim for every week in that month — break the month's work into a sensible progression across its weeks (e.g. week 1 might start something, week 2 continues it, week 3 finishes it). If there's leftover backlog, put it in the EARLIEST upcoming weeks, not spread evenly or left until the last week. If no weeks are listed, return an empty object for "weeks".

CRITICAL: in your JSON output, each key inside "weeks" must be EXACTLY the key string shown above (e.g. "${weeks[0]?.start ?? "2026-01-05"}") — just the start date, never a date range, never including the word "to" or the end date.

Respond with ONLY a JSON object of exactly this shape, nothing else, no markdown formatting:
{"months": ["<month 1 milestone>", "<month 2 milestone>", "<month 3 milestone>"], "weeks": {"<week start date exactly as given>": "<that week's task>"}}`;

      const text = await callClaude(ANTHROPIC_API_KEY, prompt, 900);
      let parsed: any = {};
      try {
        const match = text.match(/\{[\s\S]*\}/);
        parsed = JSON.parse(match ? match[0] : text);
      } catch {
        parsed = {};
      }
      const outMonths = Array.isArray(parsed.months) ? parsed.months.filter((m: unknown) => typeof m === "string").slice(0, 3) : [];
      const outWeeks: Record<string, string> = {};
      if (parsed.weeks && typeof parsed.weeks === "object") {
        const parsedKeys = Object.keys(parsed.weeks);
        for (const w of weeks) {
          // Prefer an exact key match; fall back to any key that starts with the week's
          // start date, in case the model still prepends/appends extra text to the key.
          const exactKey = parsedKeys.includes(w.start) ? w.start : parsedKeys.find((k) => k.startsWith(w.start));
          const v = exactKey ? parsed.weeks[exactKey] : undefined;
          if (typeof v === "string" && v.trim()) outWeeks[w.start] = v.trim();
        }
      }
      if (outMonths.length !== 3) return jsonResponse({ error: "Didn't get a usable breakdown back — try rephrasing the goal." }, 502);
      return jsonResponse({ months: outMonths, weeks: outWeeks });
    }

    // ---------- Default mode: goal -> 3-5 tiny next actions ----------
    const prompt = `You help someone with ADHD who feels overwhelmed and stuck on a goal. Break the following goal into 3 to 5 concrete, tiny, physically-doable next actions, ordered from first to last. Each action should be something that can be started in under 15 minutes with no ambiguity about how to begin. Do not include vague planning/thinking steps like "decide" or "figure out" — make every step a literal physical action a person could just go do.

Goal: "${goal}"

Respond with ONLY a JSON array of strings, nothing else, no markdown formatting. Example: ["Open the spreadsheet and find row 1", "Copy the email template into a new draft", "Send it to the first contact on the list"]`;

    const text = await callClaude(ANTHROPIC_API_KEY, prompt, 500);
    let steps: unknown;
    try {
      const match = text.match(/\[[\s\S]*\]/);
      steps = JSON.parse(match ? match[0] : text);
    } catch {
      steps = [];
    }
    if (!Array.isArray(steps)) steps = [];
    const cleanSteps = (steps as unknown[]).filter((s) => typeof s === "string" && s.trim().length > 0).slice(0, 6);
    if (cleanSteps.length === 0) return jsonResponse({ error: "Didn't get any steps back — try rephrasing the goal." }, 502);

    return jsonResponse({ steps: cleanSteps });
  } catch (e) {
    return jsonResponse({ error: String(e) }, 500);
  }
});
