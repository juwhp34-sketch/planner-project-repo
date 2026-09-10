// Supabase Edge Function: breakdown-goal
//
// Takes a goal/task in plain language and asks Claude to break it into 3-5 concrete,
// tiny, physically-doable next actions — written for someone with ADHD who is stuck
// and needs the very first move spelled out, not a project plan.
//
// This function exists so the Anthropic API key never has to live in the browser —
// it stays a secret on Supabase's servers, set via `supabase secrets set`.
//
// Deploy with: supabase functions deploy breakdown-goal

const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type, x-planner-secret",
};

Deno.serve(async (req: Request) => {
  if (req.method === "OPTIONS") {
    return new Response(null, { headers: corsHeaders });
  }

  try {
    const SHARED_SECRET = Deno.env.get("PLANNER_SHARED_SECRET");
    const ANTHROPIC_API_KEY = Deno.env.get("ANTHROPIC_API_KEY");

    const incomingSecret = req.headers.get("x-planner-secret");
    if (!SHARED_SECRET || incomingSecret !== SHARED_SECRET) {
      return new Response(JSON.stringify({ error: "Unauthorized" }), {
        status: 401,
        headers: { ...corsHeaders, "Content-Type": "application/json" },
      });
    }
    if (!ANTHROPIC_API_KEY) {
      return new Response(JSON.stringify({ error: "Server is missing ANTHROPIC_API_KEY — set it with `supabase secrets set`." }), {
        status: 500,
        headers: { ...corsHeaders, "Content-Type": "application/json" },
      });
    }

    const { goal } = await req.json();
    if (!goal || typeof goal !== "string" || !goal.trim()) {
      return new Response(JSON.stringify({ error: "Missing goal text." }), {
        status: 400,
        headers: { ...corsHeaders, "Content-Type": "application/json" },
      });
    }

    const prompt = `You help someone with ADHD who feels overwhelmed and stuck on a goal. Break the following goal into 3 to 5 concrete, tiny, physically-doable next actions, ordered from first to last. Each action should be something that can be started in under 15 minutes with no ambiguity about how to begin. Do not include vague planning/thinking steps like "decide" or "figure out" — make every step a literal physical action a person could just go do.

Goal: "${goal.trim()}"

Respond with ONLY a JSON array of strings, nothing else, no markdown formatting. Example: ["Open the spreadsheet and find row 1", "Copy the email template into a new draft", "Send it to the first contact on the list"]`;

    const anthropicResp = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "content-type": "application/json",
        "x-api-key": ANTHROPIC_API_KEY,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify({
        model: "claude-haiku-4-5-20251001",
        max_tokens: 500,
        messages: [{ role: "user", content: prompt }],
      }),
    });

    if (!anthropicResp.ok) {
      const errText = await anthropicResp.text();
      return new Response(JSON.stringify({ error: "Anthropic API error", detail: errText }), {
        status: 502,
        headers: { ...corsHeaders, "Content-Type": "application/json" },
      });
    }

    const data = await anthropicResp.json();
    const text: string = data?.content?.[0]?.text ?? "[]";

    let steps: unknown;
    try {
      const match = text.match(/\[[\s\S]*\]/);
      steps = JSON.parse(match ? match[0] : text);
    } catch {
      steps = [];
    }
    if (!Array.isArray(steps)) steps = [];
    const cleanSteps = (steps as unknown[])
      .filter((s) => typeof s === "string" && s.trim().length > 0)
      .slice(0, 6);

    return new Response(JSON.stringify({ steps: cleanSteps }), {
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  } catch (e) {
    return new Response(JSON.stringify({ error: String(e) }), {
      status: 500,
      headers: { ...corsHeaders, "Content-Type": "application/json" },
    });
  }
});
