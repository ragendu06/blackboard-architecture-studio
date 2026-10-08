# Blackboard Architecture Studio

An interactive companion to **Research Brief 1 — Blackboard Architecture for LLM Multi-Agent Systems**.
One self-contained web page with two linked tools, plus a detailed build & run guide.

## What's inside

| File | What it is |
|------|------------|
| `index.html` | The **studio** — one page, two tools, switchable from a top segmented control. No build step, no dependencies, no network needed. |
| `blackboard-variants.html` | Stand-alone copy of **Tool 02** (the variants control room), kept for reference. |
| `Build_and_Run_Guide.docx` | Detailed guide: how to run it and what every UI element does. |
| `tools/gen_guide.js` | The script that generates the guide (see below). |

### Tool 01 — Topology Comparison
Compares five LLM multi-agent coordination topologies (single agent, orchestrator–worker,
group chat, **blackboard**, decentralized queue). Runs seeded Monte-Carlo trials and reports a
**Welch t-test** on token cost, an accuracy-vs-tokens Pareto chart, and a leaderboard.

### Tool 02 — Blackboard Variants
Opens the blackboard itself and runs **seven control-loop variations** as a live, stepped
animation — classical agenda (Hearsay-II), control blackboard (Hayes-Roth), LLM router (bMAS),
volunteer market (Salemi et al.), a learned **cost-aware controller** (Direction A),
credibility-gated verification (Direction B / PatchBoard), and a secure partitioned board
(Direction C / Terrarium). Play / Pause / Step / Restart / Speed, with a cost-to-solve comparison.

> Both tools are mocks: the numbers come from synthetic, seeded models calibrated to the effect
> sizes reported in the brief, not from live LLM calls.

## Run it

**Option A — open the file:** double-click `index.html`.

**Option B — local server (recommended):**

```bash
cd ~/Documents/blackboard-demo
python3 -m http.server 8777
```

Then open <http://localhost:8777/index.html>.

## Regenerate the guide (optional)

```bash
cd tools
npm install docx      # one-time
node gen_guide.js     # writes ../Build_and_Run_Guide.docx
```
