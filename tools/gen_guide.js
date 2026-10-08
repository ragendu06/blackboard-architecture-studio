const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType,
  TableOfContents, PageBreak, ExternalHyperlink, LevelFormat, PositionalTab,
  PositionalTabAlignment, PositionalTabLeader
} = require('docx');

/* ---------- palette (matches the app) ---------- */
const TEAL = "0F8F80", TEAL_DK = "0A6A5F", INK = "141A20", INK2 = "404B57", INK3 = "6B7885";
const LINE = "DBE1E8", SOFT = "ECF6F4", SLATE = "1B2530";
const HUE = { single:"6B7885", orch:"5866C9", group:"D98A2B", board:"0F8F80", decen:"8A6FB0",
              good:"2E9E6B", warn:"D08A2C", crit:"D0574E" };

/* ---------- helpers ---------- */
const FONT = "Calibri", MONO = "Consolas", DISP = "Georgia";

function h1(text){ return new Paragraph({ heading: HeadingLevel.HEADING_1, spacing:{before:320,after:120},
  children:[new TextRun({text, font:DISP, bold:true, color:INK, size:30})] }); }
function h2(text){ return new Paragraph({ heading: HeadingLevel.HEADING_2, spacing:{before:220,after:80},
  children:[new TextRun({text, font:DISP, bold:true, color:TEAL_DK, size:24})] }); }
function h3(text){ return new Paragraph({ heading: HeadingLevel.HEADING_3, spacing:{before:160,after:60},
  children:[new TextRun({text, font:FONT, bold:true, color:INK, size:21})] }); }
function p(runs, opts={}){ return new Paragraph({ spacing:{after:opts.after??120, line:276}, alignment:opts.align,
  children: Array.isArray(runs)?runs:[new TextRun({text:runs, font:FONT, size:21, color:INK2})] }); }
function t(text, o={}){ return new TextRun({ text, font:o.mono?MONO:FONT, size:o.size??21, bold:o.bold, italics:o.italics,
  color:o.color??INK2 }); }
function bullet(runs, level=0){ return new Paragraph({ numbering:{reference:"bl", level}, spacing:{after:70, line:270},
  children: Array.isArray(runs)?runs:[new TextRun({text:runs, font:FONT, size:21, color:INK2})] }); }
function num(runs){ return new Paragraph({ numbering:{reference:"nl", level:0}, spacing:{after:90, line:272},
  children: Array.isArray(runs)?runs:[new TextRun({text:runs, font:FONT, size:21, color:INK2})] }); }
function code(text){ return new Paragraph({ spacing:{after:40, before:40, line:264}, shading:{type:ShadingType.CLEAR, fill:"F3F5F8"},
  border:{left:{style:BorderStyle.SINGLE, size:18, color:TEAL, space:8}},
  children:[new TextRun({text, font:MONO, size:19, color:"243B45"})] }); }
function spacer(sz=60){ return new Paragraph({spacing:{after:sz}, children:[new TextRun("")]}); }

function calloutBox(title, body, accent=TEAL){
  return new Table({
    width:{size:100, type:WidthType.PERCENTAGE}, columnWidths:[9360],
    borders:{ top:{style:BorderStyle.SINGLE,size:2,color:accent}, bottom:{style:BorderStyle.SINGLE,size:2,color:accent},
      left:{style:BorderStyle.SINGLE,size:24,color:accent}, right:{style:BorderStyle.SINGLE,size:2,color:accent},
      insideHorizontal:{style:BorderStyle.NONE}, insideVertical:{style:BorderStyle.NONE} },
    rows:[ new TableRow({ children:[ new TableCell({ width:{size:9360,type:WidthType.DXA},
      shading:{type:ShadingType.CLEAR, fill:"F7FAF9"}, margins:{top:120,bottom:120,left:160,right:160},
      children:[
        new Paragraph({spacing:{after:50}, children:[new TextRun({text:title, font:FONT, bold:true, size:20, color:accent})]}),
        ...body.map(line=> new Paragraph({spacing:{after:40, line:268}, children:[new TextRun({text:line, font:FONT, size:20, color:INK2})]}))
      ] }) ] }) ]
  });
}

/* ---------- reference table (2-col: element / what it does) ---------- */
function refTable(rows){
  const mk = (txt, header, mono) => new TableCell({
    width:{size: header?2600:6760, type:WidthType.DXA},
    shading:{type:ShadingType.CLEAR, fill: header?SLATE:"FFFFFF"},
    margins:{top:70,bottom:70,left:120,right:120},
    children:[new Paragraph({children:[new TextRun({text:txt, font: header?FONT:(mono?FONT:FONT),
      bold:header, size:header?18:19, color: header?"FFFFFF":INK2})]})]
  });
  const headerRow = new TableRow({ tableHeader:true, children:[
    new TableCell({width:{size:2600,type:WidthType.DXA}, shading:{type:ShadingType.CLEAR,fill:SLATE}, margins:{top:80,bottom:80,left:120,right:120},
      children:[new Paragraph({children:[new TextRun({text:"UI element", font:FONT, bold:true, size:18, color:"FFFFFF"})]})]}),
    new TableCell({width:{size:6760,type:WidthType.DXA}, shading:{type:ShadingType.CLEAR,fill:SLATE}, margins:{top:80,bottom:80,left:120,right:120},
      children:[new Paragraph({children:[new TextRun({text:"What it does", font:FONT, bold:true, size:18, color:"FFFFFF"})]})]})
  ]});
  const bodyRows = rows.map((r,i)=> new TableRow({ children:[
    new TableCell({width:{size:2600,type:WidthType.DXA}, shading:{type:ShadingType.CLEAR, fill: i%2?"F6F8FA":"FFFFFF"},
      margins:{top:70,bottom:70,left:120,right:120},
      children:[new Paragraph({children:[new TextRun({text:r[0], font:FONT, bold:true, size:19, color:INK})]})]}),
    new TableCell({width:{size:6760,type:WidthType.DXA}, shading:{type:ShadingType.CLEAR, fill: i%2?"F6F8FA":"FFFFFF"},
      margins:{top:70,bottom:70,left:120,right:120},
      children:[new Paragraph({spacing:{line:264}, children: Array.isArray(r[1])?r[1]:[new TextRun({text:r[1], font:FONT, size:19, color:INK2})]})]})
  ]}));
  return new Table({ width:{size:100,type:WidthType.PERCENTAGE}, columnWidths:[2600,6760],
    borders:{ top:{style:BorderStyle.SINGLE,size:2,color:LINE}, bottom:{style:BorderStyle.SINGLE,size:2,color:LINE},
      left:{style:BorderStyle.SINGLE,size:2,color:LINE}, right:{style:BorderStyle.SINGLE,size:2,color:LINE},
      insideHorizontal:{style:BorderStyle.SINGLE,size:2,color:LINE}, insideVertical:{style:BorderStyle.SINGLE,size:2,color:LINE} },
    rows:[headerRow, ...bodyRows] });
}

/* comparison table for topologies */
function topoTable(){
  const head = ["Topology","Coordination","Token cost","Accuracy on hard tasks"];
  const data = [
    ["Single Agent","One chain of thought, no peers","Lowest calls","Weakest — no second opinion", HUE.single],
    ["Orchestrator–Worker","Manager routes a fixed SOP pipeline","High","Strong when workflow is known", HUE.orch],
    ["Group Chat","All agents broadcast until they converge","Highest","Good, but noisy and redundant", HUE.group],
    ["Blackboard ★","Controller picks the next knowledge source; entries carry credibility","Low (competitive)","Top-tier, incl. ill-structured", HUE.board],
    ["Decentralized Queue","Agents pull tasks and write verified updates","Lowest compute","Good; slightly below blackboard", HUE.decen],
  ];
  const cw=[1900,3200,1500,2760];
  const hr = new TableRow({tableHeader:true, children: head.map((hd,i)=> new TableCell({
    width:{size:cw[i],type:WidthType.DXA}, shading:{type:ShadingType.CLEAR,fill:SLATE}, margins:{top:70,bottom:70,left:110,right:110},
    children:[new Paragraph({children:[new TextRun({text:hd, font:FONT, bold:true, size:17, color:"FFFFFF"})]})]})) });
  const rows = data.map((r,i)=> new TableRow({ children: r.slice(0,4).map((cell,ci)=> new TableCell({
    width:{size:cw[ci],type:WidthType.DXA}, shading:{type:ShadingType.CLEAR, fill: r[0].includes("★")?SOFT:(i%2?"F6F8FA":"FFFFFF")},
    margins:{top:60,bottom:60,left:110,right:110},
    children:[new Paragraph({spacing:{line:258}, children:[ ci===0
      ? new TextRun({text:cell, font:FONT, bold:true, size:18, color:"#"===cell[0]?INK:r[4]})
      : new TextRun({text:cell, font:FONT, size:18, color:INK2})]})]
  })) }));
  return new Table({ width:{size:100,type:WidthType.PERCENTAGE}, columnWidths:cw,
    borders:{ top:{style:BorderStyle.SINGLE,size:2,color:LINE}, bottom:{style:BorderStyle.SINGLE,size:2,color:LINE},
      left:{style:BorderStyle.SINGLE,size:2,color:LINE}, right:{style:BorderStyle.SINGLE,size:2,color:LINE},
      insideHorizontal:{style:BorderStyle.SINGLE,size:2,color:LINE}, insideVertical:{style:BorderStyle.SINGLE,size:2,color:LINE} },
    rows:[hr, ...rows] });
}

/* variants table for Tool 02 (Variant / Origin / How control decides / Tokens) */
function variantsTable(){
  const head=["Variant","Origin in the brief","How control decides who acts next","Tokens to solve"];
  const data=[
    ["Classical Agenda","Hearsay-II · Nii","Fixed-priority agenda — runs the highest-ranked triggered source, no value reasoning","2,450", HUE.single],
    ["Control Blackboard","Hayes-Roth 1985","A separate CONTROL lane cycles the focus EXPAND → VERIFY → CONCLUDE","2,250", HUE.orch],
    ["LLM Router","bMAS · Han & Zhang 2025","An LLM reads the board and names the next source; no token budget","2,600", HUE.group],
    ["Volunteer Market","Salemi et al. 2025","Every eligible source bids a self-fit score; the highest bidder acts","1,800", HUE.single],
    ["Cost-Aware Controller  ★","Direction A · this paper","Scores each source by expected value ÷ token cost, picks the best, stops early","1,600", HUE.board],
    ["Credibility-Gated","Direction B · PatchBoard","Reads gated at cred ≥ 0.70 — a low-trust claim is blocked until a Verifier raises it","1,800", HUE.good],
    ["Secure Partitioned","Direction C · Terrarium","Trusted / untrusted regions; a firewall blocks poisoned entries from privileged sources","1,900", HUE.crit],
  ];
  const cw=[1900,1900,4060,1160];
  const hr=new TableRow({tableHeader:true, children:head.map((hd,i)=> new TableCell({
    width:{size:cw[i],type:WidthType.DXA}, shading:{type:ShadingType.CLEAR,fill:SLATE}, margins:{top:70,bottom:70,left:100,right:100},
    children:[new Paragraph({children:[new TextRun({text:hd, font:FONT, bold:true, size:16, color:"FFFFFF"})]})]})) });
  const rows=data.map((r,i)=> new TableRow({ children:[0,1,2,3].map(ci=> new TableCell({
    width:{size:cw[ci],type:WidthType.DXA}, shading:{type:ShadingType.CLEAR, fill:r[0].includes("★")?SOFT:(i%2?"F6F8FA":"FFFFFF")},
    margins:{top:60,bottom:60,left:100,right:100},
    children:[new Paragraph({spacing:{line:252}, children:[ ci===0
      ? new TextRun({text:r[0], font:FONT, bold:true, size:17, color:r[4]})
      : new TextRun({text:r[ci], font: ci===3?MONO:FONT, size:ci===3?16:17, color:INK2, bold:ci===3})]})]
  })) }));
  return new Table({ width:{size:100,type:WidthType.PERCENTAGE}, columnWidths:cw,
    borders:{ top:{style:BorderStyle.SINGLE,size:2,color:LINE}, bottom:{style:BorderStyle.SINGLE,size:2,color:LINE},
      left:{style:BorderStyle.SINGLE,size:2,color:LINE}, right:{style:BorderStyle.SINGLE,size:2,color:LINE},
      insideHorizontal:{style:BorderStyle.SINGLE,size:2,color:LINE}, insideVertical:{style:BorderStyle.SINGLE,size:2,color:LINE} },
    rows:[hr,...rows] });
}

/* ---------- build document ---------- */
const children = [];

/* cover */
children.push(new Paragraph({spacing:{before:600, after:0}, children:[
  new TextRun({text:"RESEARCH DEMO · BUILD & RUN GUIDE", font:MONO, size:18, color:TEAL_DK, bold:true})]}));
children.push(new Paragraph({spacing:{before:60, after:0}, border:{bottom:{style:BorderStyle.SINGLE,size:12,color:TEAL,space:10}},
  children:[new TextRun({text:"The Blackboard Architecture Studio", font:DISP, bold:true, size:48, color:INK})]}));
children.push(new Paragraph({spacing:{before:200, after:60}, children:[
  new TextRun({text:"One page, two linked interactive tools. Tool 01 compares five LLM multi-agent coordination topologies with a live Welch t-test; Tool 02 opens the blackboard itself and runs seven variations of its control loop as a stepped animation.",
    font:FONT, size:21, color:INK2, italics:true})]}));
children.push(new Paragraph({spacing:{before:120}, children:[
  t("Prepared for Faizan  ·  ", {size:19, color:INK3}),
  t("Companion to Research Brief 1 — Blackboard Architecture for LLM Multi-Agent Systems", {size:19, color:INK3})]}));
children.push(spacer(120));
children.push(calloutBox("What you're holding", [
  "• index.html — the whole studio in one self-contained file (no build step, no dependencies, no internet needed). A segmented control at the top switches between Tool 01 (Topology Comparison) and Tool 02 (Blackboard Variants).",
  "• This guide — how to run it, and exactly what every button, slider, chart, meter and number on both tools does.",
  "Both tools are mocks: the numbers come from synthetic, seeded models calibrated to the effect sizes reported in the brief, not from live LLM calls. They exist to demonstrate and teach the architecture visually, not to benchmark real models."
], TEAL));

children.push(spacer(80));
children.push(h1("Contents"));
children.push(new TableOfContents("Contents", { hyperlink:true, headingStyleRange:"1-2" }));
children.push(new Paragraph({children:[new PageBreak()]}));

/* 1. what the app is */
children.push(h1("1. What the application demonstrates"));
children.push(p("The blackboard architecture is a 1970s AI pattern — independent specialists (knowledge sources) cooperate by reading from and writing to a shared workspace, while a control component decides who acts next. Since 2025 it has returned as a coordination model for LLM agents. Your brief's central, testable claim is that this pattern earns competitive accuracy while spending fewer tokens than the busier alternatives."));
children.push(p("This studio makes that claim tangible across two tools, reached by the segmented switcher at the top of the page:"));
children.push(bullet([t("Tool 01 — Topology Comparison. ",{bold:true,color:INK,size:21}), t("Places all five coordination topologies side by side, animates how each moves information, then runs seeded Monte-Carlo trials so you can watch the statistics — including a Welch t-test — settle in real time. Covered in Section 3.",{size:21})]));
children.push(bullet([t("Tool 02 — Blackboard Variants. ",{bold:true,color:INK,size:21}), t("Zooms inside the blackboard and runs seven different control policies as a live, stepped animation, showing how each decides who acts next and what that costs. Covered in Section 4.",{size:21})]));
children.push(p("The page opens on Tool 01. Everything below applies to both; where it matters, the tool is named."));
children.push(h2("The five topologies (Tool 01)"));
children.push(topoTable());
children.push(spacer(60));
children.push(p([t("The blackboard (marked ", {}), t("★", {bold:true, color:TEAL_DK}),
  t(") is the hypothesis under test. Everything in the interface is arranged so you can challenge it: switch approaches, make the task harder, add seeds, and see whether the blackboard's advantage holds up or dissolves into noise.", {})]));

/* 2. build & run */
children.push(h1("2. Build & run"));
children.push(p([t("Good news first: there is nothing to compile. The application is one HTML file with all styles and logic inline, so \"building\" it means, at most, opening it. Three ways to run it, easiest first.", {})]));

children.push(h2("Option A — Just open the file (fastest)"));
children.push(num("Locate the project folder at ~/Documents/blackboard-demo."));
children.push(num([t("Double-click ", {}), t("index.html", {mono:true, color:INK}), t(". It opens in your default browser and runs immediately.", {})]));
children.push(p([t("This works because the app needs no server and no network. Use this for a quick look or to hand it to someone.", {italics:true, color:INK3, size:20})]));

children.push(h2("Option B — Run a tiny local server (recommended)"));
children.push(p([t("A local server guarantees fonts and character encoding load exactly as designed. Open Terminal and run:", {})]));
children.push(code("cd ~/Documents/blackboard-demo"));
children.push(code("python3 -m http.server 8777"));
children.push(p([t("Then visit ", {}), t("http://localhost:8777/index.html", {mono:true, color:TEAL_DK}),
  t(" in your browser. Press ", {}), t("Ctrl+C", {mono:true, color:INK}), t(" in Terminal to stop it when you're done.", {})]));

children.push(h2("Option C — The hosted version"));
children.push(p([t("A published, shareable copy already lives here:", {})]));
children.push(new Paragraph({spacing:{after:120}, children:[ new ExternalHyperlink({
  link:"https://claude.ai/code/artifact/27769eb8-9b2f-4eb1-921b-84324c38d864",
  children:[new TextRun({text:"claude.ai/code/artifact/27769eb8-9b2f-4eb1-921b-84324c38d864", font:MONO, size:19, color:TEAL_DK, underline:{}})]})]}));
children.push(p([t("Opening it may ask you to sign in to Claude; the local file (Options A/B) never does.", {italics:true, color:INK3, size:20})]));

children.push(calloutBox("Requirements", [
  "• Any modern browser — Chrome, Safari, Edge, or Firefox from the last few years.",
  "• Python 3 only for Option B (macOS already has it; check with  python3 --version).",
  "• No npm, no node_modules, no API keys, no internet connection for the local file."
], HUE.warn));

/* 3. the interface tour */
children.push(new Paragraph({children:[new PageBreak()]}));
children.push(h1("3. Tool 01 — Topology Comparison, element by element"));
children.push(p([t("At the very top of the page sits the ",{}), t("tool switcher",{bold:true,color:INK}), t(" — a segmented control with ",{}), t("01 Topology Comparison",{bold:true,color:INK}), t(" and ",{}), t("02 Blackboard Variants",{bold:true,color:INK}), t(". It selects which tool fills the page; the page opens on Tool 01. This section documents Tool 01, whose screen has four zones, top to bottom: the header with global controls, the topology rail, the board stage, and the analytics section.",{})]));

children.push(h2("3.1 Header & global controls"));
children.push(refTable([
  ["Task difficulty slider", [t("Sets how hard the shared task is, from 0.00 (easy) to 1.00 (ill-structured). Higher difficulty lowers every topology's accuracy and raises its token use — and it ", {size:19,color:INK2}), t("widens the blackboard's advantage", {size:19,bold:true,color:INK}), t(", because opportunistic control helps most when no fixed pipeline fits. Releasing the slider re-runs all trials.", {size:19,color:INK2})]],
  ["Seeds / trials slider", "Sets how many independent runs (seeds) each topology gets, from 8 to 120. More seeds shrink the confidence intervals and make the t-test more decisive. This is the statistical-power dial."],
  ["Run trials button", "Re-executes the whole Monte-Carlo experiment for all five topologies at the current settings and redraws every chart, the t-test, and the leaderboard. Use it after changing settings, or just to confirm results are stable."],
  ["Theme button", "Toggles light / dark colour theme. The slate board keeps its dark surface in both; the surrounding chrome adapts."],
]));

children.push(h2("3.2 Approach rail (left column)"));
children.push(refTable([
  ["Topology cards", [t("Five clickable cards, one per approach, each colour-coded (grey, indigo, amber, ", {size:19,color:INK2}), t("teal", {size:19,bold:true,color:TEAL_DK}), t(", violet). Clicking one selects it: the board stage, the advantages/disadvantages lists, and the live metrics all switch to that topology and replay its trace.", {size:19,color:INK2})]],
  ["★ badge", "Marks the Blackboard card as \"this paper\" — the approach the whole demo is built to test."],
  ["Card sub-label", "The one-line tagline (e.g. \"Opportunistic + credibility\") and a provenance tag (\"MetaGPT-style\", \"DeLM-style\") tying the mock to the real system in your brief."],
]));

children.push(h2("3.3 Board stage (the visualization)"));
children.push(p([t("This is the heart of the demo — a dark \"slate board\" that shows the selected approach as a live diagram plus an animated trace of what the agents post.", {})]));
children.push(refTable([
  ["Live metric tiles", "Four counters — Accuracy, Tokens, Latency, LLM calls — that tick upward as the trace animates and settle on the first seed's sampled values. They give you an at-a-glance feel for the approach's cost and quality."],
  ["Diagram panel (left)", [t("An SVG schematic of the topology's shape: a lone agent, a star of workers around an orchestrator, a full mesh for group chat, a queue for the decentralized design, and — for the blackboard — the shared board with credibility-rated hypotheses at several levels plus a controller choosing the next knowledge source.", {size:19,color:INK2})]],
  ["Trace panel (right)", "A step-by-step log that animates one entry at a time, showing who acted and what they posted. It illustrates the coordination pattern in motion — a fixed SOP for the orchestrator, redundant broadcasts for group chat, opportunistic controller-gated picks for the blackboard."],
  ["Credibility chips & bars", [t("On the blackboard trace only, entries carry a ", {size:19,color:INK2}), t("cred 0.00–1.00", {size:19,mono:true,color:INK}), t(" score with a coloured bar — green when a verifier confirms an entry, red when one is doubted or retracted. This is the Hearsay-II idea your brief highlights: every hypothesis is rated, so errors can be damped before they spread.", {size:19,color:INK2})]],
  ["Legend", "Explains the colours used in the current diagram (controller, knowledge source, verified, doubted, etc.)."],
  ["Replay trace button", "Re-runs the animation for the selected topology without changing any numbers."],
]));

children.push(h2("3.4 Advantages & disadvantages"));
children.push(refTable([
  ["Advantages column (green +)", "The strengths of the selected approach, drawn from the classical blackboard literature and the 2025–26 LLM papers in your brief."],
  ["Disadvantages column (red –)", "The honest costs — for the blackboard, that means the need for a good control policy and the shared-state attack surface (your brief's security direction). The lists update the moment you switch approaches."],
]));

children.push(h2("3.5 Statistics section"));
children.push(refTable([
  ["Accuracy vs. token cost (Pareto)", [t("A scatter of all five topologies: accuracy up the y-axis, mean tokens along the x-axis. ", {size:19,color:INK2}), t("Up and to the left is better", {size:19,bold:true,color:INK}), t(" — high accuracy for few tokens. The dashed line is the Pareto frontier (the non-dominated set); the blackboard is drawn as a haloed teal dot so you can see whether it sits on that frontier.", {size:19,color:INK2})]],
  ["Mean tokens bar chart", "One bar per topology showing mean tokens per task, with a whisker for the 95% confidence interval. It's the visual companion to the t-test: if the blackboard's bar and whisker sit clearly below a baseline's, the difference is likely real."],
  ["Baseline selector", [t("A dropdown that chooses which approach the blackboard is tested against. Default is Group Chat (the most token-hungry). Switch it to Single Agent or Decentralized Queue to stress-test the claim against a cheaper rival.", {size:19,color:INK2})]],
  ["Welch t-test verdict card", "The teal panel states the outcome in plain language — \"Hypothesis supported\", \"Cheaper, but not conclusive\", or \"Not cheaper here\" — plus the token saving and the p-value. It turns amber when the mean is lower but not significant, so you can see the difference between a real effect and noise."],
  ["t-test numbers", [t("The full statistics: both means, the mean difference, the ", {size:19,color:INK2}), t("t statistic", {size:19,mono:true,color:INK}), t(", Welch–Satterthwaite degrees of freedom, the one-sided ", {size:19,color:INK2}), t("p-value", {size:19,mono:true,color:INK}), t(", and ", {size:19,color:INK2}), t("Cohen's d", {size:19,mono:true,color:INK}), t(" (the standardized effect size, labelled small/medium/large).", {size:19,color:INK2})]],
  ["Accuracy check note", "A second, two-sided test on accuracy. This is what lets the demo claim \"competitive accuracy\": if the accuracy gap is not significant, the blackboard matches quality while spending less — exactly the brief's argument."],
  ["Leaderboard", [t("All five topologies in one table — accuracy, tokens, latency, LLM calls, and a cost-adjusted score (accuracy ÷ tokens-per-1k). The best value in each column gets a ", {size:19,color:INK2}), t("★", {size:19,color:HUE.good,bold:true}), t("; the blackboard row is tinted teal.", {size:19,color:INK2})]],
]));

/* 4. Tool 02 — variants control room */
children.push(new Paragraph({children:[new PageBreak()]}));
children.push(h1("4. Tool 02 — the Blackboard Variants control room"));
children.push(p([t("Click ",{}), t("02 Blackboard Variants",{bold:true,color:INK}), t(" in the top switcher to open the second tool. Tool 01 asked whether to use a blackboard at all; Tool 02 assumes you have chosen one and asks the question your brief flags as under-studied: ",{}), t("once you have a shared workspace, how should control decide who acts next?",{italics:true,color:INK}), t(" There are many answers. This tool runs seven of them — the classical schemes, the modern LLM ones, and the three novel directions — as a live, stepped animation of the control loop.",{})]));

children.push(h2("The seven variations"));
children.push(p("Every variant solves the same task using the same six knowledge sources (Retriever, Reasoner, Solver, Verifier, Critic, Concluder). They differ only in the control policy — the shaded row is the paper's own proposal."));
children.push(variantsTable());
children.push(spacer(50));
children.push(p([t("The tokens-to-solve figures are the tool's own output at the default settings; because every variant reaches the same solved goal, they are a clean comparison of efficiency. Note the ordering: the two schemes that do not reason about cost — the ",{}), t("Classical Agenda",{bold:true,color:INK}), t(" and the ",{}), t("LLM Router",{bold:true,color:INK}), t(" — are the most expensive, while the learned ",{}), t("Cost-Aware Controller",{bold:true,color:INK}), t(" is the cheapest.",{})]));

children.push(h2("How a run animates"));
children.push(p("Open a variant (or press Play) and each control cycle plays in two visible phases, so you can follow the reasoning rather than just the result:"));
children.push(num([t("Decide — ",{bold:true,color:INK}), t("the controller reads the board and highlights the knowledge source it will run next. The Controller-decision panel shows its actual rationale: a priority rank (Classical), a strategy focus (Control Blackboard), a set of bids (Volunteer), or value÷cost figures (Cost-Aware).",{})]));
children.push(num([t("Fire — ",{bold:true,color:INK}), t("the chosen source runs. A new entry animates onto the board, its credibility bar fills, the token meter advances, and the event log gains a line.",{})]));
children.push(p([t("The loop repeats until the goal entry passes the credibility threshold (",{}), t("solved",{bold:true,color:TEAL_DK}), t(") or the token budget is exhausted (",{}), t("stopped",{bold:true,color:"D0574E"}), t("). The cost-aware controller stops the instant the goal is credible enough — that early stop, plus skipping low-value actions, is where its token saving comes from.",{})]));

children.push(h2("4.1 Control-room controls & readouts"));
children.push(refTable([
  ["Tool switcher (01 / 02)","The segmented control at the very top of the page. It swaps the whole page between Tool 01 and Tool 02; your place in each is kept as you flip between them."],
  ["Variant tabs (1–7)","Pick which control policy runs. Selecting one switches the board, the explainer card, the meters and the animation to that variant and replays it from an empty board."],
  ["Variant explainer",[t("The card under the tabs: what the selected variant is, its origin citation (e.g. ",{size:19,color:INK2}), t("Hayes-Roth 1985",{size:19,italics:true,color:INK}), t("), and a ",{size:19,color:INK2}), t("“what to watch”",{size:19,color:INK}), t(" hint pointing at the behaviour that makes it distinct.",{size:19,color:INK2})]],
  ["Meter row",[t("Five live readouts: ",{size:19,color:INK2}), t("Cycle",{size:19,bold:true,color:INK}), t(", ",{size:19,color:INK2}), t("Tokens / budget",{size:19,bold:true,color:INK}), t(" (the bar turns amber then red as spend nears the 4,200-token budget), ",{size:19,color:INK2}), t("Open goals",{size:19,bold:true,color:INK}), t(", ",{size:19,color:INK2}), t("Goal credibility",{size:19,bold:true,color:INK}), t(", and ",{size:19,color:INK2}), t("Status",{size:19,bold:true,color:INK}), t(" (running / solved / stopped).",{size:19,color:INK2})]],
  ["Play / Pause","Runs the animation continuously, or halts it so the board holds still."],
  ["Step","Advances exactly one control cycle. The best way to inspect a single decision — press it repeatedly to walk the loop by hand."],
  ["Restart","Clears the board back to empty for the current variant so you can run it again from the start."],
  ["Speed slider","Scales the animation from 0.5× to 3×. It changes only the pace, never the outcome."],
  ["The blackboard (lanes)",[t("The shared state, drawn as lanes by level of abstraction — ",{size:19,color:INK2}), t("GOAL, SUB-CLAIM, EVIDENCE",{size:19,mono:true,color:INK}), t(" — plus a purple ",{size:19,color:INK2}), t("CONTROL",{size:19,mono:true,color:INK}), t(" lane for the Control Blackboard and a red ",{size:19,color:INK2}), t("UNTRUSTED",{size:19,mono:true,color:INK}), t(" lane for the Secure Partitioned variant. Each entry animates in with a credibility bar and a status badge (posted / verified / doubted).",{size:19,color:INK2})]],
  ["Knowledge sources panel",[t("The six sources with their token cost. Triggered sources (preconditions met) glow; the one the controller selects pulses and shows ",{size:19,color:INK2}), t("“▶ firing”",{size:19,color:INK}), t(". In Volunteer and Cost-Aware modes each source also shows its bid or value÷cost score, so you can see the decision being made.",{size:19,color:INK2})]],
  ["Controller decision","Names the source chosen this cycle and the reasoning behind it. This is the single most important panel for seeing how the variants differ — the same board produces a different pick under each policy."],
  ["Event log","A running record of what each source actually did to the board, newest at the bottom."],
  ["Result banner","On finish, states the cycles used, tokens spent, percentage under budget and final goal credibility — or, if it stopped, why."],
  ["Cost-to-solve chart","One bar per variant, which fills in once you have run that variant. Lower is better; this is where the cost-aware advantage becomes visible at a glance."],
]));

children.push(h2("4.2 What the run proves"));
children.push(p([t("Run each variant to the end and the ",{}), t("Cost to solve",{bold:true,color:INK}), t(" chart fills in. Because all seven reach the same solved goal at the same credibility, the bar lengths isolate one thing — token efficiency — and they line up exactly as your brief predicts.",{})]));
children.push(calloutBox("The headline", [
  "The learned, cost-aware controller reaches the same answer for roughly 38% fewer tokens than the LLM router (1,600 vs 2,600) and about 35% fewer than the classical agenda. That is the concrete demonstration of Direction A: revive Hayes-Roth's idea of control as a first-class problem, but make the control decision learned and budget-aware — competitive quality, fewer tokens.",
  "The Credibility-Gated and Secure Partitioned variants make the point for Directions B and C: a small, visible extra cost (a forced verification, a firewalled entry) buys error-containment and security that the naive schemes simply do not have."
], TEAL));

/* 5. how the stats work */
children.push(new Paragraph({children:[new PageBreak()]}));
children.push(h1("5. How Tool 01 tests the hypothesis"));
children.push(p([t("The demo doesn't just assert the blackboard is better — it runs an experiment and applies a standard test. Here is the chain of reasoning, so you can defend every number on screen.", {})]));
children.push(h3("Step 1 — Generate seeded samples"));
children.push(p([t("For each topology, the app draws ", {}), t("N", {mono:true,color:INK}),
  t(" samples (N = the seeds slider) from a normal distribution whose mean and spread depend on the current difficulty. A seeded random generator (mulberry32) makes every run reproducible — the same settings always give the same numbers.", {})]));
children.push(h3("Step 2 — Summarise"));
children.push(p("Means, standard deviations and 95% confidence intervals are computed per topology. These feed the bars, whiskers and Pareto points."));
children.push(h3("Step 3 — Welch's two-sample t-test"));
children.push(p([t("To compare the blackboard's token cost against the chosen baseline, the app uses ", {}),
  t("Welch's t-test", {bold:true, color:INK}),
  t(" — the version that does not assume the two groups share a variance, which is the safe default. It computes the t statistic, the Welch–Satterthwaite degrees of freedom, and a p-value from the Student-t distribution (via the regularized incomplete beta function — a real implementation, not a lookup table).", {})]));
children.push(p([t("The test is ", {}), t("one-sided", {bold:true, color:INK}),
  t(": the hypothesis is specifically that the blackboard spends ", {}), t("fewer", {italics:true, color:INK}),
  t(" tokens, not merely a different number. If p < 0.05 and the blackboard's mean is lower, the demo reports the hypothesis as supported and rejects the null hypothesis (\"no difference\").", {})]));
children.push(h3("Step 4 — Guard the accuracy claim"));
children.push(p("A separate two-sided t-test checks accuracy. The blackboard's selling point is not \"more accurate at any cost\" but \"same accuracy, fewer tokens\". So the note reads the accuracy test and tells you whether quality was held (difference not significant) or traded (difference significant) — keeping the demo honest."));
children.push(calloutBox("Reading the effect size", [
  "The p-value tells you whether a difference is real; Cohen's d tells you whether it's big. d ≈ 0.2 is small, 0.5 medium, 0.8 large. With enough seeds you'll typically see a large-to-very-large d on tokens — the blackboard's saving is not just statistically detectable, it's substantial.",
], TEAL));

/* 5. suggested walkthrough */
children.push(h1("6. A guided walkthrough"));
children.push(h2("Tool 01 — Topology Comparison"));
children.push(num("Open the page. It loads on Tool 01, already populated, with the Blackboard selected and its trace animating."));
children.push(num("Click each topology card in turn. Watch the diagram and trace change — note how Group Chat floods the board while the Blackboard's controller picks only what it needs."));
children.push(num("Read the advantages/disadvantages for the Blackboard, then for Group Chat, to feel the trade-off."));
children.push(num("Drag Task difficulty up to ~0.85 and release. Watch every accuracy fall — and the blackboard's token lead grow."));
children.push(num("In the t-test, leave the baseline on Group Chat: you should see \"Hypothesis supported\", a large token saving, and p < 0.001."));
children.push(num("Switch the baseline to Decentralized Queue — the cheapest rival. Now the verdict may soften, which is the honest result: against a lean design the saving is smaller. This is exactly the nuance a reviewer will look for."));
children.push(num("Drag Seeds down to 8 and re-run. Confidence intervals widen and borderline results may flip to \"not conclusive\" — a live lesson in statistical power."));
children.push(h2("Tool 02 — Blackboard Variants"));
children.push(num("Click 02 Blackboard Variants in the top switcher. The Cost-Aware Controller runs automatically; watch the value÷cost figures and the early stop."));
children.push(num("Click the Classical Agenda tab and let it run. Notice it spends more tokens — it verifies and critiques by rote. Compare its Cost-to-solve bar with the cost-aware one."));
children.push(num("Select Credibility-Gated and use Step to walk it cycle by cycle. Watch a below-0.70 sub-claim get blocked from the goal until a Verifier raises it."));
children.push(num("Select Secure Partitioned and Step through it. Watch the poisoned entry appear in the red UNTRUSTED lane and get firewalled before it reaches the Concluder."));
children.push(num("Play the remaining variants so every Cost-to-solve bar fills in — the cost-aware bar is clearly the shortest."));

/* 7. customise */
children.push(h1("7. Extending the studio"));
children.push(p([t("Everything lives in ", {}), t("index.html", {mono:true, color:INK}),
  t(". The model is a single array named ", {}), t("TOPOS", {mono:true, color:INK}),
  t(" near the top of the script; each entry defines a topology's accuracy and token functions of difficulty, its pros/cons, its diagram legend and its animated trace. To tune the story, adjust these:", {})]));
children.push(bullet([t("acc:d=>…", {mono:true, color:INK, size:20}), t("  — mean accuracy as a function of difficulty d.", {size:20})]));
children.push(bullet([t("tok:d=>…", {mono:true, color:INK, size:20}), t("  — mean token cost as a function of difficulty.", {size:20})]));
children.push(bullet([t("traceScript()", {mono:true, color:INK, size:20}), t("  — the scripted steps shown in the animated trace.", {size:20})]));
children.push(p([t("To swap in real data, replace the sampling loop in ", {}), t("runExperiment()", {mono:true, color:INK}),
  t(" so it reads your measured accuracy and token arrays instead of drawing from the synthetic distributions — the charts and t-test will work unchanged, because they only ever see arrays of numbers.", {})]));
children.push(p([t("Tool 02 is structured the same way. Its seven policies live in a ", {}), t("VARIANTS", {mono:true, color:INK}),
  t(" array and a ", {}), t("decide()", {mono:true, color:INK}), t(" function whose branches are the control policies; the shared knowledge sources and their costs sit in a ", {}),
  t("KS", {mono:true, color:INK}), t(" object. To add an eighth variation, add an entry to ", {}), t("VARIANTS", {mono:true, color:INK}),
  t(" and a matching branch in ", {}), t("decide()", {mono:true, color:INK}), t(" — the board, meters, animation and cost chart pick it up automatically.", {})]));

/* 8. troubleshooting */
children.push(h1("8. Troubleshooting"));
children.push(refTable([
  ["Odd characters (Â·, â†’)", "The file is UTF-8. If you open it from an old server that forces a different encoding, run it via Option B instead — python's server and the file's own charset tag keep it correct."],
  ["Fonts look plain", "The display fonts load from Google Fonts, which needs internet. Offline, the app falls back to Georgia / system fonts by design — the layout and all functionality are unaffected."],
  ["Charts look empty", "Click Run trials once. If still empty, your browser may be very old; try Chrome or Safari."],
  ["Hosted link asks me to sign in", "That's expected for the claude.ai copy. Use the local index.html (Options A/B) for a no-login version."],
  ["Variants animation looks finished", "Tool 02 runs fast — a variant can solve in a few seconds. Press Restart then Step to walk it one cycle at a time, or lower the Speed slider before pressing Play."],
]));

children.push(spacer(120));
children.push(new Paragraph({border:{top:{style:BorderStyle.SINGLE,size:8,color:LINE,space:8}},
  spacing:{before:120}, children:[
  new TextRun({text:"Blackboard Architecture Studio — build & run guide  ·  companion to Research Brief 1  ·  the studio and this guide were generated together, so every element name above matches the screen exactly.",
    font:FONT, size:17, color:INK3, italics:true})]}));

/* ---------- assemble ---------- */
const doc = new Document({
  creator:"Personal Assistant", title:"Blackboard Architecture Studio — Build & Run Guide",
  styles:{ default:{ document:{ run:{ font:FONT, size:21, color:INK } } } },
  numbering:{ config:[
    { reference:"bl", levels:[{level:0, format:LevelFormat.BULLET, text:"•", alignment:AlignmentType.LEFT,
      style:{run:{color:TEAL_DK}, paragraph:{indent:{left:420, hanging:240}}}}] },
    { reference:"nl", levels:[{level:0, format:LevelFormat.DECIMAL, text:"%1.", alignment:AlignmentType.LEFT,
      style:{run:{color:TEAL_DK, bold:true}, paragraph:{indent:{left:440, hanging:280}}}}] },
  ]},
  sections:[{
    properties:{ page:{ size:{width:12240, height:15840}, margin:{top:1180, bottom:1080, left:1200, right:1200} } },
    children
  }]
});

Packer.toBuffer(doc).then(buf=>{
  fs.writeFileSync(require("path").join(__dirname, "..", "Build_and_Run_Guide.docx"), buf);
  console.log("WROTE", buf.length, "bytes");
});
