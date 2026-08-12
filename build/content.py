"""Guerrilla AI for Excel: single source of truth for all artifacts.
Dataset, day titles, hooks, demo steps, code blocks, desk actions, teases.
Every deck, workbook, cheat sheet, and copy doc pulls from here. No em dashes.
"""
import random

SERIES = "Guerrilla AI for Excel"
SUBTITLE = "Five 20-minute wins with the tools you already have"
SCHEDULE = "Monday, September 28 through Friday, October 2, 2026"
TIMESLOT = "12:00 to 12:20 PM Eastern"
FOOTER = "Live Mon Sep 28 to Fri Oct 2, 12:00 to 12:20 PM ET. Sessions are recorded."
COMPANY = "Lakeshore Office Supply"

THESIS = ("You don't need a Copilot license, a budget line, or IT's permission to work "
          "like an AI-powered analyst. Guerrilla AI means being crafty with what's already "
          "on your machine: the machine learning Excel has quietly shipped for years, the "
          "structural habits that make any AI dramatically better at reading your work, and "
          "the free tools that close the gap.")

# ---------------------------------------------------------------- dataset ----
REPS = [
    ("dana.kowalski", "Dana Kowalski"), ("marcus.webb", "Marcus Webb"),
    ("priya.raman", "Priya Raman"), ("tom.delgado", "Tom Delgado"),
    ("aisha.clark", "Aisha Clark"), ("sam.osei", "Sam Osei"),
]
EMAIL_DOMAIN = "lakeshoresupply.com"
PRODUCTS = [
    ("Copy Paper", "Office Supplies", 42.00), ("Toner Cartridge", "Office Supplies", 89.00),
    ("Standing Desk", "Furniture", 549.00), ("Task Chair", "Furniture", 219.00),
    ("Laser Printer", "Technology", 379.00), ("Conference Phone", "Technology", 269.00),
]
REGIONS = ["Midwest", "Northeast", "South", "West"]
DIRTY_REGION = {   # index -> dirty variant applied to some rows
    "Midwest": ["Midwest", "MIDWEST", "midwest "], "Northeast": ["Northeast", "northeast", "Northeast "],
    "South": ["South", "SOUTH", " South"], "West": ["West", "west", "West "],
}
RAW_COLUMNS = ["Order ID", "Order Date", "Rep Email", "Product - Category",
               "Region", "Units", "Unit Price", "Amount"]
CLEAN_COLUMNS = ["Order ID", "Order Date", "Rep Email", "Product", "Category",
                 "Region", "Units", "Unit Price", "Amount"]

def _dates_2026(rng, months, n):
    out = []
    for _ in range(n):
        m = rng.choice(months); d = rng.randint(1, 28)
        if rng.random() < 0.22: out.append(f"2026-{m:02d}-{d:02d}")       # ISO landmine
        else: out.append(f"{m}/{d}/2026")
    return out

def make_rows(seed=42, n=54, months=(1, 2, 3, 4, 5), start_id=10401):
    """Rows for the raw export, in RAW_COLUMNS order, dirty on purpose."""
    rng = random.Random(seed)
    dates = _dates_2026(rng, months, n)
    rows = []
    for i in range(n):
        rep, _ = REPS[rng.randrange(len(REPS))]
        prod, cat, price = PRODUCTS[rng.randrange(len(PRODUCTS))]
        region = rng.choice(REGIONS)
        region_val = rng.choice(DIRTY_REGION[region])
        units = rng.randint(2, 24)
        units_val = f"{units} " if rng.random() < 0.10 else units          # text units landmine
        price_val = f"${price:,.2f}" + (" " if rng.random() < 0.15 else "")  # $ text landmine
        amount = round(units * price, 2)
        rows.append([f"SO-{start_id + i}", dates[i], f"{rep}@{EMAIL_DOMAIN}",
                     f"{prod} - {cat}", region_val, units_val, price_val, amount])
    return rows

def junked(rows):
    """Insert the junk rows a real export would have: two TOTAL rows and a blank row."""
    q1 = [r for r in rows if _month(r[1]) <= 3]; rest = [r for r in rows if _month(r[1]) > 3]
    total1 = ["TOTAL Q1", "", "", "", "", "", "", round(sum(r[7] for r in q1), 2)]
    total2 = ["TOTAL Q2 TO DATE", "", "", "", "", "", "", round(sum(r[7] for r in rest), 2)]
    blank = ["", "", "", "", "", "", "", ""]
    return q1 + [total1, blank] + rest + [total2]

def _month(ds):
    return int(ds.split("-")[1]) if ds.startswith("2026-") else int(ds.split("/")[0])

def june_rows(seed=7, n=8):
    return make_rows(seed=seed, n=n, months=(6,), start_id=10496)

def clean_rows(rows):
    """What Power Query produces: split product, trimmed proper region, typed values."""
    out = []
    for r in rows:
        if not r[0] or str(r[0]).startswith("TOTAL"): continue
        prod, cat = r[3].split(" - ")
        region = str(r[4]).strip().title()
        units = int(str(r[5]).strip())
        price = float(str(r[6]).replace("$", "").replace(",", "").strip())
        out.append([r[0], iso_date(r[1]), r[2], prod, cat, region, units, price, r[7]])
    return out

def iso_date(ds):
    if ds.startswith("2026-"): return ds
    m, d, y = ds.split("/"); return f"{y}-{int(m):02d}-{int(d):02d}"

def region_summary(crows):
    tot, cnt = {}, {}
    for r in crows:
        tot[r[5]] = tot.get(r[5], 0) + r[8]; cnt[r[5]] = cnt.get(r[5], 0) + 1
    return [[reg, round(tot[reg], 2), cnt[reg]]
            for reg in sorted(tot, key=tot.get, reverse=True)]

def rep_name(email):
    local = email.split("@")[0]
    return " ".join(p.capitalize() for p in local.split("."))

# ----------------------------------------------------------- demo content ----
# Steps use short monospace-friendly lines. They appear identically in the deck's
# demo card, the workbook steps sheet, the solution notes, and the cheat sheet.

D1_FLASHFILL = [
    "# Split a clean name out of the email column",
    "1. On raw-export, click cell I1 and type: Rep Name",
    "2. In I2, type the answer for row 2: Dana Kowalski",
    "3. Press Enter, then press Ctrl+E (Flash Fill)",
    "4. Watch the whole column fill from your one example",
    "5. Spot-check two rows near the bottom",
]
D1_ANALYZE = [
    "# Ask Excel a question in plain English",
    "1. Click anywhere in the data",
    "2. Home tab > Analyze Data",
    "3. Type: total Amount by Region",
    "4. Insert the suggested PivotTable",
    "5. Fallback: Insert > Recommended PivotTables",
]
D2_TABLE = [
    "# Convert the range to a named table",
    "1. Delete the two TOTAL rows and the blank row",
    "2. Click in the data, press Ctrl+T, confirm headers",
    "3. Table Design > Table Name: type SalesData",
    "4. In a spare cell type: =SUM(SalesData[Amount])",
    "5. Read the formula out loud. It documents itself.",
]
D2_RERUN = [
    "# Same question, better answer",
    "1. Click inside the SalesData table",
    "2. Home tab > Analyze Data",
    "3. Ask again: total Amount by Region",
    "4. Compare with Monday: no junk rows, named fields",
    "5. New rows added to the table join automatically",
]
D3_TRANSFORMS = [
    "# Record the cleanup once",
    "1. Click in RawExport, then Data > From Table/Range",
    "2. Filter Order ID: uncheck blanks and TOTAL rows",
    "3. Split 'Product - Category' by delimiter ' - '",
    "4. Region: Transform > Format > Trim, then Capitalize",
    "5. Set types: Date, Whole Number, Currency",
    "6. Home > Close & Load To... a table named SalesClean",
]
D3_REFRESH = [
    "# The applause line",
    "1. Copy the June rows from the june-rows sheet",
    "2. Paste them at the bottom of RawExport",
    "3. Right-click the SalesClean table > Refresh",
    "4. Every recorded step replays on the new data",
    "5. That recurring cleanup chore is now one click",
]
D4_BLOCK1 = [
    "# Point pandas at the clean table",
    'df = xl("SalesClean[#All]", headers=True)',
    "df.describe()",
]
D4_BLOCK2 = [
    "# Total sales by region, biggest first",
    'df.groupby("Region")["Amount"].sum().sort_values(ascending=False)',
]
D4_BLOCK3 = [
    "# One chart, straight from the table",
    'totals = df.groupby("Region", as_index=False)["Amount"].sum()',
    'sns.barplot(data=totals, x="Region", y="Amount", color="#CF3338")',
]
D5_BLOB = [
    "# The blob prompt (what most people send)",
    "summarize this sales data and tell me whats important",
    "[messy paste of the whole export]",
]

def d5_structured(summary):
    lines = [
        "# Quarterly regional sales review",
        "## Role",
        "You are an FP&A analyst reviewing regional sales.",
        "## Task",
        "- Summarize performance by region",
        "- Flag any region trailing the others",
        "- Suggest two follow-up questions",
        "## Data",
        "| Region | Total Sales | Orders |",
        "| --- | ---: | ---: |",
    ]
    lines += [f"| {r} | {t:,.0f} | {c} |" for r, t, c in summary]
    lines += ["## Format", "- Three bullets, then a Markdown table"]
    return lines

D5_ROUNDTRIP = [
    "# Move tables in and out of AI tools",
    "1. Copy the RegionSummary table in Excel, paste to a chatbot",
    "2. It arrives as a table the model reads cleanly",
    "3. Ask for results 'as a Markdown table'",
    "4. Paste the reply into Excel, then Data > Text to Columns",
    "5. Delimiter: | (pipe), then trim the empty edge columns",
]

# ------------------------------------------------------------- day frames ----
DAYS = {
    1: dict(
        title="The AI already hiding in your Excel",
        hook=("Microsoft shipped machine learning inside Excel years ago and nobody told you. "
              "Flash Fill is pattern recognition. Analyze Data and Recommended PivotTables are "
              "models running against your data right now, zero licenses required. The guerrilla "
              "skill is knowing they exist and when to reach for them."),
        demos=[("Flash Fill: one example, whole column", D1_FLASHFILL),
               ("Analyze Data: ask in plain English", D1_ANALYZE)],
        desk=["Find one column you have been cleaning by hand and let Flash Fill do it: "
              "type one example, press Ctrl+E.",
              "Open a real dataset and click Home > Analyze Data. Ask it one question you "
              "actually care about.",
              "Note what it gets wrong. Bring that list tomorrow. The fix is the whole week."],
        tease="Tomorrow: the single highest-leverage prep move an analyst can make, "
              "and it costs nothing.",
        takeaway="Before asking for an AI budget, exhaust the AI you already have.",
    ),
    2: dict(
        title="Make your data AI-readable",
        hook=("Every AI tool, from Analyze Data today to Copilot and Python tomorrow, performs "
              "dramatically better on Excel tables than on raw ranges. Structured references make "
              "formulas self-documenting, so both AI and your colleagues can actually interpret "
              "your workbook. Making your data machine-readable is the highest-leverage prep move "
              "an analyst can make while everyone else waits for licenses."),
        demos=[("Ctrl+T changes everything", D2_TABLE),
               ("Rerun Monday's question", D2_RERUN)],
        desk=["Pick your most-used workbook and press Ctrl+T on the main range. Give the table "
              "a real name.",
              "Rewrite one fragile formula with structured references and read it out loud.",
              "Rerun an Analyze Data question you tried Monday and compare the answers."],
        tease="Tomorrow: the tool that delivers what people think they need AI for, "
              "and why clean inputs stop AI from lying to you.",
        takeaway="The AI-ready workbook is a craft skill, not a purchase.",
    ),
    3: dict(
        title="Clean data so AI doesn't lie to you",
        hook=("Two punches today. First, Power Query delivers the outcome people think they need "
              "AI for: automated, repeatable, refreshable data prep, with zero AI. Recorded steps "
              "are automation without programming. Second, clean inputs are what stop AI from "
              "hallucinating your numbers. Garbage in, confident-sounding garbage out. Bonus "
              "move: chatbots write M code, and Power Query is where you paste it."),
        demos=[("Record the cleanup once", D3_TRANSFORMS),
               ("Change the source, hit Refresh", D3_REFRESH)],
        desk=["Find the export you clean every week or month and rebuild the cleanup once in "
              "Power Query.",
              "Next cycle, paste the new export in and hit Refresh instead of redoing the work.",
              "Stuck on a transform? Describe it to a free chatbot and ask for the M code."],
        tease="Tomorrow: the full data science lab that has been sitting behind =PY() "
              "this whole time.",
        takeaway="One hour of Power Query replaces a recurring cleanup chore forever.",
    ),
    4: dict(
        title="A data science lab behind =PY()",
        hook=("There is a full Python runtime with pandas and seaborn sitting inside Excel. No "
              "install, no admin rights, no IT ticket. And because free chatbots write competent "
              "Python, this is where AI code generation pays off even without Copilot: describe "
              "what you want, paste the code into a PY cell, run it on your table."),
        demos=[("Descriptive stats in one cell", D4_BLOCK1 + [""] + D4_BLOCK2),
               ("A chart the chatbot wrote", D4_BLOCK3)],
        desk=["Type =PY( in any cell of a Microsoft 365 workbook and see the editor appear.",
              "Point df = xl(\"YourTable[#All]\", headers=True) at a table you own and run "
              "df.describe().",
              "Ask a free chatbot for one seaborn chart of your table and paste the code into "
              "a PY cell."],
        tease="Tomorrow: the language every AI tool speaks, and the cheapest upgrade "
              "of the week.",
        takeaway="You already have a data science environment. The craft is knowing "
                 "how to ask for the code.",
    ),
    5: dict(
        title="Markdown: the language AI speaks",
        hook=("Prompts, Copilot output, chatbot answers, agent instruction files, and half the "
              "modern web are all Markdown. Twenty minutes of syntax makes every AI tool easier "
              "to steer and its output easier to reuse. And tables round-trip between Excel and "
              "chatbots cleanly as Markdown, which is the guerrilla way to move data in and out "
              "of AI tools without connectors or licenses."),
        demos=[("Structure beats the blob", None),   # built from d5_structured at runtime
               ("The Excel round-trip", D5_ROUNDTRIP)],
        desk=["Rewrite your next AI prompt with a heading, a bulleted task list, and a Format "
              "section. Compare the answer.",
              "Paste one small Excel table into a chatbot and ask for the reply as a Markdown "
              "table.",
              "Save your best prompt as a .md file. That file is reusable craft."],
        tease=None,
        takeaway="Speaking the machine's language is the cheapest AI upgrade there is.",
    ),
}

DATASET_FLAWS = [
    "Rep appears only as an email address",
    "Product and category jammed into one column",
    "Region typed four different ways, some with stray spaces",
    "Dates in two formats, numbers stored as text, prices with $ signs",
    "Two TOTAL rows and a blank row buried in the data",
]

CTA_MEMBERSHIP = ("Replays of all five sessions live in the Stringfest Analytics membership "
                  "vault, along with the full back catalog. Ten dollars a month, cancel anytime.")
CTA_COHORT = ("Want the full-depth version? AI for Excel in a Week runs this December: five "
              "60-minute live sessions that take everything further, hands on.")
