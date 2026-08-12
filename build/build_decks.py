"""Build the five Guerrilla AI decks with verbatim speaker notes, from content.py."""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import content as C
import deckbuilder as D
from pptx.util import Inches, Pt

OUT = sys.argv[1] if len(sys.argv) > 1 else "out"

def wordmark(slide):
    """Styled-text brand mark, since no logo PNG assets are installed."""
    _, tf = D._box(slide, Inches(0.5), Inches(0.42), Inches(4.5), Inches(0.4))
    D._run(tf.paragraphs[0], "STRINGFEST ANALYTICS", 13, D.RED, D.HEAD, bold=True)

summary = C.region_summary(C.clean_rows(C.junked(C.make_rows())) + C.clean_rows(C.june_rows()))
STRUCTURED = C.d5_structured(summary)

THESIS_SPOKEN = (
    "Here is the idea behind this whole series. You do not need a Copilot license, a budget "
    "line, or IT's permission to work like an AI-powered analyst. Guerrilla AI means being "
    "crafty with what is already installed on your machine: the machine learning Excel has "
    "quietly shipped for years, the structural habits that make any AI dramatically better at "
    "reading your work, and the free tools that close the gap. ")

def opener(n, focus):
    return (f"Welcome, everyone. This is Guerrilla AI for Excel, day {n} of 5. I'm George "
            "Mount from Stringfest Analytics. " + THESIS_SPOKEN +
            f"Today's twenty-minute win: {focus} We have one live demo block and then two or "
            "three things you can do at your own desk this afternoon. Let's go.")

def recap_slide(prs, n):
    prev = C.DAYS[n - 1]
    return D.content(prs, f"Yesterday in 30 seconds",
                     [f"Day {n-1} was {prev['title']}.",
                      prev["takeaway"],
                      "Same dataset today. Every day builds on the last."],
                     subtitle="One dataset, five wins, zero licenses")

def build_day(n, slides, fname):
    prs = D.new_deck()
    for kind, args, notes in slides:
        fn = {"cover": D.cover, "divider": D.divider, "content": D.content,
              "code": D.code_slide, "two_col": D.two_col, "stat": D.stat_slide,
              "recap": lambda prs: recap_slide(prs, n)}[kind]
        s = fn(prs, *args) if kind != "recap" else fn(prs)
        if kind == "cover": wordmark(s)
        D.add_notes(s, notes)
    d = os.path.join(OUT, f"Day {n}"); os.makedirs(d, exist_ok=True)
    p = os.path.join(d, fname); prs.save(p); print("wrote", p, len(prs.slides._sldIdLst), "slides")

K = "Guerrilla AI for Excel. A live series"

# ================================================================== DAY 1 ====
d1 = C.DAYS[1]
day1 = [
("cover", (f"{K}, day 1 of 5", d1["title"],
           "Flash Fill, Analyze Data, and the models Microsoft never told you about.",
           C.FOOTER),
 opener(1, "the machine learning that has been hiding in your Excel for years.") +
 " Before we start, drop one thing in the chat: where you are joining from, and the Excel "
 "task that eats the most of your week. Keep an eye on your own answer. By Friday you will "
 "have a move for most of them."),

("content", ("How this week works",
             ["Live Monday to Friday, 12:00 to 12:20 PM Eastern. Twenty minutes, one win per day.",
              "Every session is recorded. Replays live in the Stringfest Analytics membership vault.",
              "Files and follow-ups go to the email you registered with on Eventbrite. Watch that inbox.",
              "One shared dataset all week: a genuinely ugly sales export from Lakeshore Office Supply.",
              "Demos are watch-along. Each day ends with two or three things to do at your desk after."],
             "Housekeeping, then we never speak of it again"),
 "Quick housekeeping, then we get to the good stuff. We run live every day this week, noon to "
 "twelve twenty Eastern. Twenty minutes is a promise, and I keep it. Every session is recorded, "
 "and the replays live in the Stringfest membership vault, so if you miss a day, that is where "
 "to find it. Anything I send, files, forms, follow-ups, goes to the email address you "
 "registered with on Eventbrite, so keep an eye on that inbox and reply from that address if "
 "you need me. We will use one dataset all week: a sales export from a fictional company "
 "called Lakeshore Office Supply, and it is ugly on purpose. In twenty minutes there is no "
 "time to type along, so watch the demo, then do the desk actions I give you at the end of "
 "each session. That is the deal. Let's talk about why you are here."),

("content", ("Why this counts as AI",
             ["Microsoft shipped machine learning inside Excel years ago and nobody told you.",
              "Flash Fill is pattern recognition. You give one example and it learns the rule.",
              "Analyze Data and Recommended PivotTables are models running against your data, no license required.",
              "The guerrilla skill is knowing these exist and when to reach for them.",
              d1["takeaway"]],
             "The guerrilla hook, day 1"),
 "Here is today's guerrilla hook, stated plainly. Microsoft shipped machine learning inside "
 "Excel years ago and nobody told you. Flash Fill, which has been sitting on the Data tab "
 "since Excel 2013, is pattern recognition. You type one example of what you want and it "
 "infers the rule. Analyze Data, right there on the Home tab, runs models against your "
 "data and answers questions you type in plain English. Recommended PivotTables is the same "
 "idea with training wheels. None of this needs a license, a budget request, or a "
 "conversation with IT. It is already on your machine. So before anyone in your organization "
 "asks for an AI budget, the guerrilla move is to exhaust the AI you already have. That is "
 "the whole day in one sentence. Now let me show you the dataset we will abuse all week."),

("content", ("Meet the mess: the Lakeshore export",
             C.DATASET_FLAWS,
             "One dataset, all five days"),
 "This is the Lakeshore Office Supply sales export, and if you work in finance or ops you "
 "have met this file before under a different name. Walk through the damage with me. The rep "
 "shows up only as an email address, nobody typed an actual name. Product and category are "
 "jammed into one column with a dash between them. Region has been typed four different "
 "ways, sometimes in caps, sometimes with a trailing space you cannot see. Dates arrive in "
 "two formats, some numbers are stored as text, and the prices have dollar signs baked into "
 "the cells. And someone helpfully embedded two TOTAL rows and a blank row right in the "
 "middle of the data. This exact file follows us through all five days, and each day we "
 "make it a little more useful. Today we start with the fastest win in Excel."),

("code", ("Live demo: Flash Fill",
          "One typed example teaches Excel the pattern. Watch column I.",
          [l for l in C.D1_FLASHFILL if l],
          "Machine learning, shipped in 2013, license cost: zero"),
 "Demo time. I am in the raw-export sheet. We need actual rep names, and all we have is "
 "email addresses like dana dot kowalski at lakeshoresupply dot com. Watch how fast this is. "
 "I click cell I1 and type the header, Rep Name. Now in I2, next to Dana's email, I type the "
 "answer I want: Dana Kowalski, capital D, capital K. I press Enter. Now the magic key: "
 "Control E. Flash Fill reads my one example, works out the pattern, and fills the entire "
 "column. Look at it go. Marcus Webb, Priya Raman, Tom Delgado, all correct. I will "
 "spot-check two rows near the bottom, because trust but verify applies to every tool this "
 "week. That is pattern recognition from a single training example. If you have ever spent "
 "an afternoon splitting names by hand, this just gave you your afternoon back."),

("code", ("Live demo: Analyze Data",
          "Ask a question in plain English, get a PivotTable back.",
          [l for l in C.D1_ANALYZE if l],
          "A model reads your data and answers in English"),
 "Second demo. I click anywhere in the data, go to the Home tab, and hit Analyze Data. A "
 "pane opens with suggestions Excel generated by scanning the data on its own. I will type a "
 "question instead: total Amount by Region. Excel parses my English, matches it to columns, "
 "and offers a PivotTable, which I insert with one click. Now, look closely at the result, "
 "because something is off. We will talk about it on the next slide, and I want you to see "
 "it with your own eyes first. While we are here: if the pane ever feels lost, the older "
 "sibling is Insert, Recommended PivotTables. Same machine learning, more manual. Both cost "
 "nothing and both have been sitting in front of you for years."),

("content", ("The honest moment",
             ["Analyze Data just showed seven regions. Lakeshore has four.",
              "MIDWEST, Midwest, and midwest-with-a-space read as three different places.",
              "The TOTAL rows quietly double-counted the answer.",
              "The AI is only as good as the data underneath it. Hold that thought for Wednesday.",
              "Also free: Copilot Chat in Excel, no purchased license. Licensing questions go to your IT team."],
             "Why the answer was wrong, and why that is the real lesson"),
 "Here is the honest moment, and it might be the most important slide of the day. Analyze "
 "Data just told us Lakeshore has seven regions. It has four. Because region was typed "
 "inconsistently, the model treated MIDWEST in caps, Midwest, and midwest with a trailing "
 "space as three different places. And those embedded TOTAL rows? They got counted like "
 "orders, so the numbers came back inflated. The tool did exactly what we asked. The data "
 "lied to it. That is the thread of this whole week: the AI is only as good as what it "
 "reads, and Wednesday is entirely about fixing that. One more thing while we are talking "
 "about free tools. Many of you have Copilot Chat available in Excel right now without a "
 "purchased license, and it is worth knowing what employees can already do with it. Whether "
 "and how it is enabled is a question for your IT department, so start there rather than "
 "with a purchase request."),

("content", ("Do this at your desk today", d1["desk"], "Ten minutes, tops"),
 "Your desk actions for today, and there are three. One: find a column you have been "
 "cleaning by hand, type one example of what you want next to it, and press Control E. Two: "
 "open a real dataset you care about, click Home, Analyze Data, and ask it one question "
 "that actually matters to you. Three, and this is the sneaky one: note what it gets wrong. "
 "Write the list down. Bring it tomorrow, because the fix for that list is the rest of this "
 "week. Ten minutes, all three."),

("divider", ("Tomorrow", "Make your data AI-readable",
             "The single highest-leverage prep move an analyst can make, and it costs nothing."),
 "Tomorrow we take the most boring-sounding feature in Excel, tables, and I will show you "
 "why it is the single highest-leverage AI-readiness move you can make. Same ugly dataset. "
 "You will watch Analyze Data go from confused to sharp with one keyboard shortcut. That is "
 "day 2, noon Eastern, twenty minutes."),

("divider", ("Thank you", "See you tomorrow at noon Eastern",
             "Replays and the full back catalog live in the Stringfest membership vault."),
 "That is day 1, and we are on time. Quick recap: Flash Fill is pattern recognition you "
 "already own, Analyze Data answers plain-English questions, and dirty data is what makes "
 "smart tools look dumb. If you want the replay, it lives in the Stringfest Analytics "
 "membership vault along with everything else I have recorded, ten dollars a month, link in "
 "the follow-up email. One question if we have time, otherwise drop it in the chat and I "
 "will answer by email. Thanks, everyone. See you tomorrow at noon."),
]

# ================================================================== DAY 2 ====
d2 = C.DAYS[2]
day2 = [
("cover", (f"{K}, day 2 of 5", d2["title"],
           "Tables and structured references: the prep move that makes every AI sharper.",
           C.FOOTER),
 opener(2, "making your data machine-readable with nothing but Control T and a good name.") +
 " Yesterday you saw smart tools stumble over a messy range. Today you see the same tools "
 "get noticeably sharper because of one habit."),

("recap", None,
 "Thirty-second recap for anyone just joining. Yesterday was the AI already hiding in your "
 "Excel: Flash Fill learned a pattern from one example, and Analyze Data answered a "
 "plain-English question with a PivotTable. The takeaway was to exhaust the AI you already "
 "have before asking for budget. But we also caught the tools being confidently wrong, "
 "seven regions instead of four, because the data underneath was a mess. Today we start "
 "fixing the underneath. Same Lakeshore export, one new habit."),

("content", ("Why this counts as AI",
             ["Every AI tool performs dramatically better on tables than on raw ranges.",
              "That applies to Analyze Data today, and to Copilot and Python later this week.",
              "Structured references make formulas self-documenting, readable by AI and by humans.",
              "Machine-readable data is the highest-leverage prep move while everyone else waits for licenses.",
              d2["takeaway"]],
             "The guerrilla hook, day 2"),
 "Today's guerrilla hook. Tables sound like the least exciting feature in Excel, so here is "
 "the reframe: every AI tool you will touch performs dramatically better on a table than on "
 "a raw range. That is true of Analyze Data, which you saw yesterday, and it is just as "
 "true of Copilot and of the Python we will run on Thursday. A table gives the machine "
 "named columns, a defined boundary, and one data type per field. On top of that, "
 "structured references make your formulas read like sentences, so a colleague, or an AI "
 "asked to audit your workbook, can tell what a formula means without decoding cell "
 "addresses. Here is the guerrilla part: while other teams wait for licenses to arrive, "
 "the analyst who makes data machine-readable today gets the payoff from every tool the "
 "moment it lands. The AI-ready workbook is a craft skill, not a purchase."),

("content", ("What Ctrl+T actually changes",
             ["Named columns the machine can point at, instead of guessing where data ends.",
              "One data type per column, which is how models expect data to arrive.",
              "Formulas that self-document: =SUM(SalesData[Amount]) needs no explanation.",
              "New rows join the table automatically, so answers stay current.",
              "Your colleagues get the same benefit the AI does. Readability compounds."],
             "The 13-minute demo, in two parts"),
 "Before I demo it, here is what pressing Control T actually buys you, because it is more "
 "than formatting stripes. First, named columns: the machine can point at Amount instead of "
 "guessing that column H probably means money. Second, a defined boundary with one data "
 "type per column, which is exactly the shape every model expects data to arrive in. Third, "
 "structured references: sum of SalesData Amount reads like a sentence, and formulas that "
 "read like sentences are formulas AI tools and new teammates can interpret. Fourth, tables "
 "grow. Paste new rows at the bottom and every formula, pivot, and AI answer that points at "
 "the table picks them up. Two-part demo: convert and name, then rerun Monday's question "
 "and compare."),

("code", ("Live demo: Ctrl+T and a real name",
          "Junk rows out, one shortcut, one name that means something.",
          [l for l in C.D2_TABLE if l],
          "A formula that reads like a sentence is AI-readable"),
 "Into the workbook. First, ten seconds of tidying: those two TOTAL rows and the blank row "
 "from yesterday, I am deleting them by hand, and notice that it is manual and annoying, "
 "because tomorrow we automate exactly this. Now I click anywhere in the data and press "
 "Control T. Excel finds the edges, I confirm my data has headers, and we have a table. "
 "Here is the step almost everyone skips: go to Table Design and name it. Table1 means "
 "nothing. I am typing SalesData. Now watch what that name buys us. In a spare cell I type "
 "equals SUM, open paren, and click the Amount column header. The formula writes itself: "
 "sum of SalesData Amount. Read that out loud. It documents itself. No dollar signs, no "
 "H2 colon H58, no archaeology six months from now."),

("code", ("Live demo: rerun Monday's question",
          "Same question as yesterday. Watch the quality jump.",
          [l for l in C.D2_RERUN if l],
          "Same tool, same data, better structure, better answer"),
 "Part two, and this is the before-and-after that makes the whole day. Click inside "
 "SalesData, Home tab, Analyze Data, and I ask the exact question from Monday: total Amount "
 "by Region. Compare. No junk rows polluting the totals, because the TOTAL rows are gone. "
 "The suggestions reference SalesData and its column names instead of vague ranges. The "
 "pane is faster to useful answers because it is no longer guessing where the data starts "
 "and stops. The region typos are still there, we fix those tomorrow, but the structural "
 "confusion is gone. Same tool, same twenty seconds of effort, dramatically better output. "
 "That is what machine-readable buys, and it cost you one keyboard shortcut and a name."),

("content", ("Do this at your desk today", d2["desk"], "The highest-leverage ten minutes of your week"),
 "Desk actions, three again. One: open the workbook you live in and press Control T on its "
 "main range, then give the table a real name, not Table1. Two: take one fragile formula, "
 "the kind with dollar signs and mystery ranges, rewrite it with structured references, and "
 "read it out loud. If it reads like a sentence, you did it right. Three: rerun the Analyze "
 "Data question you tried Monday and compare the answers side by side. That comparison is "
 "the proof you can show a skeptical coworker."),

("divider", ("Tomorrow", "Clean data so AI doesn't lie to you",
             "Power Query: the automation people think they need AI for, plus the fix for the seven-region lie."),
 "Tomorrow is the day the seven-region lie dies. Power Query records a cleanup once and "
 "replays it forever, which is the automated data prep people assume requires AI, delivered "
 "with zero AI. It is the applause-line demo of the week. Noon Eastern, twenty minutes, "
 "bring the export you hate cleaning."),

("divider", ("Thank you", "See you tomorrow at noon Eastern",
             "Replays live in the Stringfest membership vault."),
 "That is day 2. Recap: Control T plus a real name makes data machine-readable, structured "
 "references make formulas self-documenting, and every AI tool this week gets sharper "
 "because of it. Replay is in the membership vault. One question if time allows, otherwise "
 "chat or email. Thanks, everyone. Tomorrow the Refresh button gets a round of applause."),
]

# ================================================================== DAY 3 ====
d3 = C.DAYS[3]
day3 = [
("cover", (f"{K}, day 3 of 5", d3["title"],
           "Power Query: record the cleanup once, refresh it forever.",
           C.FOOTER),
 opener(3, "automated, repeatable data cleaning with a tool already on your Data tab.") +
 " This is the day with the applause line. Watch for the Refresh moment."),

("recap", None,
 "Quick recap. Monday: the machine learning already in Excel, and the discovery that dirty "
 "data makes smart tools lie, seven regions instead of four. Tuesday: Control T and "
 "structured references made the workbook machine-readable, and Analyze Data got sharper. "
 "The takeaway was that the AI-ready workbook is a craft skill, not a purchase. Today we "
 "kill the remaining lies in this dataset, and we do it in a way that never has to be "
 "done twice."),

("content", ("Why this counts as AI",
             ["Punch one: Power Query delivers what people think they need AI for. Automated, repeatable, refreshable prep, zero AI inside.",
              "Recorded steps are automation without programming.",
              "Punch two: clean inputs are what stop AI from hallucinating your numbers.",
              "Garbage in, confident-sounding garbage out.",
              "Bonus move: free chatbots write M code, and Power Query is where you paste it."],
             "The guerrilla hook, day 3"),
 "Today's hook lands two punches. Punch one: when people say they want AI to automate their "
 "data prep, the outcome they are describing, automated, repeatable, refreshable cleanup, "
 "has been shipping inside Excel for a decade as Power Query, and there is no AI in it at "
 "all. It records your steps like a macro recorder that actually works, and replays them on "
 "demand. Automation without programming. Punch two: remember Monday, when Analyze Data "
 "swore we had seven regions? Clean inputs are what stop AI tools from hallucinating your "
 "numbers. Garbage in, confident-sounding garbage out, and the confident tone is what makes "
 "it dangerous in front of your CFO. And a bonus guerrilla move to keep in your pocket: "
 "free chatbots write perfectly good M code, Power Query's language, and the Advanced "
 "Editor is where you paste it. Let's clean this thing for the last time."),

("code", ("Live demo: record the cleanup once",
          "Six recorded steps. Every flaw from Monday, fixed in order.",
          [l for l in C.D3_TRANSFORMS if l],
          "You just wrote a program by clicking"),
 "Main demo. I click in the RawExport table, Data tab, From Table Range, and the Power "
 "Query editor opens. Now I just clean, and it records. Step one: filter Order ID, uncheck "
 "the blanks and those two TOTAL rows. Gone, and unlike yesterday's manual delete, this is "
 "recorded. Step two: select Product dash Category, Split Column by delimiter, space dash "
 "space, and we get real Product and Category columns. Step three: Region. Transform, "
 "Format, Trim, then Capitalize Each Word. MIDWEST in caps, midwest with the trailing "
 "space, all collapse into one clean Midwest. The seven-region lie dies right there. Step "
 "four: types. Order Date to Date, Units to Whole Number, Unit Price to Currency, and the "
 "dollar signs and text numbers all resolve. Watch the Applied Steps pane on the right: "
 "that is a program you just wrote by clicking. Close and Load To, new sheet, and I name "
 "the output table SalesClean."),

("code", ("Live demo: the Refresh moment",
          "June just arrived. Watch what does not happen next.",
          [l for l in C.D3_REFRESH if l],
          "One hour of Power Query replaces a recurring chore forever"),
 "Now the applause line. It is a new month at Lakeshore and June's orders just landed, "
 "same ugly format, of course. I copy the June rows from the june-rows sheet and paste "
 "them at the bottom of RawExport. Now I right-click the SalesClean table and hit Refresh. "
 "Done. Every recorded step replayed on the new data: junk filtered, product split, "
 "regions standardized, types fixed, June included. Nobody re-cleaned anything. That "
 "monthly chore you were going to do by hand forever is now one click, and next month it "
 "is still one click. This is the single best return on an hour I know of in Excel: build "
 "the query once, refresh it for the rest of your tenure."),

("content", ("Do this at your desk today", d3["desk"], "Pick your ugliest recurring export"),
 "Desk actions. One: pick the export you clean every week or every month, the one you "
 "dread, and rebuild the cleanup one time in Power Query, clicking through the same moves "
 "you saw today. Two: next cycle, paste the new data in and press Refresh instead of "
 "redoing the work, then enjoy the feeling. Three: if you hit a transform you cannot "
 "figure out, describe it to a free chatbot in plain English and ask for the M code, then "
 "paste that into the Advanced Editor. That last one is pure guerrilla: free AI writing "
 "code for a free tool you already had."),

("divider", ("Tomorrow", "A data science lab behind =PY()",
             "pandas and seaborn are already inside Excel. No install, no admin rights, no ticket."),
 "Tomorrow might be the biggest surprise of the week for some of you. There is a full "
 "Python data science environment, pandas, seaborn, the real thing, sitting behind an "
 "Excel formula called PY. No install, no admin rights, no IT ticket. And the clean table "
 "we built today is exactly what it wants to eat. Noon Eastern."),

("divider", ("Thank you", "See you tomorrow at noon Eastern",
             "Replays live in the Stringfest membership vault."),
 "That is day 3, and the seven-region lie is officially dead. Recap: Power Query is "
 "recorded, repeatable cleanup, the Refresh click replays it on every new month, and clean "
 "inputs are what keep every AI tool this week honest. Replay in the membership vault, "
 "questions in chat or by email. Thanks, everyone. Tomorrow we write Python without "
 "leaving Excel."),
]

# ================================================================== DAY 4 ====
d4 = C.DAYS[4]
day4 = [
("cover", (f"{K}, day 4 of 5", d4["title"],
           "pandas, seaborn, and a chatbot that writes the code for you.",
           C.FOOTER),
 opener(4, "running real Python, pandas and seaborn, inside Excel with no install and no "
        "permission slip.") +
 " If you have ever been told data science requires a whole new toolchain, today is for you."),

("recap", None,
 "Recap in thirty seconds. Monday, the machine learning already in Excel. Tuesday, tables "
 "and structured references made the workbook machine-readable. Yesterday, Power Query "
 "recorded the cleanup once, and one Refresh click replayed it on June's data. The "
 "takeaway: one hour of Power Query replaces a recurring chore forever. Today we point "
 "something much more powerful at that clean SalesClean table, and it has been hiding "
 "behind an equals sign the whole time."),

("content", ("Why this counts as AI",
             ["A full Python runtime with pandas and seaborn ships inside Excel. It runs in Microsoft's cloud.",
              "No install, no admin rights, no IT ticket.",
              "Free chatbots write competent Python. Describe what you want, paste it into a PY cell, run it.",
              "This is where AI code generation pays off even without Copilot.",
              d4["takeaway"]],
             "The guerrilla hook, day 4"),
 "Today's guerrilla hook. Inside current Microsoft 365 Excel there is a formula called PY, "
 "and behind it sits a full Python runtime with pandas and seaborn, the same libraries "
 "professional data scientists use every day. It runs in Microsoft's cloud, which means no "
 "install, no admin rights, and no IT ticket. You type equals PY and you are in. Now the "
 "guerrilla part. The classic objection to Python is that you have to learn to write it. "
 "But free chatbots write competent Python, and they are especially good at exactly the "
 "small, well-described analysis tasks an analyst needs. So the workflow is: describe what "
 "you want in plain English, let the chatbot draft the code, paste it into a PY cell, run "
 "it against your table, and judge the result. You already have a data science "
 "environment. The craft is knowing how to ask for the code."),

("content", ("How PY() sees your workbook",
             ["Type =PY( in a cell and Excel switches to a Python editor.",
              "xl(\"SalesClean[#All]\", headers=True) hands your table to pandas as a DataFrame.",
              "pandas is preloaded as pd, seaborn as sns. No imports needed for today's demo.",
              "Results spill back into the grid, numbers or charts.",
              "Yesterday's clean table is exactly the input Python wants."],
             "Two code blocks and one chart, all live"),
 "One minute of orientation before the demo. When I type equals PY and an open paren in a "
 "cell, the formula bar becomes a small Python editor. The bridge between the grid and "
 "Python is a function called xl. When I write xl of SalesClean bracket hash All, headers "
 "true, Excel hands the whole table to pandas as a DataFrame, which is just Python's word "
 "for a table. The essentials are preloaded: pandas is already there as pd, seaborn as "
 "sns, so today's demo needs zero import statements. Whatever the code produces, a table "
 "of numbers or a chart, spills back into the grid like any other formula result. And "
 "notice: it wants a clean, named table. Wednesday's work was the price of admission, and "
 "we already paid it."),

("code", ("Live demo: stats in one cell",
          "describe(), then a groupby that replaces a pivot.",
          [l for l in C.D4_BLOCK1 if l] + [l for l in C.D4_BLOCK2 if l],
          "Analysis that took a pivot and six clicks, in one line"),
 "First block. I click a cell to the right of the data, type equals PY, and enter two "
 "lines: df equals xl SalesClean hash All, headers true, then df dot describe. Control "
 "Enter runs it. Out spills a statistical profile of every numeric column: count, mean, "
 "standard deviation, min, max, quartiles. That summary would take a dozen formulas the "
 "manual way. Second block, in the next cell down: df dot groupby Region, take Amount, "
 "sum it, sort it descending. One line, and it answers the question we have asked all "
 "week, total sales by region, biggest first. Compare that with Monday, when we needed a "
 "PivotTable, and notice the answer is finally trustworthy because the regions are clean. "
 "Same question all week, better tools every day."),

("code", ("Live demo: the chart a chatbot wrote",
          "The describe-to-chatbot, paste-to-cell workflow, live.",
          [l for l in C.D4_BLOCK3 if l],
          "Describe what you want. Paste. Run. Judge."),
 "Last block, and this is the workflow slide. Suppose I do not know seaborn. I ask a free "
 "chatbot: I have an Excel table called SalesClean with Region and Amount columns, give me "
 "Python for a bar chart of total sales by region, for a PY cell in Excel, colored "
 "CF3338. That hex code is Stringfest red, because brand matters even in a demo. The "
 "chatbot returns roughly what you see here: group the totals, then sns dot barplot. I "
 "paste it into a PY cell, run it, and a chart spills into the grid. Right-click, display "
 "plot over cells, and there it is, a presentable chart. The skill I actually used was "
 "describing my table and my goal precisely. Tomorrow is entirely about getting better at "
 "that describing, because it turns out there is a language for it."),

("content", ("Do this at your desk today", d4["desk"], "Prove to yourself it is really there"),
 "Desk actions. One: type equals PY open paren in any cell of a current Microsoft 365 "
 "workbook, just to watch the editor appear. If it does not, note that availability "
 "depends on your Microsoft 365 plan and channel, worth one question to IT, and everything "
 "else this week still works. Two: point df equals xl at a table you own, run df dot "
 "describe, and read the profile. Three: ask a free chatbot for one seaborn chart of your "
 "table and paste the code into a PY cell. Describe your columns when you ask. That "
 "one-sentence description is tomorrow's topic in miniature."),

("divider", ("Tomorrow", "Markdown: the language AI speaks",
             "The cheapest AI upgrade there is, and the finale of the week."),
 "Tomorrow we close the week with the cheapest upgrade in the whole series: Markdown, the "
 "formatting language every AI tool reads and writes. Twenty minutes of syntax that makes "
 "every prompt sharper and every answer easier to reuse, including a trick for moving "
 "tables between Excel and any chatbot with no connectors at all. Noon Eastern, one last "
 "time."),

("divider", ("Thank you", "See you tomorrow at noon Eastern",
             "Replays live in the Stringfest membership vault."),
 "That is day 4. Recap: a real Python lab lives behind equals PY, xl hands it your clean "
 "tables, and free chatbots will write the code if you describe the goal well. Replay in "
 "the membership vault. Thanks, everyone. Tomorrow we finish the week strong."),
]

# ================================================================== DAY 5 ====
d5 = C.DAYS[5]
day5 = [
("cover", (f"{K}, day 5 of 5", d5["title"],
           "Structured prompts, round-trip tables, and the close of the week.",
           C.FOOTER),
 opener(5, "learning the language every AI tool already speaks, and using it to move work "
        "in and out of Excel.") +
 " It is the finale, so we also recap the week and talk about where to go from here."),

("recap", None,
 "Last recap of the week. Monday, the machine learning already hiding in Excel. Tuesday, "
 "tables made the workbook machine-readable. Wednesday, Power Query recorded the cleanup "
 "and Refresh replayed it. Yesterday, real Python behind equals PY, with a chatbot "
 "writing the code. One thread ran through all of it: the better the machine can read "
 "your work, the better every tool performs. Today we apply that same idea to the words "
 "you send AI tools, because it turns out prompts have a structure too."),

("content", ("Why this counts as AI",
             ["Prompts, Copilot output, chatbot answers, and agent instruction files are all Markdown.",
              "Twenty minutes of syntax makes every AI tool easier to steer and its output easier to reuse.",
              "Tables round-trip between Excel and chatbots cleanly as Markdown.",
              "No connectors, no licenses, no IT ticket. Just structure.",
              d5["takeaway"]],
             "The guerrilla hook, day 5"),
 "The final guerrilla hook. Behind almost every AI tool you touch is the same plain-text "
 "formatting language: Markdown. Hash marks for headings, dashes for bullets, pipes for "
 "tables. Chatbot answers are written in it. Copilot emits it. The instruction files "
 "that steer AI agents are plain Markdown documents. Which means twenty minutes of "
 "syntax buys you two powers. Steering: a prompt with headings and bullet structure is "
 "dramatically easier for a model to follow than a wall of text. And portability: tables "
 "written in Markdown move cleanly between Excel and any chatbot, in either direction, "
 "with no connectors and no licenses. Speaking the machine's language is the cheapest AI "
 "upgrade there is. Two demos and then we close out the week."),

("code", ("Live demo: structure beats the blob",
          "Same request, two prompts. Watch the quality gap.",
          [l for l in C.D5_BLOB if l] + [""] + [l for l in STRUCTURED if l][:8] +
          ["... (full prompt in today's demo file)"],
          "Headings and bullets are steering, and steering is free"),
 "Demo one. On screen are two prompts asking for the same thing. The first is the blob: "
 "summarize this sales data and tell me whats important, followed by a messy paste of the "
 "whole export. We have all sent that prompt. The second is structured Markdown. A "
 "heading names the job: quarterly regional sales review. A Role section says who to be, "
 "an FP&A analyst. A Task section lists three bullets: summarize by region, flag the "
 "trailing region, suggest two follow-ups. A Data section carries our regional summary as "
 "a clean Markdown table, and a Format section dictates the shape of the answer. I send "
 "both to a free chatbot, live. The blob comes back generic and rambling. The structured "
 "prompt comes back organized, on-task, in exactly the format I dictated, and it names "
 "the trailing region without being coaxed. Same model, same data, same price of zero. "
 "Structure did all of that."),

("code", ("Live demo: the Excel round-trip",
          "Copy out, Markdown back, Text to Columns. No connectors.",
          [l for l in C.D5_ROUNDTRIP if l],
          "The guerrilla data pipeline: copy, paste, pipe, done"),
 "Demo two, the round-trip. Going out is easy: I copy the RegionSummary table in Excel "
 "and paste it straight into the chat. It lands as a table the model reads cleanly, and "
 "you saw it handle that table perfectly in demo one. Coming back is the part nobody "
 "shows you. I ask the chatbot for its results as a Markdown table, and it returns rows "
 "with pipe characters between the columns. I copy that, paste it into a blank area of "
 "the sheet, and it lands squashed into one column. Now: Data tab, Text to Columns, "
 "Delimited, and the delimiter is the pipe character. Finish, trim the empty edge "
 "columns, and there it is, a real Excel range again. Copy out, Markdown back, Text to "
 "Columns. That is a data pipeline between Excel and any AI tool on earth, built from "
 "features that predate all of them."),

("content", ("Do this at your desk today", d5["desk"], "The habit that outlasts the week"),
 "Final desk actions of the series. One: take the next real prompt you were going to "
 "send anyway and rewrite it with a heading, a bulleted task list, and a Format section, "
 "then compare answers. That comparison converts people. Two: paste one small Excel "
 "table into a chatbot and ask for the reply as a Markdown table, then bring it home "
 "with Text to Columns. Three: save your best prompt as a dot md file. That file is "
 "reusable craft, and it is the seed of every agent instruction file you will ever "
 "write."),

("stat", ("The week, in numbers",
          [("5", "twenty-minute wins"), ("0", "licenses required"),
           ("$0", "spent on tools"), ("1", "ugly export, fully tamed")],
          "Flash Fill. Tables. Power Query. Python. Markdown. All already on your machine."),
 "Here is the week in numbers. Five twenty-minute wins. Zero licenses required. Zero "
 "dollars spent. One genuinely ugly sales export, fully tamed. Monday you found the "
 "machine learning already in Excel. Tuesday you made your data machine-readable. "
 "Wednesday you recorded the cleanup once and refreshed it forever. Thursday you ran "
 "real Python behind equals PY. Today you learned the language that steers all of it. "
 "None of that came from a budget line. It came from craft with tools you already had, "
 "which was the promise of the series. So, where next?"),

("two_col", ("Where to go from here",
             "Keep the momentum",
             ["Replays of all five sessions live in the membership vault.",
              "Ten dollars a month, cancel anytime, full back catalog included.",
              "Rewatch the demos with the files in front of you."],
             "Go deeper in December",
             ["AI for Excel in a Week: five live 60-minute sessions.",
              "Everything from this week, hands-on and in depth.",
              "Registration link lands in tomorrow's thank-you email."],
             "Two paths, and they stack"),
 "Two ways to keep going, and they stack. First, the replays. All five sessions from "
 "this week live in the Stringfest Analytics membership vault, along with my full back "
 "catalog of recordings. Ten dollars a month, cancel anytime. If you want to rewatch "
 "Wednesday with the workbook open in front of you, that is the place. Second, the "
 "deeper version. This December I am running AI for Excel in a Week: five live sessions, "
 "sixty minutes each, that take everything you saw here further, hands-on, with time to "
 "actually build alongside me. If this week was the espresso shot, December is the full "
 "pour. The registration link will be in tomorrow's thank-you email, sent to the address "
 "you registered with. Membership today, December for depth, or both."),

("divider", ("Thank you", "Five days, five wins, zero licenses",
             "Watch your inbox for the thank-you email with links, files, and the December cohort."),
 "That is the series. Thank you, sincerely, for giving me twenty minutes a day of your "
 "week. Everything lands in your inbox tomorrow: the files, the links, the membership, "
 "and December's cohort. My ask is simple: pick one of the five moves and use it on real "
 "work this afternoon, because the move you use within a day is the one you keep. If "
 "this week saved you time, tell one coworker who is still cleaning columns by hand. "
 "I'm George Mount from Stringfest Analytics. Go be crafty."),
]

DECKS = {1: day1, 2: day2, 3: day3, 4: day4, 5: day5}
for n, slides in DECKS.items():
    build_day(n, slides, f"Guerrilla AI - Day {n} - Slides.pptx")
print("decks done")
