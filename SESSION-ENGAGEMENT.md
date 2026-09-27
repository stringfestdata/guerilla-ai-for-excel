# Guerrilla AI: participation guide

Five 20-minute sessions. Five short invitations per day. All work as chat votes; no external poll setup is required. Prompts are embedded in the decks. Demo questions and next actions are mirrored in the workbooks.

## Before going live

Keep chat visible beside the shared window. Put the opening slide up while people join. Read the choices aloud. During Excel demos, say or paste the short question in chat because the slide may no longer be visible. Watching is welcome; nobody needs to keep up with typing.

Ask once, wait about eight seconds, acknowledge one or two real answers, and continue. Most cues take about 15 seconds in total. The final question menu can use the remaining minute for one short explanation. If nobody answers, use the prepared line and continue without commenting on silence.

## The 20-minute shape

| Time | What happens |
|---|---|
| 0:00-1:00 | Outcome and one-word opening choice |
| 1:00-4:00 | Essential context and setup |
| 4:00-10:00 | First demo with a short chat check |
| 10:00-15:30 | Second demo with a short chat check |
| 15:30-18:00 | Choose one desk action and preview what comes next |
| 18:00-20:00 | Specific question menu, one useful explanation, close |

Times are cues, not a stopwatch script. If behind, skip the optional next-action vote and one demo vote. Keep the useful demo and the 20-minute finish. Do not wait longer to make somebody answer.

## All 25 prompts at a glance

| Day | Time | Slide | Question | Reply |
|---|---|---|---|---|
| 1 | 0:30 | 1 | Which takes more of your time? | CLEANING / REPORTING |
| 1 | 6:00 | 5 | Email changes later. Will Flash Fill update? | YES / NO |
| 1 | 11:00 | 6 | Before adding up sales: keep the TOTAL rows? | KEEP / REMOVE |
| 1 | 16:30 | 8 | Which will you try first? | NAMES / TOTALS |
| 1 | 18:30 | 10 | Which part should I revisit? | NAMES / REGIONS / TOTALS |
| 2 | 0:30 | 1 | Do you use Excel tables already? | YES / NO |
| 2 | 6:00 | 5 | Easier to recognize? | A = Table1; B = SalesData |
| 2 | 11:00 | 6 | Did Ctrl+T fix the region typos? | YES / NO |
| 2 | 16:30 | 7 | A name for your first table? | SALES / EXPENSES / OTHER |
| 2 | 18:30 | 9 | Which step should I show again? | CREATE / NAME / FORMULA |
| 3 | 0:30 | 1 | Do you repeat the same cleanup? | YES / NO |
| 3 | 6:00 | 4 | Which removes extra spaces? | A = Trim; B = Split |
| 3 | 12:00 | 5 | The next month arrives. What do we do? | REDO / REFRESH |
| 3 | 16:30 | 6 | Which cleanup would you save? | SPACES / SPLIT / TYPES |
| 3 | 18:30 | 8 | Which part should I revisit? | FILTER / SPLIT / REFRESH |
| 4 | 0:30 | 1 | Where are you with Python? | NEW / TRIED |
| 4 | 7:00 | 5 | For total sales: add or count? | SUM / COUNT |
| 4 | 12:00 | 6 | Which would you chart? | SALES / ORDERS |
| 4 | 16:30 | 7 | Which result would help you first? | SUMMARY / CHART |
| 4 | 18:30 | 9 | Which part should I revisit? | SETUP / SUMMARY / CHART |
| 5 | 0:30 | 1 | Your usual AI prompt? | SHORT / DETAILED |
| 5 | 6:00 | 4 | Easier for you to scan? | A = paragraph; B = headings |
| 5 | 12:00 | 5 | Which separator is in this table? | PIPE / COMMA |
| 5 | 16:30 | 6 | What will you add to your next prompt? | TASK / FORMAT |
| 5 | 18:30 | 9 | Which part should I revisit? | PROMPT / TABLE |

## Read-aloud lines and fallbacks

### Day 1

**0:30, slide 1**

> Let's try the chat with one word: CLEANING or REPORTING. Which takes more of your time? Either answer works. You are welcome to just watch, too.

Pause once for about eight seconds. Acknowledge one actual answer: 'Thanks. Keep that task in mind as we look at the export.'

If quiet: **I'll use cleaning as our starting point.**

**6:00, slide 5**

> Quick guess, just YES or NO: if I change the email later, will that filled name update too?

Pause once for about eight seconds. NO. Flash Fill writes values. It does not create a formula linked to the email. Show the filled cell in the formula bar, then continue.

If quiet: **Let's check the filled cell. It is a value, so it will not update with the email.**

**11:00, slide 6**

> Before we add up sales, what should happen to the TOTAL rows: KEEP or REMOVE? One word is plenty.

Pause once for about eight seconds. REMOVE. They already summarize other rows. Including them as orders can count sales twice. Today, spot the problem. Tomorrow, remove them.

If quiet: **I'll remove summary rows before adding up the individual orders.**

**16:30, slide 8**

> Pick just one thing for later: NAMES for Flash Fill, or TOTALS for checking a summary. Type either word.

Pause once for about eight seconds. Acknowledge an actual choice. NAMES: type one correct example and try Flash Fill. TOTALS: check one summary for included subtotal rows.

If quiet: **I'll start with NAMES: one column and one example.**

**18:30, slide 10**

> Want me to revisit something? Type NAMES, REGIONS, or TOTALS. A full question is welcome too.

Pause once for about eight seconds. Answer one actual request in about 30 seconds. NAMES: check the example matches its row. REGIONS: spaces and inconsistent labels need cleaning. TOTALS: remove summary rows before summing orders.

If quiet: **One useful check before you go: make sure your example matches the email on that same row.**

### Day 2

**0:30, slide 1**

> One word in chat to get us started: do you use Excel tables already, YES or NO? No explanation needed. Watching is welcome too.

Pause once for about eight seconds. Acknowledge an actual answer. 'I'll show the shortcut and the naming step, so either way you can follow.'

If quiet: **I'll show it from the beginning.**

**6:00, slide 5**

> Quick vote: which name tells you more, A for Table1 or B for SalesData? Just the letter.

Pause once for about eight seconds. B. SalesData tells us what the table contains. Show the Table Name box, then read SUM(SalesData[Amount]) out loud.

If quiet: **I'll use B, SalesData, so the formula tells us what it is adding up.**

**11:00, slide 6**

> We made a table. Did that fix the region typos, YES or NO? You can read the labels on my screen.

Pause once for about eight seconds. NO. Ctrl+T adds table structure, not data cleanup. Point to a remaining inconsistent Region label. Tomorrow's Power Query steps address that.

If quiet: **The answer is NO. The table still contains the labels we put into it.**

**16:30, slide 7**

> For your own workbook, which table name would fit: SALES, EXPENSES, or OTHER? One word is enough.

Pause once for about eight seconds. Acknowledge an actual choice. Give the main range a table name that describes its contents. This is the one desk action to prioritize.

If quiet: **I'll use SalesData for this example. You can choose a name from your own work.**

**18:30, slide 9**

> Type CREATE, NAME, or FORMULA if you would like that step again. You do not need to write a whole question.

Pause once for about eight seconds. Answer one actual request. CREATE: select the range and press Ctrl+T. NAME: Table Design, Table Name. FORMULA: =SUM(SalesData[Amount]).

If quiet: **The step I would double-check is the name: look in Table Design before writing the formula.**

### Day 3

**0:30, slide 1**

> One word to start: do you repeat the same cleanup each week or month, YES or NO? Watching is welcome too.

Pause once for about eight seconds. Acknowledge an actual answer. 'Keep that export in mind. Today we record the steps for next time.'

If quiet: **I'll use our monthly sales export as the example.**

**6:00, slide 4**

> Quick choice from the steps on screen: which removes extra spaces, A for Trim or B for Split?

Pause once for about eight seconds. A. Trim removes leading and trailing spaces. Split separates the product and category. Point to Trim in the recorded steps.

If quiet: **I'll choose A, Trim, for the spaces around the region names.**

**12:00, slide 5**

> June has arrived. One word before I click: REDO or REFRESH?

Pause once for about eight seconds. REFRESH. With new rows inside RawExport, refresh SalesClean. Show a new June order in the output so the payoff is visible.

If quiet: **I'll click REFRESH and check that a June order made it into the result.**

**16:30, slide 6**

> Which step would you most like to stop repeating: SPACES, SPLIT, or TYPES? Pick just one word.

Pause once for about eight seconds. Acknowledge an actual choice and connect it to that recorded step. The desk action is to save one small query for a recurring export.

If quiet: **I'll pick SPACES. That is a small first query you can build and refresh.**

**18:30, slide 8**

> If you want a step again, type FILTER, SPLIT, or REFRESH. A full question works too.

Pause once for about eight seconds. Answer one actual request. FILTER: remove blank and TOTAL rows. SPLIT: use the space-dash-space delimiter. REFRESH: put new rows inside RawExport, then refresh SalesClean.

If quiet: **One useful reminder: new rows need to be inside the source table before you refresh.**

### Day 4

**0:30, slide 1**

> One word in chat: NEW if Python is new to you, or TRIED if you have used it before. No code required. Watching is welcome too.

Pause once for about eight seconds. Acknowledge an actual answer. 'I'll describe what the result means as we go.'

If quiet: **I'll explain the result in everyday Excel terms.**

**7:00, slide 5**

> For total sales by region, should we SUM the amounts or COUNT the orders? Just type SUM or COUNT.

Pause once for about eight seconds. SUM. COUNT answers how many orders. Point to sum() in the code and the resulting regional totals.

If quiet: **I'll use SUM because our question is about sales dollars.**

**12:00, slide 6**

> Which would help in your work: a chart of SALES or a chart of ORDERS? One word is plenty.

Pause once for about eight seconds. Acknowledge a real choice. Today's chart shows sales. Explain that an orders chart would count order rows instead. Check the plotted sales against the table before calling the chart correct.

If quiet: **I'll show SALES today, then check the bars against our totals.**

**16:30, slide 7**

> Pick an outcome for later: SUMMARY or CHART. You can answer even if Python is not available in your Excel.

Pause once for about eight seconds. Acknowledge an actual choice. If PY is available, try one code block. If not, use the displayed results to decide what you would ask the tool to do.

If quiet: **I'll start with SUMMARY. Knowing the question comes before choosing the code.**

**18:30, slide 9**

> Type SETUP, SUMMARY, or CHART if you want that part again. You do not need to know the Python vocabulary.

Pause once for about eight seconds. Answer one actual request. SETUP: PY availability depends on the account and supported Excel setup. SUMMARY: describe gives a numeric profile, groupby sums sales by region. CHART: compare it with the totals.

If quiet: **A useful check is to compare the chart with the numbers. A chart that runs still needs a review.**

### Day 5

**0:30, slide 1**

> One word to start: is your usual AI prompt SHORT or DETAILED? Either is fine. Watching is welcome too.

Pause once for about eight seconds. Acknowledge one actual answer. 'We will give the request a clear task and output format.'

If quiet: **I'll start with a short request and make the task clearer.**

**6:00, slide 4**

> Which is easier for you to scan, A for the paragraph or B for headings? Just the letter.

Pause once for about eight seconds. Acknowledge the choice without claiming a model result is guaranteed. Show the Task and Format sections. Compare the actual AI replies for requested content and errors.

If quiet: **I'll use B, headings, so we can find the task and requested format quickly.**

**12:00, slide 5**

> Look between the Markdown table columns: is the separator PIPE or COMMA? One word.

Pause once for about eight seconds. PIPE. Show the | character. After splitting, remove the Markdown separator row and empty edge columns, and verify headers and numbers.

If quiet: **It is PIPE, the vertical line between columns. I'll use it in Text to Columns.**

**16:30, slide 6**

> Pick one improvement for your next prompt: TASK or FORMAT. Type either word.

Pause once for about eight seconds. Acknowledge an actual choice. TASK: say exactly what to do. FORMAT: say how you want the answer returned. Try one change before adding more.

If quiet: **I'll choose FORMAT, for example: return three bullets and a table.**

**18:30, slide 9**

> One last chance to steer the demo: type PROMPT or TABLE if you want that part again. Full questions are welcome too.

Pause once for about eight seconds. Answer one actual request. PROMPT: show Task and Format. TABLE: point to the pipe delimiter and remove the separator row. Thank participants once, without commenting on how many replied.

If quiet: **One useful final check: compare the returned table with the original before using its numbers.**

## Small corrections made during the engagement edit

- Day 1 now uses Marcus Webb for I2, matching the email in the first data row. The deck, workbook, and notes agree.
- Watch-along instructions now match the 20-minute format. Day 4 no longer asks people to type every line live.
- The table discussion separates table structure from cleaning values. The prompt asks whether Ctrl+T actually fixed the region labels.
- Notes tell you to inspect the real AI result instead of promising a particular region count or a guaranteed better reply.
- The Markdown return trip explicitly removes the separator row and empty edge columns.
- Availability wording distinguishes not buying Copilot from already having a qualifying Excel setup. Sources: [Microsoft: Analyze Data](https://support.microsoft.com/en-us/excel/analyze-data-in-excel) and [Microsoft: Python in Excel availability](https://support.microsoft.com/en-us/excel/python/python-in-excel-availability), checked September 27, 2026.

## Future session guidance

The standing preference is saved in the global Codex instructions and the personal `stringfest-session-engagement` skill. The course-in-a-week, PeopleSkillTraining webinar-kit, TTS preparation, and course presenter-note guidance also point to it. Future session builds should include these prompts without another reminder.
