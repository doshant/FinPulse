# FinPulse: Finance Dashboard add-in

FinPulse is a VSTO Excel add-in for exploring the bank's monthly results. A task pane lets a finance user pick a metric, business segment, country, monthly or quarterly periods and how many periods to show. The numbers, the results table and a native Excel chart on the **Dashboard** sheet update as soon as anything changes. There is no Apply button.

## How to run it

**Requirements:** Windows, desktop Excel, and Visual Studio with the *Office/SharePoint development* workload.

1. Clone the repository and open `FinPulse.slnx` in Visual Studio.
2. Set `FinPulse` as the startup project and press **F5**. Excel starts with the add-in loaded.
3. Open `Financials.xlsx` (sheets: Financials, Metrics, Dashboard).
4. On the ribbon, click **Launch FinPulse**. The pane opens on the right and the workbook jumps to the Dashboard sheet. After the first click the same button shows or hides the pane.
5. Change any setting in the pane. The table, summary and chart update straight away.

## Project structure

| Project / file | What it holds |
|---|---|
| `FinPulse` (VSTO) | `ThisAddIn.cs` hosts the task pane, reads the sheets and reloads when they change. `RibbonToggle.cs` is the ribbon button. `DashboardWriter.cs` writes the table and chart to the Dashboard sheet. |
| `FinPulse.UI` (WPF) | `FinancePane.xaml(.cs)` is the pane. `Engine/` holds all the parsing and calculation code (`DatasetBuilder.cs`, `FinanceEngine.cs`). The engine uses no Excel types, so it is independent of the UI and the workbook. |

**Data rule:** the add-in only reads the Financials and Metrics sheets and never changes them. It writes only to the Dashboard sheet: the chart (a single object named `FinPulseChart`, in columns A to L) and a data table to its right (from column N) so every plotted number is traceable.

## What the user can do

| Control | Options |
|---|---|
| Metric | Every metric on the Metrics sheet, plus a built-in **Operating profit** (revenue minus operating expenses) |
| Show | **Value**, **vs previous** period, or **vs last year** (same period a year earlier) |
| Business segment / Country | Any one, or All |
| Compare by | None, **Segment** or **Country**: one line or bar per segment or country (up to 8) |
| Reporting frequency | Monthly or Quarterly |
| Number of periods | 3, 4, 6, 12 or All |
| Chart type | Column or Line |
| Order of periods | Newest first (default) or Oldest first, in the table and on the chart |
| Utilities | **Refresh data** (re-read the sheets), **Reset**, **Copy table** (paste straight into Excel) |

### What the user sees

- **Snapshot card:** the latest value, a coloured change badge against the previous period, the same period last year, and a sparkline.
- **Results table:** period, value and change per row. An amber dot marks any period that includes an estimate or is only part of a quarter, with a tooltip explaining why.
- **Key takeaways:** a short written summary of the overall move, the biggest single move, and a note when the direction is unfavourable. In compare mode it also names the highest and the fastest-rising segment or country.
- **Data quality notice:** an amber panel listing every blank cell that was filled in, the value used and how it was worked out.
- **Dashboard chart:** the latest period is highlighted and labelled, estimates are drawn in amber, and a "Data check" column beside the chart data says which periods are complete.

### Growth views (vs previous, vs last year)

In a growth view the chart keeps the underlying values and adds the change:

- **Bars** show the raw value on the left axis.
- A **line** shows the % change on a secondary right axis. With the Line chart type both are lines and the change line is dashed.
- The table and the Copy table output carry both the value and the change.
- Cost-to-income, being a ratio, changes in **percentage points** rather than %.
- Compare mode with a growth view shows the change only, because several raw series plus several change series on two axes would be unreadable.

Good and bad moves are coloured by meaning, not direction: costs and the cost-to-income ratio are better when they fall, so a rise shows red.

## How the numbers are calculated

1. **Periods.** The window always ends at the latest month in the data (Jun 25). Monthly takes the last N months. Quarterly takes the last N calendar quarters (Jan to Mar, Apr to Jun, Jul to Sep, Oct to Dec). If fewer periods exist than asked for, all of them are shown. "All" means every period in the data.
2. **Segment and country.** Only rows matching the pick are used. "All" means every segment or every country.
3. **Loans outstanding is a balance, so a quarter shows its closing month.** Summing three month-end balances would triple-count. This is what example E1 requires (Q2 25 uses the Jun 25 balances). Revenue and operating expenses are flows, so they are summed across the months of the period. The Metrics sheet does not say which is which, so the code treats `Loans` (or any metric whose name contains "outstanding" or "balance") as a balance.
4. **Cost-to-income %.** It comes from the Metrics sheet's "Calculated from" (Expenses and Revenue) and is never read from the data. For one segment and one country it is total expenses divided by total revenue over the period. For "All" it is the **equal-weight average of the ratios of the segment-country combinations selected**.
5. **Growth views.** % change is (current minus comparison) divided by the comparison value, where the comparison is the previous period or the same period a year earlier (4 quarters or 12 months back). Periods that have nothing to compare with are left out of the window.

### Points in the brief that were unclear, and what I did

- **Cost-to-income for "All" is not defined in the brief.** Pooling all expenses over all revenue gives 53.3% for E4, but the expected answer is 57.0%. That is exactly the average of the four segments' ratios (63.0, 38.2, 87.7, 39.0), so the equal-weight average is the rule the examples require. E4 only fixes the rule across segments. **Assumption to confirm:** I applied the same rule across countries (each segment-country pair counts once) so the two filters behave the same way.
- **The Metrics sheet does not say which metrics are balances and which are flows.** I made the balance-or-flow call described in point 3 above, based on E1.
- **Blank cells are not specified.** See the next section.
- **A partly filled latest quarter is not addressed** (for example when only July has been added). See "When the data grows".

## Blank cells

Two cells are blank, and each has a note on its row: Treasury UK Expenses Jun 25 ("not closed yet") and Wealth India Revenue Feb 24 ("awaiting restatement").

**Approach: fill each blank with a simple, explainable estimate, and always show that it is an estimate.** A blank is never read as zero, because that would make a missing month look like a real collapse (for example Treasury UK cost-to-income for Q2 25 would read 24.4% instead of 36.7%).

| Metric | Blank is filled with | Edge case |
|---|---|---|
| Loans outstanding (a balance) | The previous month's balance, carried forward | No earlier month: the next real balance is copied back |
| Revenue, Operating expenses (flows) | The average of the previous and the next real month | Only one side exists (first or last month in the data): that side's value is repeated |
| Cost-to-income % | Calculated from the filled expenses and revenue | n/a |

- **Runs of blanks:** several blanks in a row use the nearest real figure on each side.
- **No basis at all:** a row with no real figure has nothing to estimate from, so it contributes nothing. If every row in a period is like that, the period has no value and the chart shows a gap.
- **Neighbours come from the whole row,** not only the months on screen. A blank on the first month of the window can still use the month before it.
- **Real data is never changed.** Filling happens only in the calculation. The Financials sheet is untouched.
- **Estimates are always visible.** The pane lists each filled cell, the value used, how it was worked out and the sheet's own note. The table has an amber dot with a tooltip. The Dashboard chart draws the affected bar or marker in amber, and the Dashboard table has a "Data check" column.

**Why these rules:** carrying a balance forward matches how a balance behaves (it changes slowly and is a position at a date). A flow in a missing month is best guessed from the months around it. Both are easy to explain and to check by hand. They are estimates, so they are flagged and should be replaced when the source data is closed or restated.

## When the data grows

The add-in is built to keep working as months are added.

- **New months are picked up automatically.** Month columns are found by their header (2025M07, 2025M08 and so on), so a new column needs no code change. The chart always ends at the latest month that has data. Edits to the Financials or Metrics sheets reload the pane by themselves, and a **Refresh data** link covers changes Excel does not report, such as inserted columns or formula results.
- **A new month is a normal period.** Monthly views show it straight away. For a quarterly view, the newest quarter is often only partly there. It is still shown, but marked as **quarter to date** in amber and described ("1 of 3 months"). A part-quarter total of revenue or expenses is never compared with a full quarter (the change shows as not comparable) and is left out of the written takeaway. Loans (a balance) and ratios are not affected by this.
- **A month column with no numbers is ignored.** A header added ahead of time does not turn into a fake latest month.
- **A month column missing between two others becomes a blank month,** so "last N months" still means calendar months. It is filled and flagged like any other blank.
- **A month that is only partly filled in** (some rows have it, some do not) uses the blank-cell rules above for the missing rows, and is flagged.
- **Layout changes are tolerated.** The header row can sit below a title, columns can be in any order, month headers are matched ignoring case and spaces, numbers stored as text (even "1,234") are read as numbers, and Excel error cells (#N/A) count as blanks. New segments, countries and metrics appear in the pickers without code changes.
- **Not covered:** a new calculated metric other than cost-to-income and operating profit would need its own rule.

## Design decisions

- **Live updates, no Apply button.** Every control change recalculates and redraws immediately, which keeps exploration fast.
- **Engine separated from Excel.** All parsing and calculation sits in `FinPulse.UI/Engine` with no Excel types. `DatasetBuilder` turns the sheet values into a plain dataset, `FinanceEngine` answers queries, and `DashboardWriter` is the only code that draws to Excel.
- **Read-only on the source sheets.** The add-in cannot damage the finance team's data. All output goes to the Dashboard sheet.
- **A native Excel chart, not an image.** Finance users can restyle it, copy it into a deck, or trace every plotted number in the table beside it.
- **Honest about data quality.** Anything estimated or incomplete is marked in the pane, the table and the chart, never hidden.

## Ideas beyond the brief

- Growth views with raw values and % change together, on a secondary axis.
- Compare by segment or country, with a ranked table and a written comparison.
- Built-in Operating profit metric.
- Same period last year, and a change badge coloured by whether the move is good or bad.
- A written takeaway for each view.
- Newest-first or oldest-first ordering, column or line chart, and a table that copies straight into Excel.
- A Dashboard data table with a "Data check" column, so the chart is fully traceable.

## Use of AI

I used AI tools as an assistant during this project: to talk through ideas and options, and to help draft parts of the code, mainly the WPF styling and boilerplate for the Excel chart formatting. 
P.S. I used AI to format the readme file