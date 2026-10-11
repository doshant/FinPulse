using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace FinPulse.UI.Engine
{
    public enum Frequency { Monthly, Quarterly }

    public enum ValueKind { Amount, Percent }

    /// <summary>One row of the Financials sheet. A null month value means the cell was blank.</summary>
    public class FinancialRow
    {
        public string RowId;
        public string Segment;
        public string Country;
        public string Metric;
        public double?[] Values;      // index matches Dataset.Months
        public string Notes;
    }

    public class MetricInfo
    {
        public string Key;            // Loans, Revenue, Expenses, CostToIncome
        public string DisplayName;
        public ValueKind Kind;
        public bool IsCalculated;     // true when "Calculated from" is filled in
        public bool IsStock;          // balance (Loans): a quarter shows its closing month, not a sum
        public bool HigherIsBetter;
        public bool IsChange;         // a growth view: the numbers are % (or percentage-point) changes, not values
        public string ChangeUnit;     // "%" or " pp"
    }

    /// <summary>Everything read from the workbook, kept free of any Excel types so it can be unit tested.</summary>
    public class Dataset
    {
        public List<string> Months = new List<string>();      // "2023M01" ... in order
        public List<FinancialRow> Rows = new List<FinancialRow>();
        public List<MetricInfo> Metrics = new List<MetricInfo>();

        public IEnumerable<string> Segments { get { return Rows.Select(r => r.Segment).Distinct().OrderBy(s => s); } }
        public IEnumerable<string> Countries { get { return Rows.Select(r => r.Country).Distinct().OrderBy(s => s); } }
    }

    /// <summary>Draw one series per segment or per country instead of a single combined series.</summary>
    public enum CompareBy { None, Segment, Country }

    /// <summary>Show the values themselves, or how much they changed.</summary>
    public enum GrowthView { Value, VsPrevious, VsYearAgo }

    public class Query
    {
        public string MetricKey;
        public string Segment = "All";
        public string Country = "All";
        public Frequency Frequency = Frequency.Quarterly;
        public int Periods = 4;
        public CompareBy CompareBy = CompareBy.None;
        public GrowthView Growth = GrowthView.Value;
    }

    /// <summary>A blank cell that fell inside a displayed period, and what was done about it.</summary>
    public class DataIssue
    {
        public string RowId;
        public string Segment;
        public string Country;
        public string Metric;
        public string Month;          // e.g. "2025M06"
        public string Note;           // text from the Notes column, if any
        public double? FilledWith;    // the estimate used, or null when there was nothing to base one on
        public string Method;         // plain-English reason, e.g. "average of Jan 24 and Mar 24"
    }

    public class PeriodPoint
    {
        public string Label;          // "Q2 25" or "Jun 25"
        public string Key;            // "2025Q2" or "2025M06"
        public double? Value;         // null = nothing to show for this period
        public double? RawValue;      // growth views: the underlying value for this period (Value holds the change)
        public bool IsIncomplete;     // true when the period contains an estimate for a blank cell (or has no figure)
        public bool IsPartial;        // a quarter with fewer than 3 months in the data (quarter to date)
        public string Detail;         // explanation shown in tooltips and the pane
    }

    public class QueryResult
    {
        public Query Query;
        public MetricInfo Metric;
        public MetricInfo BaseMetric;             // growth views: the metric the change was calculated from (null in Value view)
        public List<PeriodPoint> Points = new List<PeriodPoint>();
        public List<DataIssue> Issues = new List<DataIssue>();
        public int RowsUsed;
        public string SeriesName;                 // used when comparing, for example "Retail"
        public List<QueryResult> CompareSeries;   // one result per segment or country when comparing, otherwise null
    }

    public static class FinanceEngine
    {
        public static readonly CultureInfo Inv = CultureInfo.InvariantCulture;
        private static readonly string[] MonthNames =
            { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" };

        // ------------------------------------------------------------------
        // Public API
        // ------------------------------------------------------------------

        public const string OperatingProfitKey = "OperatingProfit";

        /// <summary>
        /// Runs a query. When q.CompareBy is set, the returned result is the overall series (with the compared
        /// dimension set to All) and CompareSeries holds one result per segment or country.
        /// </summary>
        public static QueryResult Run(Dataset data, Query q)
        {
            Query overallQuery = Clone(q);
            if (q.CompareBy == CompareBy.Segment) overallQuery.Segment = "All";
            if (q.CompareBy == CompareBy.Country) overallQuery.Country = "All";

            QueryResult overall = RunOne(data, overallQuery);
            overall.SeriesName = "Overall";

            if (q.CompareBy != CompareBy.None)
            {
                IEnumerable<string> items = q.CompareBy == CompareBy.Segment ? data.Segments : data.Countries;
                overall.CompareSeries = new List<QueryResult>();
                foreach (string item in items.Take(8))      // more than 8 lines is unreadable on a chart
                {
                    Query sq = Clone(q);
                    sq.CompareBy = CompareBy.None;
                    if (q.CompareBy == CompareBy.Segment) sq.Segment = item; else sq.Country = item;

                    QueryResult one = RunOne(data, sq);
                    one.SeriesName = item;
                    if (one.Points.Any(p => p.Value.HasValue)) overall.CompareSeries.Add(one);
                }
            }
            return overall;
        }

        private static Query Clone(Query q)
        {
            return new Query
            {
                MetricKey = q.MetricKey,
                Segment = q.Segment,
                Country = q.Country,
                Frequency = q.Frequency,
                Periods = q.Periods,
                CompareBy = q.CompareBy,
                Growth = q.Growth
            };
        }

        // ------------------------------------------------------------------
        // Growth view: % change vs the previous period or the same period last year.
        // Ratios (cost-to-income) change in percentage points. The first periods, which have nothing to compare
        // with, are left out. A quarter-to-date total is never compared with a full quarter.
        // The underlying value of each period is kept in RawValue, so the chart can show it beside the change.
        // ------------------------------------------------------------------

        private static QueryResult RunOne(Dataset data, Query q)
        {
            if (q.Growth == GrowthView.Value) return RunBase(data, q);

            int shift = q.Growth == GrowthView.VsPrevious ? 1 : (q.Frequency == Frequency.Monthly ? 12 : 4);
            Query bq = Clone(q);
            bq.Growth = GrowthView.Value;
            bq.CompareBy = CompareBy.None;
            bq.Periods = q.Periods + shift;
            QueryResult b = RunBase(data, bq);

            MetricInfo baseMetric = b.Metric;
            bool ratio = baseMetric.Kind == ValueKind.Percent;
            bool flowTotal = baseMetric.Kind == ValueKind.Amount && !baseMetric.IsStock;

            var res = new QueryResult
            {
                Query = q,
                RowsUsed = b.RowsUsed,
                Metric = new MetricInfo
                {
                    Key = baseMetric.Key + "|" + q.Growth,
                    DisplayName = baseMetric.DisplayName + (q.Growth == GrowthView.VsPrevious
                        ? ", change vs previous period" : ", change vs same period last year"),
                    Kind = ValueKind.Percent,
                    IsCalculated = true,
                    IsChange = true,
                    ChangeUnit = ratio ? " pp" : "%",
                    HigherIsBetter = baseMetric.HigherIsBetter
                }
            };
            res.Issues.AddRange(b.Issues);
            res.BaseMetric = baseMetric;

            var list = new List<PeriodPoint>();
            for (int i = shift; i < b.Points.Count; i++)
            {
                PeriodPoint cur = b.Points[i], prv = b.Points[i - shift];
                var p = new PeriodPoint
                {
                    Label = cur.Label,
                    Key = cur.Key,
                    RawValue = cur.Value,
                    IsPartial = cur.IsPartial,
                    IsIncomplete = cur.IsIncomplete || prv.IsIncomplete
                };

                var notes = new List<string>();
                if (!string.IsNullOrEmpty(cur.Detail)) notes.Add(cur.Detail);
                if (!string.IsNullOrEmpty(prv.Detail)) notes.Add("Compared period, " + prv.Label + ": " + prv.Detail);

                if (flowTotal && (cur.IsPartial || prv.IsPartial))
                {
                    notes.Add("A quarter to date cannot be compared with a full quarter.");
                }
                else if (cur.Value.HasValue && prv.Value.HasValue)
                {
                    if (ratio) p.Value = cur.Value.Value - prv.Value.Value;
                    else if (Math.Abs(prv.Value.Value) > 1e-9) p.Value = (cur.Value.Value - prv.Value.Value) / Math.Abs(prv.Value.Value) * 100.0;
                }

                p.Detail = notes.Count == 0 ? null : string.Join(" ", notes);
                list.Add(p);
            }

            res.Points = list.Skip(Math.Max(0, list.Count - q.Periods)).ToList();
            return res;
        }

        // ------------------------------------------------------------------
        // Values for one series
        // ------------------------------------------------------------------

        private static QueryResult RunBase(Dataset data, Query q)
        {
            MetricInfo metric = data.Metrics.First(m => m.Key == q.MetricKey);
            var result = new QueryResult { Query = q, Metric = metric };

            List<List<int>> periods = BuildPeriods(data.Months, q.Frequency, q.Periods);

            IEnumerable<FinancialRow> scope = data.Rows.Where(r =>
                (q.Segment == "All" || r.Segment == q.Segment) &&
                (q.Country == "All" || r.Country == q.Country));

            if (metric.Key == OperatingProfitKey)
            {
                RunOperatingProfit(data, scope.ToList(), periods, q.Frequency, result);
            }
            else if (metric.IsCalculated)
            {
                RunCostToIncome(data, scope.ToList(), periods, q.Frequency, result);
            }
            else
            {
                List<FinancialRow> rows = scope.Where(r => r.Metric == metric.Key).ToList();
                result.RowsUsed = rows.Count;
                RunSimple(data, rows, periods, q.Frequency, metric, result);
            }
            return result;
        }

        // ------------------------------------------------------------------
        // Periods: always end at the latest month in the data
        // ------------------------------------------------------------------

        /// <summary>Returns, oldest first, the month indexes that make up each requested period.</summary>
        public static List<List<int>> BuildPeriods(IList<string> months, Frequency freq, int count)
        {
            int last = months.Count - 1;
            var periods = new List<List<int>>();

            if (freq == Frequency.Monthly)
            {
                for (int i = Math.Max(0, last - count + 1); i <= last; i++)
                    periods.Add(new List<int> { i });
                return periods;
            }

            // Quarterly: calendar quarters (Jan-Mar, Apr-Jun, Jul-Sep, Oct-Dec).
            // The newest quarter is the one containing the latest month. If that month is not the quarter's
            // last month, the quarter is partial and only the months that exist are used.
            var byQuarter = new List<KeyValuePair<string, List<int>>>();
            for (int i = 0; i <= last; i++)
            {
                string key = QuarterKey(months[i]);
                if (byQuarter.Count == 0 || byQuarter[byQuarter.Count - 1].Key != key)
                    byQuarter.Add(new KeyValuePair<string, List<int>>(key, new List<int>()));
                byQuarter[byQuarter.Count - 1].Value.Add(i);
            }

            int start = Math.Max(0, byQuarter.Count - count);
            for (int i = start; i < byQuarter.Count; i++) periods.Add(byQuarter[i].Value);
            return periods;
        }

        public static string QuarterKey(string month)
        {
            int year = int.Parse(month.Substring(0, 4), Inv);
            int m = int.Parse(month.Substring(5, 2), Inv);
            return year.ToString(Inv) + "Q" + ((m - 1) / 3 + 1).ToString(Inv);
        }

        public static string PeriodLabel(string firstMonth, Frequency freq)
        {
            int year = int.Parse(firstMonth.Substring(0, 4), Inv);
            int m = int.Parse(firstMonth.Substring(5, 2), Inv);
            string yy = (year % 100).ToString("00", Inv);
            return freq == Frequency.Quarterly
                ? "Q" + ((m - 1) / 3 + 1).ToString(Inv) + " " + yy
                : MonthNames[m - 1] + " " + yy;
        }

        // ------------------------------------------------------------------
        // Filling blanks
        //
        // A blank cell is replaced by an estimate so totals stay comparable. Every estimate is recorded and
        // flagged, so the user can always see which figures are real and which are filled in.
        //
        //   Loans (a balance)   : the previous month's balance carried forward.
        //                         Edge case, nothing earlier exists: the next real balance is copied back.
        //   Revenue, Expenses   : the average of the previous and the next real month.
        //                         Edge cases, only one side exists (first or last month in the data):
        //                         that side's value is repeated.
        //   Several blanks in a row use the nearest real figure on each side.
        //   A row with no real figure at all has nothing to base an estimate on, so it gives no value.
        //
        // Neighbours are looked up in the whole row, not just the months on screen.
        // ------------------------------------------------------------------

        private struct Resolved
        {
            public double? Value;
            public bool Estimated;
            public string Method;
        }

        private static Resolved Resolve(Dataset data, FinancialRow r, int i, bool isBalance)
        {
            double? raw = r.Values[i];
            if (raw.HasValue) return new Resolved { Value = raw };

            int prev = -1, next = -1;
            for (int k = i - 1; k >= 0; k--) if (r.Values[k].HasValue) { prev = k; break; }
            for (int k = i + 1; k < r.Values.Length; k++) if (r.Values[k].HasValue) { next = k; break; }

            if (isBalance)
            {
                if (prev >= 0)
                    return Estimate(r.Values[prev].Value, "carried forward from " + MonthText(data, prev));
                if (next >= 0)
                    return Estimate(r.Values[next].Value, "no earlier month, so copied back from " + MonthText(data, next));
            }
            else
            {
                if (prev >= 0 && next >= 0)
                    return Estimate((r.Values[prev].Value + r.Values[next].Value) / 2.0,
                                    "average of " + MonthText(data, prev) + " and " + MonthText(data, next));
                if (prev >= 0)
                    return Estimate(r.Values[prev].Value, "no later month, so repeated from " + MonthText(data, prev));
                if (next >= 0)
                    return Estimate(r.Values[next].Value, "no earlier month, so repeated from " + MonthText(data, next));
            }

            return new Resolved { Value = null, Estimated = false, Method = "no figures at all in this row" };
        }

        private static Resolved Estimate(double? value, string method)
        {
            return new Resolved { Value = value, Estimated = true, Method = method };
        }

        private static string MonthText(Dataset data, int index)
        {
            return PeriodLabel(data.Months[index], Frequency.Monthly);
        }

        private static DataIssue Issue(Dataset data, FinancialRow r, int monthIndex, Resolved res)
        {
            return new DataIssue
            {
                RowId = r.RowId,
                Segment = r.Segment,
                Country = r.Country,
                Metric = r.Metric,
                Month = data.Months[monthIndex],
                Note = r.Notes,
                FilledWith = res.Value,
                Method = res.Method
            };
        }

        // ------------------------------------------------------------------
        // Loans / Revenue / Expenses
        // ------------------------------------------------------------------

        private static void RunSimple(Dataset data, List<FinancialRow> rows, List<List<int>> periods,
                                      Frequency freq, MetricInfo metric, QueryResult result)
        {
            foreach (List<int> months in periods)
            {
                var point = NewPoint(data, months, freq);
                double total = 0;
                int contributions = 0;
                var issues = new List<DataIssue>();

                foreach (FinancialRow r in rows)
                {
                    if (metric.IsStock)
                    {
                        // A balance is a position at a date. A quarter shows its closing month.
                        int closing = months[months.Count - 1];
                        Resolved res = Resolve(data, r, closing, true);
                        if (res.Value.HasValue) { total += res.Value.Value; contributions++; }
                        if (!r.Values[closing].HasValue) issues.Add(Issue(data, r, closing, res));
                    }
                    else
                    {
                        // A flow adds up over the months of the period.
                        bool any = false;
                        foreach (int mi in months)
                        {
                            Resolved res = Resolve(data, r, mi, false);
                            if (res.Value.HasValue) { total += res.Value.Value; any = true; }
                            if (!r.Values[mi].HasValue) issues.Add(Issue(data, r, mi, res));
                        }
                        if (any) contributions++;
                    }
                }

                point.Value = contributions > 0 ? (double?)total : null;
                Flag(point, issues, result);
                result.Points.Add(point);
            }
        }

        // ------------------------------------------------------------------
        // Operating profit = revenue minus operating expenses, each with blanks filled in as above.
        // ------------------------------------------------------------------

        private static void RunOperatingProfit(Dataset data, List<FinancialRow> scope, List<List<int>> periods,
                                               Frequency freq, QueryResult result)
        {
            MetricInfo revMetric = data.Metrics.First(m => m.Key == "Revenue");
            MetricInfo expMetric = data.Metrics.First(m => m.Key == "Expenses");
            List<FinancialRow> revRows = scope.Where(r => r.Metric == "Revenue").ToList();
            List<FinancialRow> expRows = scope.Where(r => r.Metric == "Expenses").ToList();
            result.RowsUsed = revRows.Count + expRows.Count;

            var rev = new QueryResult { Metric = revMetric };
            var exp = new QueryResult { Metric = expMetric };
            RunSimple(data, revRows, periods, freq, revMetric, rev);
            RunSimple(data, expRows, periods, freq, expMetric, exp);

            for (int i = 0; i < rev.Points.Count; i++)
            {
                PeriodPoint a = rev.Points[i], b = exp.Points[i];
                var p = new PeriodPoint
                {
                    Label = a.Label,
                    Key = a.Key,
                    IsPartial = a.IsPartial || b.IsPartial,
                    IsIncomplete = a.IsIncomplete || b.IsIncomplete,
                    Value = a.Value.HasValue && b.Value.HasValue ? (double?)(a.Value.Value - b.Value.Value) : null
                };
                var notes = new[] { a.Detail, b.Detail }.Where(d => !string.IsNullOrEmpty(d)).Distinct().ToList();
                p.Detail = notes.Count == 0 ? null : string.Join(" ", notes);
                result.Points.Add(p);
            }
            result.Issues.AddRange(rev.Issues);
            result.Issues.AddRange(exp.Issues);
        }

        // ------------------------------------------------------------------
        // Cost-to-income %
        //
        // Each (segment, country) group gets its own ratio: expenses / revenue over the period, using the
        // filled-in figures. "All" for segment or country averages those group ratios with equal weight.
        // (This is the rule that reproduces worked example E4.)
        // ------------------------------------------------------------------

        private static void RunCostToIncome(Dataset data, List<FinancialRow> scope, List<List<int>> periods,
                                            Frequency freq, QueryResult result)
        {
            var groups = scope.GroupBy(r => r.Segment + "|" + r.Country).ToList();
            result.RowsUsed = scope.Count(r => r.Metric == "Expenses" || r.Metric == "Revenue");

            foreach (List<int> months in periods)
            {
                var point = NewPoint(data, months, freq);
                var ratios = new List<double>();
                var issues = new List<DataIssue>();

                foreach (var g in groups)
                {
                    FinancialRow exp = g.FirstOrDefault(r => r.Metric == "Expenses");
                    FinancialRow rev = g.FirstOrDefault(r => r.Metric == "Revenue");
                    if (exp == null || rev == null) continue;

                    double e = 0, v = 0;
                    bool usable = false;
                    foreach (int mi in months)
                    {
                        Resolved re = Resolve(data, exp, mi, false);
                        Resolved rr = Resolve(data, rev, mi, false);
                        if (!exp.Values[mi].HasValue) issues.Add(Issue(data, exp, mi, re));
                        if (!rev.Values[mi].HasValue) issues.Add(Issue(data, rev, mi, rr));
                        if (re.Value.HasValue && rr.Value.HasValue) { e += re.Value.Value; v += rr.Value.Value; usable = true; }
                    }
                    if (usable && Math.Abs(v) > 1e-12) ratios.Add(e / v * 100.0);
                }

                point.Value = ratios.Count > 0 ? (double?)ratios.Average() : null;
                Flag(point, issues, result);
                result.Points.Add(point);
            }
        }

        // ------------------------------------------------------------------
        // Helpers
        // ------------------------------------------------------------------

        private static PeriodPoint NewPoint(Dataset data, List<int> months, Frequency freq)
        {
            string first = data.Months[months[0]];
            var point = new PeriodPoint
            {
                Label = PeriodLabel(first, freq),
                Key = freq == Frequency.Quarterly ? QuarterKey(first) : first
            };

            // New months arrive over time, so the newest quarter is often only partly there (for example only July).
            // It is still shown, but marked, so a one-month total is never mistaken for a full quarter.
            if (freq == Frequency.Quarterly && months.Count < 3)
            {
                point.IsPartial = true;
                point.IsIncomplete = true;
                string span = months.Count == 1
                    ? MonthText(data.Months[months[0]])
                    : MonthText(data.Months[months[0]]) + " to " + MonthText(data.Months[months[months.Count - 1]]);
                point.Detail = "Quarter to date: only " + span + " (" + months.Count + " of 3 months) is in the data.";
            }
            return point;
        }

        private static void Flag(PeriodPoint point, List<DataIssue> issues, QueryResult result)
        {
            if (issues.Count == 0) return;
            point.IsIncomplete = true;

            var parts = issues
                .Select(i => i.Segment + " " + i.Country + " " + i.Metric + " " + MonthText(i.Month) +
                             (i.FilledWith.HasValue
                                 ? " filled with " + Math.Round(i.FilledWith.Value, MidpointRounding.AwayFromZero).ToString("N0", Inv) + " (" + i.Method + ")"
                                 : " has no figure (" + i.Method + ")") +
                             (string.IsNullOrWhiteSpace(i.Note) ? string.Empty : " [" + i.Note.Trim() + "]"))
                .Distinct().Take(3).ToList();
            string estimated = "Estimated: " + string.Join("; ", parts) + (issues.Count > 3 ? "; and more." : ".");
            point.Detail = string.IsNullOrEmpty(point.Detail) ? estimated : point.Detail + " " + estimated;

            result.Issues.AddRange(issues);
        }

        private static string MonthText(string month)
        {
            return PeriodLabel(month, Frequency.Monthly);
        }

        // ------------------------------------------------------------------
        // Comparisons used by the pane
        // ------------------------------------------------------------------

        /// <summary>The same period one year earlier, or null when the data does not reach back that far.</summary>
        public static double? SamePeriodLastYear(Dataset data, Query q)
        {
            int shift = q.Frequency == Frequency.Monthly ? 12 : 4;
            var longer = new Query
            {
                MetricKey = q.MetricKey,
                Segment = q.Segment,
                Country = q.Country,
                Frequency = q.Frequency,
                Periods = q.Periods + shift
            };
            QueryResult r = Run(data, longer);
            if (r.Points.Count < shift + 1) return null;
            return r.Points[r.Points.Count - 1 - shift].Value;
        }
    }
}