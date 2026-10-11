using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace FinPulse.UI.Engine
{
    /// <summary>
    /// Turns the values of the Financials and Metrics sheets into a Dataset.
    /// Takes plain object[,] arrays (the shape Excel's Range.Value2 returns, 1-based) so no Excel
    /// reference is needed and the add-in never writes to those sheets.
    ///
    /// Built to cope with the sheet growing:
    ///  - New month columns (2025M07, 2025M08 ...) are found by their header, wherever they are added.
    ///  - The month axis is continuous. A month column that is missing between two others becomes a blank month.
    ///  - A month column with no numbers at all (for example a header added ahead of time) is ignored, so the
    ///    chart still ends at the latest month that really has data.
    ///  - The header row does not have to be the first row, and columns can be in any order.
    ///  - Numbers stored as text (including "1,234") are read as numbers. Other text such as "n/a" is a blank.
    /// </summary>
    public static class DatasetBuilder
    {
        private static readonly Regex MonthHeader =
            new Regex(@"^\s*(\d{4})\s*[Mm]\s*(\d{1,2})\s*$", RegexOptions.Compiled);

        public static Dataset Build(object[,] financials, object[,] metrics)
        {
            var data = new Dataset();
            ReadMetrics(metrics, data);
            AddBuiltInMetrics(data);
            ReadFinancials(financials, data);
            return data;
        }

        /// <summary>
        /// Metrics the add-in calculates itself. They are added next to the ones on the Metrics sheet,
        /// which is never changed.
        /// </summary>
        private static void AddBuiltInMetrics(Dataset data)
        {
            bool hasRevenue = data.Metrics.Any(m => m.Key.Equals("Revenue", StringComparison.OrdinalIgnoreCase));
            bool hasExpenses = data.Metrics.Any(m => m.Key.Equals("Expenses", StringComparison.OrdinalIgnoreCase));
            if (hasRevenue && hasExpenses && !data.Metrics.Any(m => m.Key == FinanceEngine.OperatingProfitKey))
            {
                data.Metrics.Add(new MetricInfo
                {
                    Key = FinanceEngine.OperatingProfitKey,
                    DisplayName = "Operating profit",
                    Kind = ValueKind.Amount,
                    IsCalculated = true,
                    HigherIsBetter = true
                });
            }
        }

        // ------------------------------------------------------------------

        private static void ReadMetrics(object[,] a, Dataset data)
        {
            int c0 = a.GetLowerBound(1), c1 = a.GetUpperBound(1);
            int r1 = a.GetUpperBound(0);

            int hr = FindHeaderRow(a, "Metrics", "metric", "display name");

            int cKey = FindColumn(a, hr, c0, c1, "metric");
            int cName = FindColumn(a, hr, c0, c1, "display name");
            int cShown = FindColumn(a, hr, c0, c1, "shown as");
            int cCalc = FindColumn(a, hr, c0, c1, "calculated from");

            for (int r = hr + 1; r <= r1; r++)
            {
                string key = Text(a[r, cKey]);
                if (key.Length == 0) continue;

                string name = Text(a[r, cName]);
                string shown = cShown >= 0 ? Text(a[r, cShown]) : string.Empty;
                string calc = cCalc >= 0 ? Text(a[r, cCalc]) : string.Empty;

                data.Metrics.Add(new MetricInfo
                {
                    Key = key,
                    DisplayName = name.Length > 0 ? name : key,
                    Kind = shown.IndexOf("percent", StringComparison.OrdinalIgnoreCase) >= 0 ? ValueKind.Percent : ValueKind.Amount,
                    IsCalculated = calc.Length > 0,
                    // The sheet does not say which metrics are balances. Loans outstanding is the only one here.
                    IsStock = key.Equals("Loans", StringComparison.OrdinalIgnoreCase) ||
                              name.IndexOf("outstanding", StringComparison.OrdinalIgnoreCase) >= 0 ||
                              name.IndexOf("balance", StringComparison.OrdinalIgnoreCase) >= 0,
                    // Costs and the cost-to-income ratio are better when they fall.
                    HigherIsBetter = !(key.IndexOf("expense", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                       key.IndexOf("cost", StringComparison.OrdinalIgnoreCase) >= 0)
                });
            }
        }

        private static void ReadFinancials(object[,] a, Dataset data)
        {
            int c0 = a.GetLowerBound(1), c1 = a.GetUpperBound(1);
            int r1 = a.GetUpperBound(0);

            int hr = FindHeaderRow(a, "Financials", "segment", "country", "metric");

            int cId = FindColumn(a, hr, c0, c1, "rowid");
            int cSeg = FindColumn(a, hr, c0, c1, "segment");
            int cCty = FindColumn(a, hr, c0, c1, "country");
            int cMet = FindColumn(a, hr, c0, c1, "metric");
            int cNotes = FindColumn(a, hr, c0, c1, "notes");

            // ---- Month columns, found by header. The first column wins if a month appears twice. ----
            var columnOf = new Dictionary<int, int>();          // month number (year*12 + month-1) -> sheet column
            for (int c = c0; c <= c1; c++)
            {
                Match m = MonthHeader.Match(Text(a[hr, c]));
                if (!m.Success) continue;
                int year = int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture);
                int month = int.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture);
                if (month < 1 || month > 12) continue;
                int n = year * 12 + (month - 1);
                if (!columnOf.ContainsKey(n)) columnOf[n] = c;
            }
            if (columnOf.Count == 0)
                throw new InvalidOperationException("No month columns (like 2023M01) were found on the Financials sheet.");

            // ---- Continuous month axis from the earliest to the latest header ----
            int first = columnOf.Keys.Min(), last = columnOf.Keys.Max();
            int span = last - first + 1;

            var rows = new List<FinancialRow>();
            for (int r = hr + 1; r <= r1; r++)
            {
                string metric = Text(a[r, cMet]);
                if (metric.Length == 0) continue;

                var row = new FinancialRow
                {
                    RowId = cId >= 0 ? Text(a[r, cId]) : "row " + r,
                    Segment = Text(a[r, cSeg]),
                    Country = Text(a[r, cCty]),
                    Metric = metric,
                    Notes = cNotes >= 0 ? Text(a[r, cNotes]) : string.Empty,
                    Values = new double?[span]
                };
                for (int i = 0; i < span; i++)
                {
                    int col;
                    row.Values[i] = columnOf.TryGetValue(first + i, out col) ? Number(a[r, col]) : null;
                }
                rows.Add(row);
            }

            // ---- Ignore month columns with no numbers at all at either end (for example a header added ahead of time) ----
            int firstWithData = -1, lastWithData = -1;
            for (int i = 0; i < span; i++)
            {
                if (!rows.Any(rw => rw.Values[i].HasValue)) continue;
                if (firstWithData < 0) firstWithData = i;
                lastWithData = i;
            }
            if (firstWithData < 0)
                throw new InvalidOperationException("The Financials sheet has month columns but no numbers in them.");

            int keep = lastWithData - firstWithData + 1;
            foreach (FinancialRow row in rows)
            {
                var trimmed = new double?[keep];
                Array.Copy(row.Values, firstWithData, trimmed, 0, keep);
                row.Values = trimmed;
            }

            for (int i = firstWithData; i <= lastWithData; i++)
            {
                int n = first + i;
                data.Months.Add((n / 12).ToString("0000", CultureInfo.InvariantCulture) + "M" +
                                (n % 12 + 1).ToString("00", CultureInfo.InvariantCulture));
            }
            data.Rows.AddRange(rows);
        }

        // ------------------------------------------------------------------

        /// <summary>Finds the header row: the first of the top rows that contains every required column name.</summary>
        private static int FindHeaderRow(object[,] a, string sheetName, params string[] required)
        {
            int r0 = a.GetLowerBound(0), r1 = a.GetUpperBound(0);
            int c0 = a.GetLowerBound(1), c1 = a.GetUpperBound(1);

            for (int r = r0; r <= Math.Min(r1, r0 + 24); r++)
            {
                if (required.All(name => FindColumn(a, r, c0, c1, name) >= 0)) return r;
            }
            throw new InvalidOperationException("The " + sheetName + " sheet needs these column headers: " +
                                                string.Join(", ", required) + ".");
        }

        private static int FindColumn(object[,] a, int headerRow, int c0, int c1, string name)
        {
            for (int c = c0; c <= c1; c++)
            {
                string h = Text(a[headerRow, c]).TrimStart('*', ' ');
                if (h.Equals(name, StringComparison.OrdinalIgnoreCase)) return c;
            }
            return -1;
        }

        private static string Text(object o)
        {
            return o == null ? string.Empty : Convert.ToString(o, CultureInfo.InvariantCulture).Trim();
        }

        /// <summary>Blank or non-numeric cells become null. They are never read as zero.</summary>
        private static double? Number(object o)
        {
            if (o == null) return null;
            if (o is double) return (double)o;
            if (o is int) return null;     // Excel hands back error cells (#N/A, #DIV/0!) as ints. They are blanks, not numbers.
            string s = Convert.ToString(o, CultureInfo.InvariantCulture).Trim();
            double d;
            return s.Length > 0 &&
                   double.TryParse(s, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out d)
                ? (double?)d : null;
        }
    }
}