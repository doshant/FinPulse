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
    /// </summary>
    public static class DatasetBuilder
    {
        private static readonly Regex MonthHeader = new Regex(@"^\d{4}M\d{2}$", RegexOptions.Compiled);

        public static Dataset Build(object[,] financials, object[,] metrics)
        {
            var data = new Dataset();
            ReadMetrics(metrics, data);
            ReadFinancials(financials, data);
            return data;
        }

        // ------------------------------------------------------------------

        private static void ReadMetrics(object[,] a, Dataset data)
        {
            int r0 = a.GetLowerBound(0), r1 = a.GetUpperBound(0);
            int c0 = a.GetLowerBound(1), c1 = a.GetUpperBound(1);

            int cKey = FindColumn(a, r0, c0, c1, "metric");
            int cName = FindColumn(a, r0, c0, c1, "display name");
            int cShown = FindColumn(a, r0, c0, c1, "shown as");
            int cCalc = FindColumn(a, r0, c0, c1, "calculated from");
            if (cKey < 0 || cName < 0)
                throw new InvalidOperationException("The Metrics sheet needs 'Metric' and 'Display Name' columns.");

            for (int r = r0 + 1; r <= r1; r++)
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
            int r0 = a.GetLowerBound(0), r1 = a.GetUpperBound(0);
            int c0 = a.GetLowerBound(1), c1 = a.GetUpperBound(1);

            int cId = FindColumn(a, r0, c0, c1, "rowid");
            int cSeg = FindColumn(a, r0, c0, c1, "segment");
            int cCty = FindColumn(a, r0, c0, c1, "country");
            int cMet = FindColumn(a, r0, c0, c1, "metric");
            int cNotes = FindColumn(a, r0, c0, c1, "notes");
            if (cSeg < 0 || cCty < 0 || cMet < 0)
                throw new InvalidOperationException("The Financials sheet needs 'Segment', 'Country' and 'Metric' columns.");

            // Month columns are found by their header (2023M01 ...) and put in time order.
            var monthCols = new List<KeyValuePair<string, int>>();
            for (int c = c0; c <= c1; c++)
            {
                string h = Text(a[r0, c]);
                if (MonthHeader.IsMatch(h)) monthCols.Add(new KeyValuePair<string, int>(h, c));
            }
            monthCols = monthCols.OrderBy(m => m.Key, StringComparer.Ordinal).ToList();
            if (monthCols.Count == 0)
                throw new InvalidOperationException("No month columns (like 2023M01) were found on the Financials sheet.");

            data.Months.AddRange(monthCols.Select(m => m.Key));

            for (int r = r0 + 1; r <= r1; r++)
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
                    Values = new double?[monthCols.Count]
                };
                for (int i = 0; i < monthCols.Count; i++) row.Values[i] = Number(a[r, monthCols[i].Value]);
                data.Rows.Add(row);
            }
        }

        // ------------------------------------------------------------------

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
            if (o is int) return (int)o;
            string s = Convert.ToString(o, CultureInfo.InvariantCulture).Trim();
            double d;
            return s.Length > 0 && double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out d)
                ? (double?)d : null;
        }
    }
}
