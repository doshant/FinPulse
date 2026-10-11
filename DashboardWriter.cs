using FinPulse.UI;
using FinPulse.UI.Engine;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace FinPulse
{
    /// <summary>
    /// Writes the current view to the Dashboard sheet and keeps one native Excel chart in sync with it.
    /// Layout: the chart fills the first columns (A to L) and the data table sits to its right (from N).
    /// It only ever touches the Dashboard sheet. Financials and Metrics are read, never written.
    /// </summary>
    public class DashboardWriter
    {
        public const string SheetName = "Dashboard";
        private const string ChartObjectName = "FinPulseChart";
        private const int HeaderRow = 5;

        // The table starts at column N (14). Column M is a gutter between the chart and the table.
        private const int TableFirstColumn = 14;
        private const string ChartLastColumn = "L";
        private const int ChartRows = 24;
        private const int MaxTableColumns = 12;

        private static readonly int Navy = Rgb(15, 42, 74);
        private static readonly int Blue = Rgb(37, 99, 235);
        private static readonly int PaleBlue = Rgb(147, 180, 245);   // single series: every period except the latest
        private static readonly int Amber = Rgb(224, 161, 0);        // estimate or quarter to date
        private static readonly int Muted = Rgb(100, 116, 139);

        // One colour per compared segment or country. Amber is kept free for warnings.
        private static readonly int[] Palette =
        {
            Rgb(15, 42, 74), Rgb(37, 99, 235), Rgb(13, 148, 136), Rgb(124, 58, 237),
            Rgb(219, 39, 119), Rgb(100, 116, 139), Rgb(22, 163, 74), Rgb(120, 53, 15)
        };

        /// <param name="newestFirst">Latest period first (the default) or oldest first.</param>
        /// <param name="activate">Bring the Dashboard sheet to the front, so the chart is the first thing seen.</param>
        /// <returns>False when the workbook has no Dashboard sheet.</returns>
        public bool Write(Excel.Workbook wb, QueryResult result, bool lineChart, bool newestFirst, bool activate)
        {
            Excel.Worksheet ws = FindSheet(wb, SheetName);
            if (ws == null) return false;

            Excel.Application app = ws.Application;
            bool oldUpdating = app.ScreenUpdating;
            app.ScreenUpdating = false;
            try
            {
                bool compare = result.CompareSeries != null && result.CompareSeries.Count > 0;
                List<QueryResult> seriesList = compare ? result.CompareSeries : new List<QueryResult> { result };
                int k = seriesList.Count;                       // number of value columns
                int n = seriesList[0].Points.Count;             // number of periods
                MetricInfo metric = result.Metric;

                // Growth views keep the raw values: bars for the value, a line for the % change on a secondary axis.
                bool dual = !compare && metric.IsChange && result.BaseMetric != null;
                MetricInfo valueMetric = dual ? result.BaseMetric : metric;
                if (dual) k = 2;

                string numberFormat = NumberFormatFor(valueMetric);
                string changeFormat = NumberFormatFor(metric);
                string axisTitle = valueMetric.IsChange ? (valueMetric.ChangeUnit.Trim() == "pp" ? "Change (percentage points)" : "Change (%)")
                                 : valueMetric.Kind == ValueKind.Percent ? "Percent" : "USD thousands";

                // Slot j is the j-th row of the table. newestFirst puts the latest period in slot 0.
                Func<int, int> periodAt = j => newestFirst ? n - 1 - j : j;
                int latestSlot = newestFirst ? 0 : n - 1;
                PeriodPoint latest = result.Points[result.Points.Count - 1];

                string colPeriod = Col(0);
                string colLast = Col(k);                        // last value column
                string colCheck = Col(k + 1);

                // ---- Clear what this add-in owns ----

                if (Convert.ToString(ws.Range["A" + HeaderRow].Value2) == "Period")
                    ws.Range["A1:C80"].Clear();
                ws.Range[Col(0) + "1:" + Col(MaxTableColumns) + "500"].Clear();

                // ---- Title block ----
                string scope = compare
                    ? metric.DisplayName + " by " + (result.Query.CompareBy == CompareBy.Segment ? "segment" : "country") + ", " +
                      (result.Query.CompareBy == CompareBy.Segment ? CountryText(result.Query) : SegmentText(result.Query))
                    : metric.DisplayName + ", " + Scope(result.Query);

                ws.Range[colPeriod + "1"].Value2 = scope;
                ws.Range[colPeriod + "1"].Font.Bold = true;
                ws.Range[colPeriod + "1"].Font.Size = 13;
                ws.Range[colPeriod + "1"].Font.Color = Navy;

                string freqText = (result.Query.Frequency == Frequency.Quarterly ? "quarterly" : "monthly") + ", " + n + " periods";
                string unitText = metric.IsChange ? (metric.ChangeUnit.Trim() == "pp" ? "Percentage points" : "Percent change")
                                : metric.Kind == ValueKind.Percent ? "Percent" : "USD thousands";
                ws.Range[colPeriod + "2"].Value2 = dual
                    ? "Bars: " + (valueMetric.Kind == ValueKind.Percent ? "percent" : "USD thousands") + ". Line: " +
                      (metric.ChangeUnit.Trim() == "pp" ? "change in percentage points" : "% change") + " (right axis), " + freqText
                    : unitText + ", " + freqText;
                ws.Range[colPeriod + "2"].Font.Color = Muted;
                ws.Range[colPeriod + "3"].Value2 = (newestFirst ? "Newest period first. " : "Oldest period first. ") +
                                                   "The latest period (" + latest.Label + ") is highlighted.";
                ws.Range[colPeriod + "3"].Font.Color = Muted;
                ws.Range[colPeriod + "3"].Font.Italic = true;

                // ---- Data table: Period, one column per series, Data check ----
                ws.Range[colPeriod + HeaderRow].Value2 = "Period";
                if (dual)
                {
                    ws.Range[Col(1) + HeaderRow].Value2 = valueMetric.Kind == ValueKind.Percent ? "Value (%)" : "Value (USD thousands)";
                    ws.Range[Col(2) + HeaderRow].Value2 = (result.Query.Growth == GrowthView.VsPrevious ? "Change vs previous" : "Change vs last year")
                                                          + " (" + metric.ChangeUnit.Trim() + ")";
                }
                else
                {
                    for (int s = 0; s < k; s++)
                    {
                        ws.Range[Col(1 + s) + HeaderRow].Value2 = compare
                            ? seriesList[s].SeriesName
                            : (metric.IsChange ? "Change" : metric.Kind == ValueKind.Percent ? "Value (%)" : "Value (USD thousands)");
                    }
                }
                ws.Range[colCheck + HeaderRow].Value2 = "Data check";
                Excel.Range header = ws.Range[colPeriod + HeaderRow + ":" + colCheck + HeaderRow];
                header.Font.Bold = true;
                header.Font.Color = Rgb(255, 255, 255);
                header.Interior.Color = Navy;

                var values = new object[n, k + 2];
                for (int j = 0; j < n; j++)
                {
                    int pi = periodAt(j);
                    values[j, 0] = seriesList[0].Points[pi].Label;
                    var checks = new List<string>();
                    if (dual)
                    {
                        PeriodPoint p = result.Points[pi];
                        values[j, 1] = p.RawValue.HasValue ? (object)p.RawValue.Value : null;
                        values[j, 2] = p.Value.HasValue ? (object)p.Value.Value : null;
                        if (p.IsIncomplete)
                            checks.Add((p.RawValue.HasValue ? "" : "No figure. ") + p.Detail);
                    }
                    else
                    {
                        for (int s = 0; s < k; s++)
                        {
                            PeriodPoint p = seriesList[s].Points[pi];
                            values[j, 1 + s] = p.Value.HasValue ? (object)p.Value.Value : null;     // blank cell = gap in chart
                            if (p.IsIncomplete)
                                checks.Add((compare ? seriesList[s].SeriesName + ": " : "") + (p.Value.HasValue ? "" : "No figure. ") + p.Detail);
                        }
                    }
                    values[j, k + 1] = checks.Count == 0 ? "Complete" : string.Join(" | ", checks);
                }

                int first = HeaderRow + 1, last = HeaderRow + n;
                ws.Range[colPeriod + first + ":" + colPeriod + last].NumberFormat = "@";
                ws.Range[colPeriod + first + ":" + colCheck + last].Value2 = values;
                if (dual)
                {
                    ws.Range[Col(1) + first + ":" + Col(1) + last].NumberFormat = numberFormat;
                    ws.Range[Col(2) + first + ":" + Col(2) + last].NumberFormat = changeFormat;
                }
                else
                {
                    ws.Range[Col(1) + first + ":" + colLast + last].NumberFormat = numberFormat;
                }
                ws.Range[colCheck + first + ":" + colCheck + last].WrapText = true;
                ws.Range[colPeriod + first + ":" + colCheck + last].VerticalAlignment = Excel.XlVAlign.xlVAlignTop;

                Excel.Range latestRow = ws.Range[colPeriod + (first + latestSlot) + ":" + colCheck + (first + latestSlot)];
                latestRow.Font.Bold = true;
                latestRow.Interior.Color = Rgb(227, 236, 255);

                ((Excel.Range)ws.Columns["M"]).ColumnWidth = 3;
                ((Excel.Range)ws.Columns[colPeriod]).ColumnWidth = 12;
                for (int s = 0; s < k; s++) ((Excel.Range)ws.Columns[Col(1 + s)]).ColumnWidth = compare ? 14 : 22;
                ((Excel.Range)ws.Columns[colCheck]).ColumnWidth = 34;
                ((Excel.Range)ws.Range[colPeriod + first + ":" + colCheck + last].Rows).AutoFit();

                // ---- Chart in the first columns ----
                // Recreated each time, so a leftover secondary axis never survives a switch back to Value view.
                Excel.Range anchor = ws.Range["A1"];
                double chartWidth = ws.Range["A1:" + ChartLastColumn + "1"].Width;
                double chartHeight = ws.Range["A1:A" + ChartRows].Height;
                Excel.ChartObject co = FindChartObject(ws);
                if (co != null) co.Delete();
                co = ((Excel.ChartObjects)ws.ChartObjects(Type.Missing))
                    .Add(anchor.Left, anchor.Top, chartWidth, chartHeight);
                co.Name = ChartObjectName;
                co.Left = anchor.Left;
                co.Top = anchor.Top;
                co.Width = chartWidth;
                co.Height = chartHeight;

                Excel.Chart chart = co.Chart;
                chart.ChartType = lineChart ? Excel.XlChartType.xlLineMarkers : Excel.XlChartType.xlColumnClustered;
                chart.SetSourceData(ws.Range[colPeriod + HeaderRow + ":" + colLast + last], Excel.XlRowCol.xlColumns);
                chart.DisplayBlanksAs = Excel.XlDisplayBlanksAs.xlNotPlotted;
                chart.HasTitle = true;

                string latestLine;
                if (compare) latestLine = "Latest period, " + latest.Label + ", is highlighted";
                else if (dual)
                    latestLine = "Latest, " + latest.Label + ": " +
                                 (latest.RawValue.HasValue ? FinancePane.FormatValue(latest.RawValue.Value, valueMetric) : "no figure") +
                                 (latest.Value.HasValue ? " (" + FinancePane.FormatValue(latest.Value.Value, metric) + ")" : "");
                else latestLine = latest.Value.HasValue
                        ? "Latest, " + latest.Label + ": " + FinancePane.FormatValue(latest.Value.Value, metric)
                        : "Latest, " + latest.Label + ": no figure yet";
                chart.ChartTitle.Text = scope + "\n" + latestLine;
                chart.ChartTitle.Font.Size = 12;
                chart.ChartTitle.Font.Bold = true;
                chart.ChartTitle.Font.Color = Navy;

                if (compare || dual)
                {
                    chart.HasLegend = true;
                    chart.Legend.Position = Excel.XlLegendPosition.xlLegendPositionBottom;
                }
                else
                {
                    chart.HasLegend = false;
                }

                Excel.Axis valueAxis = (Excel.Axis)chart.Axes(Excel.XlAxisType.xlValue);
                valueAxis.HasTitle = true;
                valueAxis.AxisTitle.Text = axisTitle;
                valueAxis.TickLabels.NumberFormat = numberFormat;
                valueAxis.MajorGridlines.Format.Line.ForeColor.RGB = Rgb(226, 232, 240);

                if (dual)
                {
                    FormatDualSeries(chart, result, n, periodAt, latestSlot, lineChart, numberFormat, changeFormat);
                    Excel.Axis sec = (Excel.Axis)chart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlSecondary);
                    sec.HasTitle = true;
                    sec.AxisTitle.Text = metric.ChangeUnit.Trim() == "pp" ? "Change (percentage points)" : "Change (%)";
                    sec.TickLabels.NumberFormat = changeFormat;
                    sec.HasMajorGridlines = false;
                }
                else if (compare) FormatCompareSeries(chart, seriesList, n, periodAt, latestSlot, lineChart, numberFormat);
                else FormatSingleSeries(chart, seriesList[0], n, periodAt, latestSlot, lineChart, numberFormat);

                if (seriesList.Any(sr => sr.Points.Any(p => p.IsIncomplete)))
                {
                    int noteRow = last + 2;
                    ws.Range[colPeriod + noteRow].Value2 =
                        "Amber = quarter to date, or includes an estimate for a blank input. Loans repeat the previous month; revenue and expenses use the average of the months either side.";
                    ws.Range[colPeriod + noteRow].Font.Italic = true;
                    ws.Range[colPeriod + noteRow].Font.Color = Muted;
                }

                if (activate)
                {
                    ((Excel._Worksheet)ws).Activate();
                    ws.Range["A1"].Select();
                    try { app.ActiveWindow.ScrollColumn = 1; app.ActiveWindow.ScrollRow = 1; } catch { }
                }
            }
            finally
            {
                app.ScreenUpdating = oldUpdating;
            }
            return true;
        }

        // ------------------------------------------------------------------
        // One series: everything is pale except the latest period, which is dark and labelled.
        // ------------------------------------------------------------------

        private static void FormatSingleSeries(Excel.Chart chart, QueryResult data, int n, Func<int, int> periodAt,
                                               int latestSlot, bool lineChart, string numberFormat)
        {
            Excel.Series series = (Excel.Series)chart.SeriesCollection(1);
            series.HasDataLabels = true;
            Excel.DataLabels labels = (Excel.DataLabels)series.DataLabels(Type.Missing);
            labels.NumberFormat = numberFormat;
            labels.Font.Size = 9;
            labels.Font.Color = Muted;

            if (lineChart)
            {
                series.Format.Line.ForeColor.RGB = PaleBlue;
                series.Format.Line.Weight = 2.25f;
                series.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleCircle;
                series.MarkerSize = 7;
                series.MarkerBackgroundColor = PaleBlue;
                series.MarkerForegroundColor = PaleBlue;
            }
            else
            {
                series.Format.Fill.ForeColor.RGB = PaleBlue;
            }

            for (int j = 0; j < n; j++)
            {
                PeriodPoint p = data.Points[periodAt(j)];
                bool isLatest = j == latestSlot;
                if (!p.Value.HasValue) continue;
                if (!isLatest && !p.IsIncomplete) continue;

                Excel.Point pt = (Excel.Point)series.Points(j + 1);
                int color = p.IsIncomplete ? Amber : Navy;       // amber = partial data, wins over the plain highlight

                if (lineChart)
                {
                    pt.MarkerBackgroundColor = color;
                    pt.MarkerForegroundColor = isLatest ? Navy : color;
                    if (isLatest) pt.MarkerSize = 13;
                }
                else
                {
                    pt.Format.Fill.ForeColor.RGB = color;
                    if (isLatest)
                    {
                        pt.Format.Line.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                        pt.Format.Line.ForeColor.RGB = Navy;
                        pt.Format.Line.Weight = 2.5f;
                    }
                }

                if (isLatest)
                {
                    pt.HasDataLabel = true;
                    pt.DataLabel.Font.Bold = true;
                    pt.DataLabel.Font.Size = 12;
                    pt.DataLabel.Font.Color = Navy;
                }
            }
        }

        // ------------------------------------------------------------------
        // Growth view: raw value as bars (or a line) on the left axis, % change as a line on the right axis.
        // ------------------------------------------------------------------

        private static void FormatDualSeries(Excel.Chart chart, QueryResult data, int n, Func<int, int> periodAt,
                                             int latestSlot, bool lineChart, string valueFormat, string changeFormat)
        {
            Excel.Series raw = (Excel.Series)chart.SeriesCollection(1);
            Excel.Series chg = (Excel.Series)chart.SeriesCollection(2);

            raw.ChartType = lineChart ? Excel.XlChartType.xlLineMarkers : Excel.XlChartType.xlColumnClustered;
            chg.ChartType = Excel.XlChartType.xlLineMarkers;
            chg.AxisGroup = Excel.XlAxisGroup.xlSecondary;

            // ---- Raw value ----
            raw.HasDataLabels = true;
            Excel.DataLabels rl = (Excel.DataLabels)raw.DataLabels(Type.Missing);
            rl.NumberFormat = valueFormat;
            rl.Font.Size = 9;
            rl.Font.Color = Muted;

            if (lineChart)
            {
                rl.Position = Excel.XlDataLabelPosition.xlLabelPositionBelow;
                raw.Format.Line.ForeColor.RGB = Navy;
                raw.Format.Line.Weight = 2.25f;
                raw.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleCircle;
                raw.MarkerSize = 7;
                raw.MarkerBackgroundColor = Navy;
                raw.MarkerForegroundColor = Navy;
            }
            else
            {
                raw.Format.Fill.ForeColor.RGB = PaleBlue;
            }

            for (int j = 0; j < n; j++)
            {
                PeriodPoint p = data.Points[periodAt(j)];
                bool isLatest = j == latestSlot;
                if (!p.RawValue.HasValue) continue;
                if (!isLatest && !p.IsIncomplete) continue;

                Excel.Point pt = (Excel.Point)raw.Points(j + 1);
                int color = p.IsIncomplete ? Amber : Navy;

                if (lineChart)
                {
                    pt.MarkerBackgroundColor = color;
                    pt.MarkerForegroundColor = color;
                    if (isLatest) pt.MarkerSize = 12;
                }
                else
                {
                    pt.Format.Fill.ForeColor.RGB = color;
                    if (isLatest)
                    {
                        pt.Format.Line.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                        pt.Format.Line.ForeColor.RGB = Navy;
                        pt.Format.Line.Weight = 2.5f;
                    }
                }

                if (isLatest)
                {
                    pt.HasDataLabel = true;
                    pt.DataLabel.Font.Bold = true;
                    pt.DataLabel.Font.Size = 11;
                    pt.DataLabel.Font.Color = Navy;
                }
            }

            // ---- % change (secondary axis) ----
            chg.Format.Line.ForeColor.RGB = Blue;
            chg.Format.Line.Weight = 2.25f;
            if (lineChart) chg.Format.Line.DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineDash;
            chg.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleCircle;
            chg.MarkerSize = 7;
            chg.MarkerBackgroundColor = Blue;
            chg.MarkerForegroundColor = Blue;

            chg.HasDataLabels = true;
            Excel.DataLabels cl = (Excel.DataLabels)chg.DataLabels(Type.Missing);
            cl.NumberFormat = changeFormat;
            cl.Position = Excel.XlDataLabelPosition.xlLabelPositionAbove;
            cl.Font.Size = 9;
            cl.Font.Bold = true;
            cl.Font.Color = Blue;

            for (int j = 0; j < n; j++)
            {
                PeriodPoint p = data.Points[periodAt(j)];
                if (!p.Value.HasValue || !p.IsIncomplete) continue;
                Excel.Point pt = (Excel.Point)chg.Points(j + 1);
                pt.MarkerBackgroundColor = Amber;
                pt.MarkerForegroundColor = Amber;
            }
        }

        // ------------------------------------------------------------------
        // Several series (one per segment or country): each has its own colour. Labels appear only on the
        // latest period, which is also drawn larger (lines) or outlined (columns).
        // Amber still marks any point that is an estimate or a quarter to date.
        // ------------------------------------------------------------------

        private static void FormatCompareSeries(Excel.Chart chart, List<QueryResult> seriesList, int n, Func<int, int> periodAt,
                                                int latestSlot, bool lineChart, string numberFormat)
        {
            for (int s = 0; s < seriesList.Count; s++)
            {
                Excel.Series series = (Excel.Series)chart.SeriesCollection(s + 1);
                int color = Palette[s % Palette.Length];
                series.HasDataLabels = false;

                if (lineChart)
                {
                    series.Format.Line.ForeColor.RGB = color;
                    series.Format.Line.Weight = 2.25f;
                    series.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleCircle;
                    series.MarkerSize = 6;
                    series.MarkerBackgroundColor = color;
                    series.MarkerForegroundColor = color;
                }
                else
                {
                    series.Format.Fill.ForeColor.RGB = color;
                }

                for (int j = 0; j < n; j++)
                {
                    PeriodPoint p = seriesList[s].Points[periodAt(j)];
                    bool isLatest = j == latestSlot;
                    if (!p.Value.HasValue) continue;
                    if (!isLatest && !p.IsIncomplete) continue;

                    Excel.Point pt = (Excel.Point)series.Points(j + 1);

                    if (lineChart)
                    {
                        if (isLatest) pt.MarkerSize = 12;
                        if (p.IsIncomplete) { pt.MarkerBackgroundColor = Amber; pt.MarkerForegroundColor = Amber; }
                    }
                    else if (p.IsIncomplete || isLatest)
                    {
                        pt.Format.Line.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                        pt.Format.Line.ForeColor.RGB = p.IsIncomplete ? Amber : Navy;
                        pt.Format.Line.Weight = 2.5f;
                    }

                    if (isLatest)
                    {
                        pt.HasDataLabel = true;
                        pt.DataLabel.NumberFormat = numberFormat;
                        pt.DataLabel.Font.Bold = true;
                        pt.DataLabel.Font.Size = 10;
                        pt.DataLabel.Font.Color = color;
                    }
                }
            }
        }

        // ------------------------------------------------------------------

        private static string NumberFormatFor(MetricInfo metric)
        {
            if (metric.IsChange)
            {
                string unit = metric.ChangeUnit.Trim() == "pp" ? " pp" : "%";
                return "+0.0\"" + unit + "\";-0.0\"" + unit + "\";0.0\"" + unit + "\"";
            }
            return metric.Kind == ValueKind.Percent ? "0.0\"%\"" : "#,##0";
        }

        private static string Scope(Query q)
        {
            return SegmentText(q) + ", " + CountryText(q);
        }

        private static string SegmentText(Query q) { return q.Segment == "All" ? "all segments" : q.Segment; }

        private static string CountryText(Query q) { return q.Country == "All" ? "all countries" : q.Country; }

        /// <summary>Column letter for the table: offset 0 is column N.</summary>
        private static string Col(int offset)
        {
            int index = TableFirstColumn + offset;       // 1-based column number
            string name = string.Empty;
            while (index > 0)
            {
                int rem = (index - 1) % 26;
                name = (char)('A' + rem) + name;
                index = (index - 1) / 26;
            }
            return name;
        }

        private static Excel.ChartObject FindChartObject(Excel.Worksheet ws)
        {
            Excel.ChartObjects all = (Excel.ChartObjects)ws.ChartObjects(Type.Missing);
            foreach (Excel.ChartObject c in all)
                if (c.Name == ChartObjectName) return c;
            return null;
        }

        public static Excel.Worksheet FindSheet(Excel.Workbook wb, string name)
        {
            foreach (Excel.Worksheet ws in wb.Worksheets)
                if (string.Equals(ws.Name, name, StringComparison.OrdinalIgnoreCase)) return ws;
            return null;
        }

        private static int Rgb(int r, int g, int b) { return r | (g << 8) | (b << 16); }
    }
}