using System;
using System.Collections.Generic;
using System.Linq;
using FinPulse.UI;
using FinPulse.UI.Engine;
using Excel = Microsoft.Office.Interop.Excel;

namespace FinPulse
{
    /// <summary>
    /// Writes the current view to the Dashboard sheet and keeps one native Excel chart in sync with it.
    /// Layout: the chart fills the first columns (A to L) and the data table sits to its right (N to P).
    /// It only ever touches the Dashboard sheet. Financials and Metrics are read, never written.
    /// </summary>
    public class DashboardWriter
    {
        public const string SheetName = "Dashboard";
        private const string ChartObjectName = "FinPulseChart";
        private const int HeaderRow = 5;

        // Layout: the chart fills the first columns (A to L) so it is the first thing seen.
        // The data table sits to its right, in N:P, with column M as a gutter.
        private const string ColA = "N";   // Period
        private const string ColB = "O";   // Value
        private const string ColC = "P";   // Data check
        private const string ChartLastColumn = "L";
        private const int ChartRows = 24;

        private static readonly int Navy = Rgb(15, 42, 74);
        private static readonly int Blue = Rgb(37, 99, 235);
        private static readonly int PaleBlue = Rgb(147, 180, 245);   // every period except the latest
        private static readonly int Amber = Rgb(224, 161, 0);
        private static readonly int Muted = Rgb(100, 116, 139);

        /// <param name="newestFirst">Latest period first (the default) or oldest first.</param>
        /// <param name="activate">Bring the Dashboard sheet to the front, so the chart is the first thing seen.</param>
        public void Write(Excel.Workbook wb, QueryResult result, bool lineChart, bool newestFirst, bool activate)
        {
            Excel.Worksheet ws = FindSheet(wb, SheetName);
            if (ws == null) return;

            Excel.Application app = ws.Application;
            bool oldUpdating = app.ScreenUpdating;
            app.ScreenUpdating = false;
            try
            {
                int n = result.Points.Count;
                bool percent = result.Metric.Kind == ValueKind.Percent;
                string numberFormat = percent ? "0.0\"%\"" : "#,##0";

                // Sheet order: newest first by default. ordered[k] is the period shown in row k.
                List<PeriodPoint> ordered = newestFirst
                    ? Enumerable.Reverse(result.Points).ToList()
                    : result.Points.ToList();
                int latestSlot = newestFirst ? 0 : n - 1;
                PeriodPoint latest = result.Points[n - 1];

                // ---- Title block ----
                // Earlier versions put the table in A:C. Remove it once so it does not sit under the chart.
                if (Convert.ToString(ws.Range["A" + HeaderRow].Value2) == "Period")
                    ws.Range["A1:C80"].Clear();
                ws.Range[ColA + "1:" + ColC + "80"].Clear();      // also resets old highlights and wrapping

                ws.Range[ColA + "1"].Value2 = result.Metric.DisplayName + ", " + Scope(result.Query);
                ws.Range[ColA + "1"].Font.Bold = true;
                ws.Range[ColA + "1"].Font.Size = 13;
                ws.Range[ColA + "1"].Font.Color = Navy;
                ws.Range[ColA + "2"].Value2 = (percent ? "Percent" : "USD thousands") + ", " +
                                        (result.Query.Frequency == Frequency.Quarterly ? "quarterly" : "monthly") +
                                        ", " + n + " periods";
                ws.Range[ColA + "2"].Font.Color = Muted;
                ws.Range[ColA + "3"].Value2 = (newestFirst ? "Newest period first. " : "Oldest period first. ") +
                                        "The latest period (" + latest.Label + ") is highlighted.";
                ws.Range[ColA + "3"].Font.Color = Muted;
                ws.Range[ColA + "3"].Font.Italic = true;

                // ---- Data table, to the right of the chart ----
                ws.Range[ColA + HeaderRow].Value2 = "Period";
                ws.Range[ColB + HeaderRow].Value2 = percent ? "Value (%)" : "Value (USD thousands)";
                ws.Range[ColC + HeaderRow].Value2 = "Data check";
                Excel.Range header = ws.Range[ColA + HeaderRow + ":" + ColC + HeaderRow];
                header.Font.Bold = true;
                header.Font.Color = Rgb(255, 255, 255);
                header.Interior.Color = Navy;

                var values = new object[n, 3];
                for (int k = 0; k < n; k++)
                {
                    PeriodPoint p = ordered[k];
                    values[k, 0] = p.Label;
                    values[k, 1] = p.Value.HasValue ? (object)p.Value.Value : null;      // blank cell = gap in chart
                    values[k, 2] = p.IsIncomplete
                        ? (p.Value.HasValue ? "" : "No figure. ") + p.Detail
                        : "Complete";
                }
                int first = HeaderRow + 1, last = HeaderRow + n;
                ws.Range[ColA + first + ":" + ColA + last].NumberFormat = "@";
                ws.Range[ColA + first + ":" + ColC + last].Value2 = values;
                ws.Range[ColB + first + ":" + ColB + last].NumberFormat = numberFormat;
                ws.Range[ColC + first + ":" + ColC + last].WrapText = true;
                ws.Range[ColA + first + ":" + ColC + last].VerticalAlignment = Excel.XlVAlign.xlVAlignTop;

                // Highlight the latest period in the table as well.
                Excel.Range latestRow = ws.Range[ColA + (first + latestSlot) + ":" + ColC + (first + latestSlot)];
                latestRow.Font.Bold = true;
                latestRow.Interior.Color = Rgb(227, 236, 255);

                ((Excel.Range)ws.Columns["M"]).ColumnWidth = 3;      // gutter between chart and table
                ((Excel.Range)ws.Columns[ColA]).ColumnWidth = 12;
                ((Excel.Range)ws.Columns[ColB]).ColumnWidth = 22;
                ((Excel.Range)ws.Columns[ColC]).ColumnWidth = 34;
                ((Excel.Range)ws.Range[ColA + first + ":" + ColC + last].Rows).AutoFit();

                // ---- Chart in the first columns ----
                Excel.Range anchor = ws.Range["A1"];
                double chartWidth = ws.Range["A1:" + ChartLastColumn + "1"].Width;
                double chartHeight = ws.Range["A1:A" + ChartRows].Height;
                Excel.ChartObject co = FindChartObject(ws);
                if (co == null)
                {
                    co = ((Excel.ChartObjects)ws.ChartObjects(Type.Missing))
                        .Add(anchor.Left, anchor.Top, chartWidth, chartHeight);
                    co.Name = ChartObjectName;
                }
                co.Left = anchor.Left;
                co.Top = anchor.Top;
                co.Width = chartWidth;
                co.Height = chartHeight;

                Excel.Chart chart = co.Chart;
                chart.ChartType = lineChart ? Excel.XlChartType.xlLineMarkers : Excel.XlChartType.xlColumnClustered;
                chart.SetSourceData(ws.Range[ColA + HeaderRow + ":" + ColB + last], Excel.XlRowCol.xlColumns);
                chart.DisplayBlanksAs = Excel.XlDisplayBlanksAs.xlNotPlotted;
                chart.HasLegend = false;
                chart.HasTitle = true;
                chart.ChartTitle.Text = result.Metric.DisplayName + ", " + Scope(result.Query) +
                    (latest.Value.HasValue
                        ? "\nLatest, " + latest.Label + ": " + FinancePane.FormatValue(latest.Value.Value, result.Metric)
                        : "\nLatest, " + latest.Label + ": no figure yet");
                chart.ChartTitle.Font.Size = 12;
                chart.ChartTitle.Font.Bold = true;
                chart.ChartTitle.Font.Color = Navy;

                Excel.Axis valueAxis = (Excel.Axis)chart.Axes(Excel.XlAxisType.xlValue);
                valueAxis.HasTitle = true;
                valueAxis.AxisTitle.Text = percent ? "Percent" : "USD thousands";
                valueAxis.TickLabels.NumberFormat = numberFormat;
                valueAxis.MajorGridlines.Format.Line.ForeColor.RGB = Rgb(226, 232, 240);

                Excel.Series series = (Excel.Series)chart.SeriesCollection(1);
                series.HasDataLabels = true;
                Excel.DataLabels labels = (Excel.DataLabels)series.DataLabels(Type.Missing);
                labels.NumberFormat = numberFormat;
                labels.Font.Size = 9;
                labels.Font.Color = Muted;

                // Everything except the latest period is toned down, so the latest one stands out.
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

                for (int k = 0; k < n; k++)
                {
                    PeriodPoint p = ordered[k];
                    bool isLatest = k == latestSlot;
                    if (!isLatest && !(p.IsIncomplete && p.Value.HasValue)) continue;
                    if (!p.Value.HasValue) continue;

                    Excel.Point pt = (Excel.Point)series.Points(k + 1);
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

                if (result.Points.Any(p => p.IsIncomplete))
                {
                    int noteRow = last + 2;
                    ws.Range[ColA + noteRow].Value2 = "Amber = includes an estimate for a blank input. Loans repeat the previous month; revenue and expenses use the average of the months either side.";
                    ws.Range[ColA + noteRow].Font.Italic = true;
                    ws.Range[ColA + noteRow].Font.Color = Muted;
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
        }

        private static string Scope(Query q)
        {
            return (q.Segment == "All" ? "all segments" : q.Segment) + ", " +
                   (q.Country == "All" ? "all countries" : q.Country);
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