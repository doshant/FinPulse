using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using FinPulse.UI.Engine;

namespace FinPulse.UI
{
    /// <summary>One line in the results table.</summary>
    public class ResultRow
    {
        public string Period { get; set; }
        public string ValueText { get; set; }
        public string ChangeText { get; set; }
        public Brush ChangeBrush { get; set; }
        public bool IsLatest { get; set; }
        public string Detail { get; set; }                 // tooltip: what is missing
        public Visibility FlagVisibility { get; set; }     // the amber dot for incomplete periods

        public double? RawValue { get; set; }
        public double? RawChange { get; set; }
    }

    public partial class FinancePane : UserControl
    {
        private static readonly Brush MutedBrush = Frozen("#64748B");
        private static readonly Brush PositiveBrush = Frozen("#0F7B53");
        private static readonly Brush NegativeBrush = Frozen("#C23B3B");
        private static readonly Brush WashBrush = Frozen("#EDF1F7");
        private static readonly Brush PositiveTint = Frozen("#E3F5EC");
        private static readonly Brush NegativeTint = Frozen("#FBE7E7");

        private Dataset _data;
        private bool _ready;          // false while the XAML is loading or the lists are being filled
        private QueryResult _result;
        private List<ResultRow> _rows;

        /// <summary>Raised whenever the numbers change. The add-in draws the Excel chart from this.</summary>
        public event EventHandler<QueryResult> ResultChanged;

        public FinancePane()
        {
            InitializeComponent();
            _ready = true;
            ClearResults("Open a workbook with Financials and Metrics sheets to begin.");
        }

        public bool UseLineChart { get { return LineOption.IsChecked == true; } }

        /// <summary>True (the default) lists the latest period first, in the table and on the chart.</summary>
        public bool NewestFirst { get { return NewestOption.IsChecked == true; } }
        public QueryResult CurrentResult { get { return _result; } }

        // ------------------------------------------------------------------
        // Called by the add-in
        // ------------------------------------------------------------------

        /// <summary>Fills the pickers from the workbook (so new segments or countries just appear) and refreshes.</summary>
        public void LoadDataset(Dataset data)
        {
            _data = data;
            bool wasReady = _ready;
            _ready = false;

            string metric = SelectedTag(MetricCombo), segment = SelectedTag(SegmentCombo), country = SelectedTag(CountryCombo);

            Fill(MetricCombo, data.Metrics.Select(m => new KeyValuePair<string, string>(m.Key, m.DisplayName)), metric);
            Fill(SegmentCombo,
                 new[] { new KeyValuePair<string, string>("All", "All segments") }
                     .Concat(data.Segments.Select(s => new KeyValuePair<string, string>(s, s))), segment);
            Fill(CountryCombo,
                 new[] { new KeyValuePair<string, string>("All", "All countries") }
                     .Concat(data.Countries.Select(c => new KeyValuePair<string, string>(c, c))), country);

            _ready = wasReady;
            Refresh();
        }

        /// <summary>Shows a plain message instead of results, for example when the sheets are missing.</summary>
        public void ShowMessage(string message)
        {
            _data = null;
            ClearResults(message);
        }

        /// <summary>Shows a one-line status without clearing the results.</summary>
        public void ShowStatus(string message, bool isError)
        {
            StatusText.Text = message;
            StatusText.Foreground = isError ? NegativeBrush : MutedBrush;
        }

        // ------------------------------------------------------------------
        // Events
        // ------------------------------------------------------------------

        private void Filter_Changed(object sender, RoutedEventArgs e) { Refresh(); }

        private void ChartType_Changed(object sender, RoutedEventArgs e) { Refresh(); }

        private void ResetButton_Click(object sender, RoutedEventArgs e)
        {
            _ready = false;
            MetricCombo.SelectedIndex = 0;
            SegmentCombo.SelectedIndex = 0;
            CountryCombo.SelectedIndex = 0;
            QuarterlyOption.IsChecked = true;
            foreach (RadioButton rb in PeriodsHost.Children.OfType<RadioButton>())
                rb.IsChecked = Convert.ToString(rb.Tag) == "6";
            ColumnOption.IsChecked = true;
            NewestOption.IsChecked = true;
            _ready = true;
            Refresh();
        }

        private void CopyButton_Click(object sender, RoutedEventArgs e)
        {
            if (_rows == null || _result == null) return;

            bool percent = _result.Metric.Kind == ValueKind.Percent;
            var sb = new StringBuilder();
            sb.AppendLine("Period\t" + (percent ? "Value (%)" : "Value (USD thousands)") +
                          "\t" + (percent ? "Change (pp)" : "Change (%)") + "\tComplete");

            foreach (ResultRow row in _rows)   // same order as the table on screen
            {
                sb.AppendLine(row.Period + "\t" +
                              (row.RawValue.HasValue ? row.RawValue.Value.ToString("0.###", CultureInfo.InvariantCulture) : "") + "\t" +
                              (row.RawChange.HasValue ? Math.Round(row.RawChange.Value, 2).ToString(CultureInfo.InvariantCulture) : "") + "\t" +
                              (row.FlagVisibility == Visibility.Visible ? "No" : "Yes"));
            }

            try
            {
                Clipboard.SetText(sb.ToString());
                StatusText.Text = "Copied " + _rows.Count + " rows. Paste them into Excel.";
                StatusText.Foreground = MutedBrush;
            }
            catch (ExternalException)
            {
                StatusText.Text = "Another app is using the clipboard. Try again in a moment.";
                StatusText.Foreground = NegativeBrush;
            }
        }

        // ------------------------------------------------------------------
        // Core refresh
        // ------------------------------------------------------------------

        private void Refresh()
        {
            if (!_ready || _data == null || MetricCombo.SelectedItem == null) return;

            try
            {
                Query q = BuildQuery();
                QueryResult result = FinanceEngine.Run(_data, q);

                ShowResult(result, FinanceEngine.SamePeriodLastYear(_data, q));

                EventHandler<QueryResult> handler = ResultChanged;
                if (handler != null) handler(this, result);
            }
            catch (Exception ex)
            {
                StatusText.Text = "Couldn't calculate this view. " + ex.Message;
                StatusText.Foreground = NegativeBrush;
            }
        }

        private Query BuildQuery()
        {
            RadioButton periods = PeriodsHost.Children.OfType<RadioButton>().FirstOrDefault(r => r.IsChecked == true);
            int n;
            if (periods == null || !int.TryParse(Convert.ToString(periods.Tag), out n)) n = 6;

            Frequency freq = QuarterlyOption.IsChecked == true ? Frequency.Quarterly : Frequency.Monthly;
            if (n == 0)   // "All": every period the data holds
            {
                n = freq == Frequency.Quarterly
                    ? FinanceEngine.BuildPeriods(_data.Months, Frequency.Quarterly, int.MaxValue).Count
                    : _data.Months.Count;
            }

            return new Query
            {
                MetricKey = SelectedTag(MetricCombo),
                Segment = SelectedTag(SegmentCombo),
                Country = SelectedTag(CountryCombo),
                Frequency = freq,
                Periods = n
            };
        }

        // ------------------------------------------------------------------
        // Rendering
        // ------------------------------------------------------------------

        private void ShowResult(QueryResult result, double? yearAgo)
        {
            _result = result;
            MetricInfo metric = result.Metric;
            List<PeriodPoint> pts = result.Points;
            bool percent = metric.Kind == ValueKind.Percent;
            int n = pts.Count;

            PeriodPoint latest = pts[n - 1];
            PeriodPoint previous = n > 1 ? pts[n - 2] : null;

            // --- Snapshot
            string scope = ScopeText(result.Query);
            KpiLabel.Text = metric.DisplayName + ", " + scope;
            KpiValueRun.Text = latest.Value.HasValue ? FormatValue(latest.Value.Value, metric) : "—";
            KpiUnitRun.Text = percent || !latest.Value.HasValue ? string.Empty : "  USD thousands";

            Delta d;
            if (previous != null && latest.Value.HasValue && previous.Value.HasValue)
            {
                d = ComputeDelta(latest.Value.Value, previous.Value.Value, metric);
                ChangeCaption.Text = "vs " + previous.Label;
                ChangeCaption.Visibility = Visibility.Visible;
            }
            else
            {
                d = new Delta { Text = "No comparison", Brush = MutedBrush, Chip = WashBrush };
                ChangeCaption.Visibility = Visibility.Collapsed;
            }
            ChangeChip.Background = d.Chip;
            ChangeChipText.Text = d.Text;
            ChangeChipText.Foreground = d.Brush;

            if (yearAgo.HasValue && latest.Value.HasValue)
            {
                Delta y = ComputeDelta(latest.Value.Value, yearAgo.Value, metric);
                YoyText.Text = "Same period last year: " + FormatValue(yearAgo.Value, metric) + "  (" + y.Text + ")";
                YoyText.Visibility = Visibility.Visible;
            }
            else
            {
                YoyText.Visibility = Visibility.Collapsed;
            }

            DrawSparkline(NewestFirst ? Enumerable.Reverse(pts).ToList() : pts);
            KpiHint.Text = latest.Value.HasValue
                ? "Latest period: " + latest.Label + (latest.IsIncomplete ? ". Includes estimated figures." : ".")
                : latest.Label + " has no figure yet. " + (latest.Detail ?? string.Empty);

            SourceBadge.Visibility = Visibility.Collapsed;

            // --- Table, in the order chosen in the pane (newest first by default)
            var rows = new List<ResultRow>();
            IEnumerable<int> order = NewestFirst ? Enumerable.Range(0, n).Reverse() : Enumerable.Range(0, n);
            foreach (int i in order)
            {
                PeriodPoint p = pts[i];
                var row = new ResultRow
                {
                    Period = p.Label,
                    ValueText = p.Value.HasValue ? FormatValue(p.Value.Value, metric) : "—",
                    RawValue = p.Value,
                    IsLatest = i == n - 1,
                    Detail = p.Detail,
                    FlagVisibility = p.IsIncomplete ? Visibility.Visible : Visibility.Collapsed,
                    ChangeText = "—",
                    ChangeBrush = MutedBrush
                };
                if (i > 0 && p.Value.HasValue && pts[i - 1].Value.HasValue)
                {
                    Delta dd = ComputeDelta(p.Value.Value, pts[i - 1].Value.Value, metric);
                    row.ChangeText = dd.Text;
                    row.ChangeBrush = dd.Brush;
                    row.RawChange = dd.Raw;
                }
                rows.Add(row);
            }
            _rows = rows;

            ValueHeader.Text = percent ? "Value (%)" : "Value (USD k)";
            ResultsList.ItemsSource = rows;
            ResultsHeader.Visibility = Visibility.Visible;
            EmptyResults.Visibility = Visibility.Collapsed;
            CopyButton.IsEnabled = true;

            // --- Data quality
            ShowQuality(result);

            // --- Takeaways and status
            InsightText.Text = BuildInsight(result);
            InsightNoteBox.Visibility = Visibility.Collapsed;

            StatusText.Text = n + (result.Query.Frequency == Frequency.Quarterly ? " quarters" : " months") +
                              " ending " + latest.Label + ". Chart updated on the Dashboard sheet.";
            StatusText.Foreground = MutedBrush;
        }

        private void ShowQuality(QueryResult result)
        {
            var distinct = result.Issues
                .GroupBy(i => i.RowId + "|" + i.Month)
                .Select(g => g.First())
                .ToList();

            if (distinct.Count == 0)
            {
                QualityBox.Visibility = Visibility.Collapsed;
                return;
            }

            var sb = new StringBuilder();
            sb.Append(distinct.Count + (distinct.Count == 1 ? " blank cell was" : " blank cells were") +
                      " filled in so totals stay comparable. Loans repeat the previous month; revenue and expenses use the average of the months either side. Periods marked ● include these estimates.");
            foreach (DataIssue i in distinct.Take(3))
            {
                sb.AppendLine();
                sb.Append("• " + i.Segment + ", " + i.Country + ", " + i.Metric + ", " +
                          FinanceEngine.PeriodLabel(i.Month, Frequency.Monthly) + ": " +
                          (i.FilledWith.HasValue
                              ? Math.Round(i.FilledWith.Value, MidpointRounding.AwayFromZero).ToString("N0", CultureInfo.InvariantCulture) + " (" + i.Method + ")"
                              : "no figure (" + i.Method + ")") +
                          (string.IsNullOrWhiteSpace(i.Note) ? string.Empty : ". Sheet note: " + i.Note.Trim()));
            }
            if (distinct.Count > 3) { sb.AppendLine(); sb.Append("• and " + (distinct.Count - 3) + " more"); }

            QualityText.Text = sb.ToString();
            QualityBox.Visibility = Visibility.Visible;
        }

        private void ClearResults(string message)
        {
            _result = null;
            _rows = null;

            KpiLabel.Text = "Latest period value";
            KpiValueRun.Text = "—";
            KpiUnitRun.Text = string.Empty;
            ChangeChip.Background = WashBrush;
            ChangeChipText.Text = "No data yet";
            ChangeChipText.Foreground = MutedBrush;
            ChangeCaption.Visibility = Visibility.Collapsed;
            YoyText.Visibility = Visibility.Collapsed;
            Sparkline.Visibility = Visibility.Collapsed;
            SourceBadge.Visibility = Visibility.Collapsed;
            QualityBox.Visibility = Visibility.Collapsed;
            KpiHint.Text = message;

            ResultsList.ItemsSource = null;
            ResultsHeader.Visibility = Visibility.Collapsed;
            EmptyResults.Visibility = Visibility.Visible;
            CopyButton.IsEnabled = false;

            InsightText.Text = "Pick a metric, segment and country to see how the numbers moved.";
            InsightNoteBox.Visibility = Visibility.Collapsed;
            StatusText.Text = message;
            StatusText.Foreground = MutedBrush;
        }

        private void DrawSparkline(IList<PeriodPoint> pts)
        {
            var have = pts.Select((p, i) => new { p.Value, Index = i }).Where(x => x.Value.HasValue).ToList();
            if (have.Count < 2) { Sparkline.Visibility = Visibility.Collapsed; return; }

            double min = have.Min(x => x.Value.Value), max = have.Max(x => x.Value.Value);
            double range = max - min;
            int last = pts.Count - 1;

            var points = new PointCollection();
            foreach (var x in have)
            {
                double px = last == 0 ? 0 : x.Index * 100.0 / last;
                double py = range < 1e-9 ? 20.0 : 40.0 - ((x.Value.Value - min) / range) * 40.0;
                points.Add(new Point(px, py));
            }
            Sparkline.Points = points;
            Sparkline.Visibility = Visibility.Visible;
        }

        private static string BuildInsight(QueryResult result)
        {
            MetricInfo metric = result.Metric;
            var have = result.Points.Where(p => p.Value.HasValue).ToList();
            if (have.Count < 2) return "There are not enough figures in this view to describe a trend.";

            string unitWord = result.Query.Frequency == Frequency.Quarterly ? "quarters" : "months";
            PeriodPoint first = have[0], last = have[have.Count - 1];
            Delta overall = ComputeDelta(last.Value.Value, first.Value.Value, metric);

            var sb = new StringBuilder();
            if (Math.Abs(overall.Raw) < 0.05)
                sb.Append(metric.DisplayName + " was flat between " + first.Label + " and " + last.Label + ".");
            else
                sb.Append(metric.DisplayName + (overall.Raw > 0 ? " rose " : " fell ") + overall.Magnitude +
                          " from " + first.Label + " to " + last.Label + ", reaching " +
                          FormatValue(last.Value.Value, metric) + (metric.Kind == ValueKind.Percent ? "" : " thousand") + ".");

            int bi = 1; double biggest = -1;
            for (int i = 1; i < have.Count; i++)
            {
                double c = Math.Abs(ComputeDelta(have[i].Value.Value, have[i - 1].Value.Value, metric).Raw);
                if (c > biggest) { biggest = c; bi = i; }
            }
            Delta bd = ComputeDelta(have[bi].Value.Value, have[bi - 1].Value.Value, metric);
            sb.Append(" The biggest move was in " + have[bi].Label + " (" + (bd.Raw > 0 ? "+" : "−") + bd.Magnitude + ").");

            if (Math.Abs(overall.Raw) >= 0.05 && (overall.Raw > 0) != metric.HigherIsBetter)
                sb.Append(" For this measure, that direction is unfavorable.");

            return sb.ToString();
        }

        // ------------------------------------------------------------------
        // Helpers
        // ------------------------------------------------------------------

        private static string ScopeText(Query q)
        {
            string seg = q.Segment == "All" ? "all segments" : q.Segment;
            string cty = q.Country == "All" ? "all countries" : q.Country;
            return seg + ", " + cty;
        }

        private static void Fill(ComboBox combo, IEnumerable<KeyValuePair<string, string>> items, string keep)
        {
            combo.Items.Clear();
            foreach (var kv in items) combo.Items.Add(new ComboBoxItem { Content = kv.Value, Tag = kv.Key });

            ComboBoxItem match = combo.Items.OfType<ComboBoxItem>().FirstOrDefault(i => Convert.ToString(i.Tag) == keep);
            combo.SelectedItem = match ?? combo.Items[0];
        }

        private static string SelectedTag(ComboBox combo)
        {
            ComboBoxItem item = combo.SelectedItem as ComboBoxItem;
            return item != null ? Convert.ToString(item.Tag) : string.Empty;
        }

        private struct Delta
        {
            public string Text;       
            public string Magnitude;   // "2.4%"
            public double Raw;         // signed change in percent, or percentage points for ratios
            public Brush Brush;
            public Brush Chip;
        }

        private static Delta ComputeDelta(double current, double previous, MetricInfo metric)
        {
            bool percent = metric.Kind == ValueKind.Percent;
            double change = percent ? current - previous
                                    : (Math.Abs(previous) < 1e-9 ? 0.0 : (current - previous) / previous * 100.0);
            string magnitude = Math.Abs(change).ToString("0.0", CultureInfo.InvariantCulture) + (percent ? " pp" : "%");

            if (Math.Abs(change) < 0.05)
                return new Delta { Text = "— " + magnitude, Magnitude = magnitude, Raw = 0, Brush = MutedBrush, Chip = WashBrush };

            bool up = change > 0;
            bool good = up == metric.HigherIsBetter;
            return new Delta
            {
                Text = (up ? "▲ " : "▼ ") + magnitude,
                Magnitude = magnitude,
                Raw = change,
                Brush = good ? PositiveBrush : NegativeBrush,
                Chip = good ? PositiveTint : NegativeTint
            };
        }

        public static string FormatValue(double value, MetricInfo metric)
        {
            return metric.Kind == ValueKind.Percent
                ? value.ToString("0.0", CultureInfo.InvariantCulture) + "%"
                : Math.Round(value, MidpointRounding.AwayFromZero).ToString("N0", CultureInfo.InvariantCulture);
        }

        private static Brush Frozen(string hex)
        {
            var brush = new SolidColorBrush((Color)ColorConverter.ConvertFromString(hex));
            brush.Freeze();
            return brush;
        }
    }
}