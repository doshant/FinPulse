using System;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using FinPulse.UI;
using FinPulse.UI.Engine;
using Microsoft.Office.Tools;
using Excel = Microsoft.Office.Interop.Excel;

namespace FinPulse
{
    public partial class ThisAddIn
    {
        private FinancePane _pane;
        private CustomTaskPane _taskPane;
        private bool _started;
        private bool _showDashboardNext;     // true right after Launch, so the chart is the first thing the user sees
        private readonly DashboardWriter _dashboard = new DashboardWriter();
        private System.Windows.Forms.Timer _reloadTimer;     // waits for edits to settle before re-reading the sheets

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            // Nothing starts automatically. The user clicks the ribbon "Launch FinPulse" button.
            _started = false;
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
        }

        private void Application_WorkbookActivate(Excel.Workbook wb)
        {
            LoadFrom(wb);
        }

        /// <summary>Reads the Financials and Metrics sheets (read only) and hands the data to the pane.</summary>
        private void LoadFrom(Excel.Workbook wb)
        {
            if (_pane == null || wb == null)
                return;

            try
            {
                Excel.Worksheet fin = DashboardWriter.FindSheet(wb, "Financials");
                Excel.Worksheet met = DashboardWriter.FindSheet(wb, "Metrics");
                if (fin == null || met == null)
                {
                    _pane.ShowMessage("This workbook has no Financials and Metrics sheets, so there is nothing to chart yet.");
                    return;
                }

                object[,] finValues = fin.UsedRange.Value2 as object[,];
                object[,] metValues = met.UsedRange.Value2 as object[,];
                if (finValues == null || metValues == null)
                {
                    _pane.ShowMessage("The Financials or Metrics sheet is empty.");
                    return;
                }

                _pane.LoadDataset(DatasetBuilder.Build(finValues, metValues));
            }
            catch (Exception ex)
            {
                _pane.ShowMessage("Couldn't read the workbook. " + ex.Message);
            }
        }

        private void Pane_ResultChanged(object sender, QueryResult result)
        {
            Excel.Workbook wb = Application.ActiveWorkbook;
            if (wb == null || result == null) return;

            try
            {
                bool activate = _showDashboardNext;
                _showDashboardNext = false;
                if (!_dashboard.Write(wb, result, _pane.UseLineChart, _pane.NewestFirst, activate))
                    _pane.ShowStatus("This workbook has no sheet named Dashboard, so there is nowhere to draw the chart.", true);
            }
            catch (Exception ex)
            {
                _pane.ShowStatus("Couldn't update the Dashboard chart. " + ex.Message, true);
            }
        }

        /// <summary>
        /// Fires on every cell edit. Only edits on the Financials or Metrics sheets matter. A short delay groups a
        /// paste of many cells into one reload.
        /// </summary>
        private void Application_SheetChange(object sh, Excel.Range target)
        {
            Excel.Worksheet ws = sh as Excel.Worksheet;
            if (ws == null) return;
            if (!string.Equals(ws.Name, "Financials", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(ws.Name, "Metrics", StringComparison.OrdinalIgnoreCase)) return;

            if (_reloadTimer == null)
            {
                _reloadTimer = new System.Windows.Forms.Timer { Interval = 500 };
                _reloadTimer.Tick += delegate
                {
                    _reloadTimer.Stop();
                    Excel.Workbook wb = Application.ActiveWorkbook;
                    if (wb != null) LoadFrom(wb);
                };
            }
            _reloadTimer.Stop();
            _reloadTimer.Start();
        }

        private void Pane_RefreshRequested(object sender, EventArgs e)
        {
            Excel.Workbook wb = Application.ActiveWorkbook;
            if (wb != null) LoadFrom(wb);
        }

        #region VSTO generated code
        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }
        #endregion

        /// <summary>
        /// Called by the ribbon button. The first click launches FinPulse. After that it shows or hides the pane.
        /// </summary>
        public void ToggleTaskPane()
        {
            try
            {
                if (!_started)
                {
                    Launch();
                    return;
                }

                if (_taskPane == null)
                {
                    ShowTaskPane();
                    return;
                }

                if (_taskPane.Visible)
                {
                    HideTaskPane();
                }
                else
                {
                    ShowTaskPane();

                    Excel.Workbook wb = Application.ActiveWorkbook;
                    if (wb != null) LoadFrom(wb);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("FinPulse toggle failed: " + ex.Message);
            }
        }

        /// <summary>Ensure the FinPulse custom task pane exists and is visible.</summary>
        public void ShowTaskPane()
        {
            try
            {
                if (_taskPane == null)
                {
                    foreach (CustomTaskPane ctp in CustomTaskPanes)
                    {
                        if (string.Equals(ctp.Title, "FinPulse", StringComparison.OrdinalIgnoreCase))
                        {
                            _taskPane = ctp;
                            break;
                        }
                    }
                }

                if (_taskPane == null)
                {
                    EnsurePaneCreated();
                }

                try { _taskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight; } catch { }
                _taskPane.Width = 380;
                _taskPane.Visible = true;

                // If showing failed, recreate as a last resort
                if (!_taskPane.Visible)
                {
                    var old = _taskPane;
                    try { CustomTaskPanes.Remove(old); } catch { }
                    _taskPane = null;
                    EnsurePaneCreated();
                    try { _taskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight; } catch { }
                    _taskPane.Width = 380;
                    _taskPane.Visible = true;
                }
            }
            catch (Exception)
            {
                // ignore
            }
        }

        /// <summary>Hide the FinPulse task pane if it exists.</summary>
        public void HideTaskPane()
        {
            try
            {
                if (_taskPane == null)
                {
                    foreach (CustomTaskPane ctp in CustomTaskPanes)
                    {
                        if (string.Equals(ctp.Title, "FinPulse", StringComparison.OrdinalIgnoreCase))
                        {
                            _taskPane = ctp;
                            break;
                        }
                    }
                }

                if (_taskPane != null)
                {
                    _taskPane.Visible = false;
                }
            }
            catch (Exception)
            {
                // ignore
            }
        }

        /// <summary>
        /// Launch the add-in: show the pane, hook workbook events and load the active workbook.
        /// Safe to call more than once; it only initializes once.
        /// </summary>
        public void Launch()
        {
            if (_started) return;
            try
            {
                ShowTaskPane();

                Application.WorkbookActivate += Application_WorkbookActivate;
                Application.WorkbookOpen += Application_WorkbookActivate;
                Application.SheetChange += Application_SheetChange;       // new months typed or pasted in show up by themselves

                _showDashboardNext = true;       // after loading, jump to the Dashboard sheet so the chart is seen first
                if (Application.ActiveWorkbook != null) LoadFrom(Application.ActiveWorkbook);

                _started = true;
            }
            catch (Exception)
            {
                // ignore
            }
        }

        private void EnsurePaneCreated()
        {
            if (_pane == null)
            {
                _pane = new FinancePane();
                _pane.ResultChanged += Pane_ResultChanged;
                _pane.RefreshRequested += Pane_RefreshRequested;
            }

            var host = new ElementHost { Dock = DockStyle.Fill, Child = _pane };
            var hostContainer = new System.Windows.Forms.UserControl();
            hostContainer.Controls.Add(host);
            _taskPane = CustomTaskPanes.Add(hostContainer, "FinPulse");
            _taskPane.Width = 380;
        }
    }
}