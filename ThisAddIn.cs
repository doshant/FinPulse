
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using FinPulse.UI;
using Excel = Microsoft.Office.Interop.Excel;

namespace FinPulse
{
    public partial class ThisAddIn
    {
        private readonly Dictionary<int, Microsoft.Office.Tools.CustomTaskPane>
            _taskPanes =
                new Dictionary<int, Microsoft.Office.Tools.CustomTaskPane>();

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            ((Excel.AppEvents_Event)this.Application).NewWorkbook += OnNewWorkbook;
            this.Application.WorkbookOpen += OnWorkbookOpen;
            this.Application.WindowActivate += OnWindowActivate;

            // Handle a workbook that is already active at startup.
            if (this.Application.ActiveWorkbook != null &&
                this.Application.ActiveWindow != null)
            {
                EnsurePane(this.Application.ActiveWindow);
            }
        }

        private void OnNewWorkbook(Excel.Workbook workbook)
        {
            if (workbook != null && workbook.Windows.Count > 0)
            {
                EnsurePane(workbook.Windows[1]);
            }
        }

        private void OnWorkbookOpen(Excel.Workbook workbook)
        {
            if (workbook != null && workbook.Windows.Count > 0)
            {
                EnsurePane(workbook.Windows[1]);
            }
        }

        private void OnWindowActivate(
            Excel.Workbook workbook,
            Excel.Window window)
        {
            EnsurePane(window);
        }

        private void EnsurePane(Excel.Window window)
        {
            if (window == null)
                return;

            int hwnd = window.Hwnd;

            if (!_taskPanes.TryGetValue(hwnd, out var pane))
            {
                var wpfPane = new FinancePane();

                var elementHost = new ElementHost
                {
                    Dock = DockStyle.Fill,
                    Child = wpfPane
                };

                var hostControl = new UserControl
                {
                    Dock = DockStyle.Fill
                };

                hostControl.Controls.Add(elementHost);

                pane = this.CustomTaskPanes.Add(
                    hostControl,
                    "FinPulse",
                    window
                );

                pane.Width = 360;

                _taskPanes.Add(hwnd, pane);
            }

            pane.Visible = true;
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            ((Excel.AppEvents_Event)this.Application).NewWorkbook -= OnNewWorkbook;
            this.Application.WorkbookOpen -= OnWorkbookOpen;
            this.Application.WindowActivate -= OnWindowActivate;

            foreach (var pane in _taskPanes.Values)
            {
                pane.Visible = false;
            }

            _taskPanes.Clear();
        }

        #region VSTO generated code

        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
