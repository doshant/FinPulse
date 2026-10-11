
using System;
using Microsoft.Office.Tools.Ribbon;

namespace FinPulse
{
    public partial class RibbonToggle : RibbonBase
    {
        private RibbonTab _tab;
        private RibbonGroup _group;
        private RibbonButton _btnToggle;

        public RibbonToggle()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            // Show the ribbon in Excel workbooks.
            this.RibbonType = "Microsoft.Excel.Workbook";

            // Create the FinPulse tab.
            _tab = this.Factory.CreateRibbonTab();
            _tab.Label = "FinPulse";

            // Create the Dashboard group.
            _group = this.Factory.CreateRibbonGroup();
            _group.Label = "Dashboard";

            // Create the Show/Hide button.
            _btnToggle = this.Factory.CreateRibbonButton();
            _btnToggle.Label = "Launch FinPulse";
            _btnToggle.Click += BtnToggle_Click;

            // Add the button to the group and the group to the tab.
            _group.Items.Add(_btnToggle);
            _tab.Groups.Add(_group);

            // Register the tab.
            this.Tabs.Add(_tab);
        }

        private void BtnToggle_Click(
            object sender,
            RibbonControlEventArgs e)
        {
            try
            {
                Globals.ThisAddIn.ToggleTaskPane();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    "FinPulse ribbon error: " + ex);
            }
        }
    }
}
