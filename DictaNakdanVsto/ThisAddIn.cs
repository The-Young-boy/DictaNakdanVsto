using System;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using DictaNakdanVsto.Views;
using DictaNakdanVsto.ViewModels;

namespace DictaNakdanVsto
{
    public partial class ThisAddIn
    {
        public Microsoft.Office.Tools.CustomTaskPane NakdanTaskPane;

        internal class PictureDispConverter : System.Windows.Forms.AxHost
        {
            private PictureDispConverter() : base("") { }
            public static stdole.IPictureDisp ToIPictureDisp(System.Drawing.Image image)
            {
                return (stdole.IPictureDisp)GetIPictureDispFromPicture(image);
            }
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            var myHost = new NakdanTaskPaneHost();
            NakdanTaskPane = this.CustomTaskPanes.Add(myHost, "נקדן דיקטה");
            NakdanTaskPane.Width = 360;
            NakdanTaskPane.Visible = false;

            this.Application.WindowBeforeRightClick += Application_WindowBeforeRightClick;
        }

        private void Application_WindowBeforeRightClick(Word.Selection Sel, ref bool Cancel)
        {
            Office.CommandBar contextMenu = this.Application.CommandBars["Text"];

            foreach (Office.CommandBarControl ctrl in contextMenu.Controls)
            {
                if (ctrl.Caption == "נקד באמצעות נקדן דיקטה") ctrl.Delete();
            }

            Office.CommandBarButton customButton = (Office.CommandBarButton)contextMenu.Controls.Add(
                Office.MsoControlType.msoControlButton, missing, missing, missing, true);

            customButton.Caption = "נקד באמצעות נקדן דיקטה";
            customButton.Tag = "DictaNakdanBtn";

            try
            {
                customButton.Picture = PictureDispConverter.ToIPictureDisp(Properties.Resources.icon_16);
            }
            catch { }

            customButton.Click += (Office.CommandBarButton Ctrl, ref bool CancelDefault) =>
            {
                NakdanTaskPane.Visible = true;

                // התיקון: שולפים את ה-ElementHost מתוך פקדי ה-Host
                var host = (NakdanTaskPaneHost)NakdanTaskPane.Control;
                var elementHost = (System.Windows.Forms.Integration.ElementHost)host.Controls[0];
                var view = (NakdanView)elementHost.Child;
                var vm = (NakdanViewModel)view.DataContext;

                if (vm.PunctuateCommand.CanExecute(null))
                {
                    vm.PunctuateCommand.Execute(null);
                }
            };
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e) { }

        #region VSTO generated code
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}