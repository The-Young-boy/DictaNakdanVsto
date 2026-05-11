using Microsoft.Office.Tools.Ribbon;

namespace DictaNakdanVsto
{
    public partial class NakdanRibbon
    {
        private void NakdanRibbon_Load(object sender, RibbonUIEventArgs e) { }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            // פתיחה והסתרה של החלונית בלחיצה
            Globals.ThisAddIn.NakdanTaskPane.Visible = !Globals.ThisAddIn.NakdanTaskPane.Visible;
        }
    }
}