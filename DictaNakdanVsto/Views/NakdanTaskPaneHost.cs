using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace DictaNakdanVsto.Views
{
    public partial class NakdanTaskPaneHost : UserControl
    {
        public NakdanTaskPaneHost()
        {
            InitializeComponent();

            // כאן אנחנו יוצרים את ה"מארח" שלוקח את עיצוב ה-WPF שכתבנו, ומכניס אותו לקופסה
            ElementHost host = new ElementHost
            {
                Dock = DockStyle.Fill,
                Child = new NakdanView() // זהו מסך ה-WPF המקורי שלנו
            };

            this.Controls.Add(host);
        }
    }
}