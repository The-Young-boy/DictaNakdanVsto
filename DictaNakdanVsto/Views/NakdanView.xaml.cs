using System.Windows.Controls;
using DictaNakdanVsto.ViewModels;

namespace DictaNakdanVsto.Views
{
    public partial class NakdanView : UserControl
    {
        public NakdanView()
        {
            InitializeComponent();

            // אנחנו אומרים למסך: ה"מוח" שמנהל אותך הוא ה-ViewModel שלנו
            this.DataContext = new NakdanViewModel();
        }
    }
}