using System.Windows;

namespace OfficeWorkAssistant.Views
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            Root.Navigate(new HomePage());
        }
    }
}
