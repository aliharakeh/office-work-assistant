using System.Windows;

namespace WorkAssistant.Views
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
