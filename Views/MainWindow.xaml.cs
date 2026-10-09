using System.ComponentModel;
using System.Windows;
using OfficeWorkAssistant.Features.Pipeline;

namespace OfficeWorkAssistant.Views
{
    // The pipeline canvas is the app's main page; step editors are shown in its right-hand Step panel.
    public partial class MainWindow : Window
    {
        readonly PipelinePage _pipeline = new PipelinePage();

        public MainWindow()
        {
            InitializeComponent();
            Root.Navigate(_pipeline);
            Closing += MainWindow_Closing;
        }

        void MainWindow_Closing(object sender, CancelEventArgs e)
        {
            if (!_pipeline.CanClose())
                e.Cancel = true;
        }
    }
}
