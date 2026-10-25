using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using OfficeWorkAssistant.Features.Pipeline;

namespace OfficeWorkAssistant.Views
{
    // The window holds the two modes: Builder (the pipeline canvas and its step settings) and
    // Run (a saved pipeline with only the file and folder choices the person running it needs).
    public partial class MainWindow : Window
    {
        readonly PipelinePage _pipeline = new PipelinePage();
        readonly RunnerPage _runner = new RunnerPage();

        public MainWindow()
        {
            InitializeComponent();
            Closing += MainWindow_Closing;
            ShowMode();
        }

        void Mode_Checked(object sender, RoutedEventArgs e)
        {
            ShowMode();
        }

        void ShowMode()
        {
            // The radios are set while the XAML is still loading.
            if (Root == null || ModeHint == null)
                return;
            var builder = BuilderMode.IsChecked != false;
            ModeHint.Text = builder
                ? "Build and edit pipelines: add steps, connect them, and fill in their settings."
                : "Run a saved pipeline: choose the files and folders it needs, then press Run.";
            Page page = builder ? (Page)_pipeline : (Page)_runner;
            if (!ReferenceEquals(Root.Content, page))
                Root.Navigate(page);
        }

        void MainWindow_Closing(object sender, CancelEventArgs e)
        {
            if (!_pipeline.CanClose())
                e.Cancel = true;
        }
    }
}
