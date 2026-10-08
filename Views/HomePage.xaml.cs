using System.Windows;
using System.Windows.Controls;
using OfficeWorkAssistant.Features.ExcelProcessing;
using OfficeWorkAssistant.Features.CopyFiles;
using OfficeWorkAssistant.Features.MergeDuplicates;
using OfficeWorkAssistant.Features.Templates;
using OfficeWorkAssistant.Features.FilterSort;
using OfficeWorkAssistant.Features.FormulaGuide;
using OfficeWorkAssistant.Features.FillColumns;
using OfficeWorkAssistant.Features.Pipeline;

namespace OfficeWorkAssistant.Views
{
    public partial class HomePage : Page
    {
        public HomePage()
        {
            InitializeComponent();
        }

        void ExcelCard_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new ExcelProcessingPage());
        }

        void FilesCard_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new CopyFilesPage());
        }

        void FoldersCard_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new MergeDuplicatesPage());
        }

        void TemplatesCard_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new TemplatesPage());
        }

        void FilterSortCard_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new FilterSortPage());
        }

        void FormulaGuideCard_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new FormulaGuidePage());
        }

        void FillColumnsCard_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new FillColumnsPage());
        }

        void PipelineCard_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new PipelinePage());
        }
    }
}
