using System.Windows;
using System.Windows.Controls;
using WorkAssistant.Features.ExcelProcessing;
using WorkAssistant.Features.CopyFiles;
using WorkAssistant.Features.MergeDuplicates;
using WorkAssistant.Features.Templates;
using WorkAssistant.Features.FilterSort;

namespace WorkAssistant.Views
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
    }
}
