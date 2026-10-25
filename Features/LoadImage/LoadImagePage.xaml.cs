using System;
using System.Windows;
using System.Windows.Controls;
using OfficeWorkAssistant.Excel;

namespace OfficeWorkAssistant.Features.LoadImage
{
    public partial class LoadImagePage : UserControl
    {
        readonly Action<LoadImageSettings> _use;

        public LoadImagePage(LoadImageSettings settings, Action<LoadImageSettings> use)
        {
            InitializeComponent();
            _use = use;
            PathBox.Text = settings != null ? settings.Path : "";
        }

        void Apply_Click(object sender, RoutedEventArgs e)
        {
            var settings = new LoadImageSettings { Path = PathBox.Text.Trim() };
            try
            {
                LoadImageWork.Run(settings);
            }
            catch (Exception ex)
            {
                Notices.Show(this, ex.Message, "Load image", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            Status.Text = "One image, ready for OCR.";
            _use(settings);
        }
    }
}
