using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using WorkAssistant.Expressions;
using WorkAssistant.Views;

namespace WorkAssistant.Features.FormulaGuide
{
    public partial class FormulaGuidePage : Page
    {
        static readonly Brush OkBrush = new SolidColorBrush(Color.FromRgb(0x00, 0x6B, 0x3C));
        static readonly Brush BadBrush = new SolidColorBrush(Color.FromRgb(0xB0, 0x00, 0x20));

        readonly List<SampleValue> _samples = FormulaGuideWork.DefaultSamples();

        public FormulaGuidePage()
        {
            InitializeComponent();
            SampleGrid.ItemsSource = _samples;
            LessonList.ItemsSource = FormulaGuideWork.BuildLessons(_samples);
            RunTest();
        }

        void Home_Click(object sender, RoutedEventArgs e)
        {
            if (NavigationService != null && NavigationService.CanGoBack)
                NavigationService.GoBack();
            else if (NavigationService != null)
                NavigationService.Navigate(new HomePage());
        }

        void Variables_Click(object sender, RoutedEventArgs e)
        {
            ExpressionHelp.Show(Window.GetWindow(this),
                "Built-in names that work in any formula. The sample values on this page stand in for the columns and fields the other tools offer.");
        }

        void Test_Click(object sender, RoutedEventArgs e)
        {
            RunTest();
        }

        void FormulaBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
                RunTest();
        }

        void TryStep_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            if (button == null)
                return;
            FormulaBox.Text = Convert.ToString(button.Tag);
            FormulaBox.Focus();
            FormulaBox.SelectAll();
            RunTest();
        }

        void SampleGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            // The cell writes back after this event, so rebuild once the edit lands.
            Dispatcher.BeginInvoke(DispatcherPriority.Background, new Action(delegate
            {
                LessonList.ItemsSource = FormulaGuideWork.BuildLessons(_samples);
            }));
        }

        void RunTest()
        {
            bool error;
            string result = FormulaGuideWork.Run(FormulaBox.Text, _samples, out error);
            TestResult.Text = result;
            TestResult.Foreground = error ? BadBrush : OkBrush;
        }
    }
}
