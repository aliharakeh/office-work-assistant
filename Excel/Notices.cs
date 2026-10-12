using System.Windows;

namespace OfficeWorkAssistant.Excel
{
    // Warnings from a step editor. The pipeline applies an editor's settings on its own while the user types,
    // so during that a warning must not open a dialog: it is kept in Last and shown in the step's section.
    public static class Notices
    {
        public static bool Quiet;
        public static string Last;

        public static MessageBoxResult Show(DependencyObject owner, string message, string title, MessageBoxButton buttons, MessageBoxImage icon)
        {
            if (Quiet)
            {
                Last = message;
                return MessageBoxResult.None;
            }
            return MessageBox.Show(Window.GetWindow(owner), message, title, buttons, icon);
        }
    }
}
