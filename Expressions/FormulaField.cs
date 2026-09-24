using System;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace OfficeWorkAssistant.Expressions
{
    public static class FormulaField
    {
        static readonly ConditionalWeakTable<TextBox, State> States = new ConditionalWeakTable<TextBox, State>();
        static readonly Brush ErrorBorder = Freeze(0xC6, 0x28, 0x28);
        static readonly Brush ErrorFill = Freeze(0xFF, 0xEB, 0xEE);

        public static void Watch(TextBox box)
        {
            Watch(box, RequiredYes);
        }

        public static void WatchOptional(TextBox box)
        {
            Watch(box, RequiredNo);
        }

        public static void Watch(TextBox box, Func<bool> required)
        {
            if (box == null)
                return;

            State state;
            if (States.TryGetValue(box, out state))
            {
                state.Required = required ?? RequiredYes;
                Paint(box);
                return;
            }

            state = new State
            {
                Hint = box.ToolTip,
                Required = required ?? RequiredYes
            };
            States.Add(box, state);
            box.TextChanged += OnChanged;
            Paint(box);
        }

        public static void Refresh(TextBox box)
        {
            Paint(box);
        }

        public static bool Ok(TextBox box)
        {
            return ErrorText(box) == null;
        }

        public static bool Check(TextBox box)
        {
            Paint(box);
            if (Ok(box))
                return true;
            if (box != null)
                box.Focus();
            return false;
        }

        static void OnChanged(object sender, TextChangedEventArgs e)
        {
            Paint(sender as TextBox);
        }

        static void Paint(TextBox box)
        {
            if (box == null)
                return;
            State state;
            if (!States.TryGetValue(box, out state))
                return;

            string error = ErrorText(box);
            if (error != null)
            {
                box.BorderBrush = ErrorBorder;
                box.BorderThickness = new Thickness(1.5);
                box.Background = ErrorFill;
                box.ToolTip = error;
                ToolTipService.SetInitialShowDelay(box, 200);
                ToolTipService.SetShowDuration(box, 20000);
            }
            else
            {
                box.ClearValue(Control.BorderBrushProperty);
                box.ClearValue(Control.BorderThicknessProperty);
                box.ClearValue(Control.BackgroundProperty);
                box.ToolTip = state.Hint;
            }
        }

        static string ErrorText(TextBox box)
        {
            if (box == null)
                return "enter a formula like $A * $B.";
            State state;
            bool need = true;
            if (States.TryGetValue(box, out state) && state.Required != null)
                need = state.Required();
            if (string.IsNullOrWhiteSpace(box.Text) && !need)
                return null;
            return ExpressionEngine.Validate(box.Text);
        }

        static bool RequiredYes()
        {
            return true;
        }

        static bool RequiredNo()
        {
            return false;
        }

        static Brush Freeze(byte r, byte g, byte b)
        {
            var brush = new SolidColorBrush(Color.FromRgb(r, g, b));
            brush.Freeze();
            return brush;
        }

        sealed class State
        {
            public object Hint;
            public Func<bool> Required;
        }
    }
}
