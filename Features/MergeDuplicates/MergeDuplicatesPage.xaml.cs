using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using Forms = System.Windows.Forms;
using WorkAssistant.Expressions;
using WorkAssistant.Views;

namespace WorkAssistant.Features.MergeDuplicates
{
    public partial class MergeDuplicatesPage : Page
    {
        List<MergeDuplicatesGroup> _groups;
        List<string> _targets;
        List<List<MemberPick>> _memberPicks;
        List<MergeDuplicatesPlan> _plans;
        bool _syncTarget;

        sealed class MemberPick
        {
            public string Name { get; set; }
            public string Label { get; set; }
            public bool IsTarget { get; set; }
            public bool Include { get; set; }
            public bool CanMerge { get { return !IsTarget; } }
        }

        public MergeDuplicatesPage()
        {
            InitializeComponent();
        }

        void Home_Click(object sender, RoutedEventArgs e)
        {
            if (NavigationService != null && NavigationService.CanGoBack)
                NavigationService.GoBack();
            else if (NavigationService != null)
                NavigationService.Navigate(new HomePage());
        }

        void BrowseParent_Click(object sender, RoutedEventArgs e)
        {
            using (var dlg = new Forms.FolderBrowserDialog())
            {
                dlg.Description = "Parent folder holding the duplicates";
                dlg.ShowNewFolderButton = false;
                if (dlg.ShowDialog() == Forms.DialogResult.OK)
                    PathParent.Text = dlg.SelectedPath;
            }
        }

        void Variables_Click(object sender, RoutedEventArgs e)
        {
            ExpressionHelp.Show(Window.GetWindow(this),
                "Match two folders. {A} and {B} are the folder names. After the separator split, {A1} is part 1 of A, {B2} part 2 of B. $A1 and $B1 work the same. {Today} and functions like CONTAINS() also work.",
                new[]
                {
                    new VariableHelp { Name = "{A}", Description = "Name of folder A." },
                    new VariableHelp { Name = "{B}", Description = "Name of folder B." },
                    new VariableHelp { Name = "{A1}", Description = "Split part 1 of folder A. {A2} is part 2." },
                    new VariableHelp { Name = "{B1}", Description = "Split part 1 of folder B." },
                    new VariableHelp { Name = "$A1", Description = "Same as {A1}." },
                    new VariableHelp { Name = "$B1", Description = "Same as {B1}." },
                });
        }

        void SepBox_Changed(object sender, SelectionChangedEventArgs e)
        {
            if (CustomSep != null)
                CustomSep.IsEnabled = SepBox.SelectedIndex == 3;
        }

        void Preview_Click(object sender, RoutedEventArgs e)
        {
            char sep;
            if (!TryReady(out sep))
                return;

            try
            {
                _groups = MergeDuplicatesWork.FindGroups(PathParent.Text, sep, ConditionBox.Text);
                _targets = new List<string>(_groups.Count);
                _memberPicks = new List<List<MemberPick>>(_groups.Count);
                foreach (var g in _groups)
                {
                    _targets.Add(g.Members[0].Name);
                    var picks = new List<MemberPick>(g.Members.Count);
                    foreach (var m in g.Members)
                    {
                        bool isTarget = string.Equals(m.Name, g.Members[0].Name, StringComparison.OrdinalIgnoreCase);
                        picks.Add(new MemberPick
                        {
                            Name = m.Name,
                            Label = m.Name + "  (" + m.Parts.Length + " parts)" + (isTarget ? " [keep]" : ""),
                            IsTarget = isTarget,
                            Include = !isTarget
                        });
                    }
                    _memberPicks.Add(picks);
                }
                _plans = null;
                PlansGrid.ItemsSource = null;
                RefreshGroups();
            }
            catch (Exception ex)
            {
                Alert(ex.Message, "Could not preview", MessageBoxImage.Error);
            }
        }

        void RefreshGroups()
        {
            GroupsList.ItemsSource = null;
            if (_groups == null || _groups.Count == 0)
            {
                GroupCount.Text = "No duplicate groups found.";
                MembersList.ItemsSource = null;
                TargetBox.ItemsSource = null;
                return;
            }

            var labels = new List<string>(_groups.Count);
            for (var i = 0; i < _groups.Count; i++)
            {
                var names = new List<string>();
                foreach (var m in _groups[i].Members)
                    names.Add(m.Name);
                labels.Add("Group " + (i + 1) + " (" + _groups[i].Members.Count + "): " + string.Join(", ", names));
            }
            GroupsList.ItemsSource = labels;
            GroupsList.SelectedIndex = 0;
            GroupCount.Text = _groups.Count + (_groups.Count == 1 ? " group" : " groups") + " found. Pick which folder to keep in each group.";
        }

        void GroupsList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var index = GroupsList.SelectedIndex;
            if (_groups == null || index < 0 || index >= _groups.Count)
            {
                MembersList.ItemsSource = null;
                TargetBox.ItemsSource = null;
                return;
            }

            var g = _groups[index];
            MembersList.ItemsSource = null;
            MembersList.ItemsSource = _memberPicks[index];

            _syncTarget = true;
            try
            {
                var targets = new List<string>();
                foreach (var m in g.Members)
                    targets.Add(m.Name);
                TargetBox.ItemsSource = targets;
                TargetBox.SelectedItem = _targets[index];
                if (TargetBox.SelectedIndex < 0)
                    TargetBox.SelectedIndex = 0;
            }
            finally
            {
                _syncTarget = false;
            }
        }

        void TargetBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_syncTarget)
                return;
            var index = GroupsList.SelectedIndex;
            if (_groups == null || index < 0 || index >= _groups.Count)
                return;
            var picked = TargetBox.SelectedItem as string;
            if (string.IsNullOrEmpty(picked))
                return;
            var oldTarget = _targets[index];
            _targets[index] = picked;
            var picks = _memberPicks[index];
            var grp = _groups[index];
            foreach (var pick in picks)
            {
                bool isTarget = string.Equals(pick.Name, picked, StringComparison.OrdinalIgnoreCase);
                bool wasTarget = string.Equals(pick.Name, oldTarget, StringComparison.OrdinalIgnoreCase);
                pick.IsTarget = isTarget;
                if (isTarget)
                    pick.Include = false;
                else if (wasTarget)
                    pick.Include = true;
                int partCount = 0;
                foreach (var m in grp.Members)
                {
                    if (string.Equals(m.Name, pick.Name, StringComparison.OrdinalIgnoreCase))
                    {
                        partCount = m.Parts.Length;
                        break;
                    }
                }
                pick.Label = pick.Name + "  (" + partCount + " parts)" + (isTarget ? " [keep]" : "");
            }
            MembersList.ItemsSource = null;
            MembersList.ItemsSource = picks;
        }

        void SelectAll_Click(object sender, RoutedEventArgs e)
        {
            SetAllPicks(true);
        }

        void SelectNone_Click(object sender, RoutedEventArgs e)
        {
            SetAllPicks(false);
        }

        void SetAllPicks(bool include)
        {
            var index = GroupsList.SelectedIndex;
            if (_memberPicks == null || index < 0 || index >= _memberPicks.Count)
                return;
            foreach (var pick in _memberPicks[index])
            {
                if (!pick.IsTarget)
                    pick.Include = include;
            }
            MembersList.ItemsSource = null;
            MembersList.ItemsSource = _memberPicks[index];
        }

        void Merge_Click(object sender, RoutedEventArgs e)
        {
            if (_groups == null || _groups.Count == 0)
            {
                Alert("Preview first.", "No groups", MessageBoxImage.Warning);
                return;
            }

            var selected = new List<IList<string>>(_groups.Count);
            for (var i = 0; i < _groups.Count; i++)
            {
                var names = new List<string>();
                if (_memberPicks != null && i < _memberPicks.Count && _memberPicks[i] != null)
                {
                    foreach (var pick in _memberPicks[i])
                    {
                        if (pick.Include && !pick.IsTarget)
                            names.Add(pick.Name);
                    }
                }
                else
                {
                    foreach (var m in _groups[i].Members)
                    {
                        if (!string.Equals(m.Name, _targets[i], StringComparison.OrdinalIgnoreCase))
                            names.Add(m.Name);
                    }
                }
                selected.Add(names);
            }

            List<MergeDuplicatesPlan> plans;
            try
            {
                plans = MergeDuplicatesWork.BuildMergePlans(_groups, _targets, selected);
            }
            catch (Exception ex)
            {
                Alert(ex.Message, "Could not plan merge", MessageBoxImage.Error);
                return;
            }

            if (plans.Count == 0)
            {
                Alert("Nothing to merge.", "Done", MessageBoxImage.Information);
                return;
            }

            var ask = MessageBox.Show(Window.GetWindow(this),
                "Merge " + plans.Count + (plans.Count == 1 ? " folder" : " folders") +
                " into the chosen targets? Files move into the kept folder; name clashes are renamed like \"name (2).ext\".",
                "Merge folders", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (ask != MessageBoxResult.Yes)
                return;

            MergeDuplicatesWork.ExecuteMerge(plans);
            _plans = plans;
            PlansGrid.ItemsSource = null;
            PlansGrid.ItemsSource = _plans;

            var merged = 0;
            var failed = 0;
            foreach (var p in plans)
            {
                if (p.Status != null && p.Status.StartsWith("Merged", StringComparison.Ordinal))
                    merged++;
                else if (p.Status != null && p.Status.StartsWith("Failed", StringComparison.Ordinal))
                    failed++;
            }
            Alert("Merged " + merged + " of " + plans.Count + ", failed " + failed + ".",
                "Done", failed > 0 ? MessageBoxImage.Warning : MessageBoxImage.Information);
        }

        bool TryReady(out char separator)
        {
            separator = SelectedSeparator();
            if (string.IsNullOrWhiteSpace(PathParent.Text))
            {
                Alert("Choose the parent folder.", "Missing folder", MessageBoxImage.Warning);
                return false;
            }
            if (SepBox.SelectedIndex == 3 && CustomSep.Text.Length == 0)
            {
                Alert("Enter a custom separator.", "Missing separator", MessageBoxImage.Warning);
                return false;
            }
                if (string.IsNullOrWhiteSpace(ConditionBox.Text))
            {
                Alert("Enter a match formula such as {A1} == {B1}.", "Missing match", MessageBoxImage.Warning);
                return false;
            }
            try
            {
                MergeDuplicatesWork.CheckCondition(ConditionBox.Text);
            }
            catch (Exception ex)
            {
                Alert(ex.Message, "Bad condition", MessageBoxImage.Warning);
                return false;
            }
            return true;
        }

        char SelectedSeparator()
        {
            switch (SepBox.SelectedIndex)
            {
                case 0: return ' ';
                case 2: return '-';
                case 3: return CustomSep.Text.Length > 0 ? CustomSep.Text[0] : '_';
                default: return '_';
            }
        }

        void Alert(string message, string title, MessageBoxImage icon)
        {
            MessageBox.Show(Window.GetWindow(this), message, title, MessageBoxButton.OK, icon);
        }
    }
}
