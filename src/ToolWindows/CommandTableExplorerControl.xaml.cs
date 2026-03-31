using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;

namespace CommandTableInfo.ToolWindows
{
    public partial class CommandTableExplorerControl : UserControl, IDisposable
    {
        private readonly CommandTableExplorerDTO _dto;
        private readonly EnvDTE.CommandEvents _cmdEvents;
        private bool _hasUsedInspectMode;
        private CollectionView _view;
        private bool _isDisposed;

        public CommandTableExplorerControl(CommandTableExplorerDTO dto)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
            _dto = dto;
            _cmdEvents = dto.DTE.Events.CommandEvents;
            Commands = _dto.DteCommands;
            DataContext = this;
            Loaded += OnLoaded;

            InitializeComponent();
        }

        public IEnumerable<EnvDTE.Command> Commands { get; }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            if (_view == null)
            {
                details.Visibility = Visibility.Hidden;
                _view = (CollectionView)CollectionViewSource.GetDefaultView(list.ItemsSource);
                _view.Filter = UserFilter;
            }
        }

        private void OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
            var cmd = (EnvDTE.Command)list.SelectedValue;

            ResetDetails();

            if (cmd != null)
            {
                txtName.Content = cmd.Name;
                txtGuid.Text = cmd.Guid;
                txtId.Text = "0x" + cmd.ID.ToString("x") + $" ({cmd.ID})";
                txtBindings.Text = string.Join(Environment.NewLine, GetBindings(cmd.Bindings as object[]));
                
                details.Visibility = Visibility.Visible;
            }

            UpdateCopyButtonsVisibility();

        }

        private static IEnumerable<string> GetBindings(IEnumerable<object> bindings)
        {
            IEnumerable<string> result = bindings.Select(binding => binding.ToString().IndexOf("::") >= 0
                ? binding.ToString().Substring(binding.ToString().IndexOf("::") + 2)
                : binding.ToString()).Distinct();

            return result;
        }

        private bool UserFilter(object item)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();

            if (string.IsNullOrEmpty(txtFilter.Text))
            {
                return true;
            }
            else
            {
                var cmd = (EnvDTE.Command)item;
                return cmd.Name.IndexOf(txtFilter.Text, StringComparison.OrdinalIgnoreCase) >= 0 ||
                       cmd.Guid.IndexOf(txtFilter.Text, StringComparison.OrdinalIgnoreCase) >= 0;
            }
        }

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            string text = ((TextBox)sender).Text;
            _ = RefreshAsync(text).ConfigureAwait(false);
        }

        private async Task RefreshAsync(string text)
        {
            await Task.Delay(300);
            await Microsoft.VisualStudio.Shell.ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            if (text == txtFilter.Text)
            {
                try
                {
                    _view.Refresh();
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.Write(ex);
                }

                EnvDTE.Command cmd = _dto.DteCommands.FirstOrDefault(c => { Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread(); return c.Name.Equals(text.Trim(), StringComparison.OrdinalIgnoreCase); });

                if (cmd != null)
                {
                    list.SelectedItem = cmd;
                    list.ScrollIntoView(list.SelectedItem);
                }
            }
        }

        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
            var cb = (CheckBox)sender;

            if (cb.IsChecked == true)
            {
                _cmdEvents.BeforeExecute += CommandEvents_BeforeExecute;

                if (!_hasUsedInspectMode)
                {
                    MessageBox.Show("Hold down Ctrl+Shift and execute any command in any menu or toolbar to inspect", Vsix.Name, MessageBoxButton.OK, MessageBoxImage.Information);
                    _hasUsedInspectMode = true;
                }
            }
            else
            {
                _cmdEvents.BeforeExecute -= CommandEvents_BeforeExecute;
            }
        }

        private void CommandEvents_BeforeExecute(string Guid, int ID, object CustomIn, object CustomOut, ref bool CancelDefault)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();

            if ((Keyboard.IsKeyDown(Key.LeftShift) || Keyboard.IsKeyDown(Key.RightShift)) &&
                (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl)))
            {
                CancelDefault = true;
                EnvDTE.Command cmd = _dto.DteCommands.FirstOrDefault(c => { Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread(); return c.Guid == Guid && c.ID == ID; });

                if (cmd != null)
                {
                    txtFilter.Text = cmd.Name;
                }
            }
        }

        private void Hyperlink_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
            var cmd = (EnvDTE.Command)list.SelectedValue;

            try
            {
                if (cmd != null && cmd.IsAvailable)
                {
                    _dto.DTE.Commands.Raise(cmd.Guid, cmd.ID, null, null);
                    _dto.DTE.StatusBar.Clear();
                }
                else
                {
                    _dto.DTE.StatusBar.Text = $"The command '{cmd.Name}' is not available in the current context";
                }
            }
            catch (Exception)
            {
                _dto.DTE.StatusBar.Text = $"The command '{cmd.Name}' is not available in the current context";
            }
        }

        private void CopyName_Click(object sender, RoutedEventArgs e)
        {
            CopyValueToClipboard(txtName.Content?.ToString() ?? string.Empty);
        }

        private void CopyGuid_Click(object sender, RoutedEventArgs e)
        {
            CopyValueToClipboard(txtGuid.Text);
        }

        private void CopyId_Click(object sender, RoutedEventArgs e)
        {
            CopyValueToClipboard(txtId.Text);
        }

        private void CopyBindings_Click(object sender, RoutedEventArgs e)
        {
            CopyValueToClipboard(txtBindings.Text);
        }

        private void CopyValueToClipboard(string value)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();

            if (string.IsNullOrWhiteSpace(value))
            {
                return;
            }

            Clipboard.SetText(value);
            _dto.DTE.StatusBar.Text = "Value copied to clipboard";
        }

        private void UpdateCopyButtonsVisibility()
        {
            btnCopyName.Visibility = string.IsNullOrWhiteSpace(txtName.Content?.ToString()) ? Visibility.Collapsed : Visibility.Visible;
            btnCopyGuid.Visibility = string.IsNullOrWhiteSpace(txtGuid.Text) ? Visibility.Collapsed : Visibility.Visible;
            btnCopyId.Visibility = string.IsNullOrWhiteSpace(txtId.Text) ? Visibility.Collapsed : Visibility.Visible;
            btnCopyBindings.Visibility = string.IsNullOrWhiteSpace(txtBindings.Text) ? Visibility.Collapsed : Visibility.Visible;
        }

        private void ResetDetails()
        {
            txtName.Content = "loading...";
            txtGuid.Text = "n/a";
            txtId.Text = "n/a";
            txtBindings.Text = "n/a";

            UpdateCopyButtonsVisibility();

            details.Visibility = Visibility.Hidden;
        }

        public void Dispose()
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
            if (!_isDisposed)
            {
                _cmdEvents.BeforeExecute -= CommandEvents_BeforeExecute;
                Loaded -= OnLoaded;
            }

            _isDisposed = true;
        }
    }
}
