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
            //CommandTreeItem.ItemSelected += CommandTreeItem_ItemSelected;
        }

        public IEnumerable<EnvDTE.Command> Commands { get; }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            if (_view == null)
            {
                details.Visibility = Visibility.Hidden;
                //groupDetails.Visibility = Visibility.Hidden;

                _view = (CollectionView)CollectionViewSource.GetDefaultView(list.ItemsSource);
                _view.Filter = UserFilter;
            }
        }

        //private void CommandTreeItem_ItemSelected(object sender, CommandTreeItem e)
        //{
        //    txtPlacementSymbolicGuid.Text = e.Command.SymbolicItemId.SymbolicGuidName.Replace("guidSolutionExplorerMenu", "guidSHLMainMenu");
        //    txtPlacementSymbolicId.Text = e.Command.SymbolicItemId.SymbolicDWordName;
        //    txtPlacementGuid.Text = e.Command.ItemId.Guid.ToString();
        //    txtPlacementId.Text = "0x" + e.Command.ItemId.DWord.ToString("x") + $" ({e.Command.ItemId.DWord})";
        //    txtPlacementPriority.Text = "0x" + e.Command.Priority.ToString("x") + $" ({e.Command.Priority})";
        //    txtPlacementType.Text = "n/a";

        //    groupDetails.Visibility = Visibility.Visible;

        //    //if (e.Command is CommandContainer menu)
        //    //{
        //    //    txtPlacementType.Text = menu.Type.ToString();
        //    //}
        //}

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

                EnvDTE.Command cmd = _dto.DteCommands.FirstOrDefault(c => c.Name.Equals(text.Trim(), StringComparison.OrdinalIgnoreCase));

                await Microsoft.VisualStudio.Shell.ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

                if (cmd != null)
                {
                    list.SelectedItem = cmd;
                    list.ScrollIntoView(list.SelectedItem);
                }

                _cmdEvents.BeforeExecute -= CommandEvents_BeforeExecute;

                if (cbInspect.IsChecked == true)
                {
                    _cmdEvents.BeforeExecute += CommandEvents_BeforeExecute;
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
                EnvDTE.Command cmd = _dto.DteCommands.FirstOrDefault(c => c.Guid == Guid && c.ID == ID);
                _cmdEvents.BeforeExecute -= CommandEvents_BeforeExecute;

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

        private void ResetDetails()
        {
            txtName.Content = "loading...";
            txtGuid.Text = "n/a";
            txtId.Text = "n/a";
            txtBindings.Text = "n/a";

            details.Visibility = Visibility.Hidden;
        }

        public void Dispose()
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
            if (!_isDisposed)
            {
                //CommandTreeItem.ItemSelected -= CommandTreeItem_ItemSelected;
                _cmdEvents.BeforeExecute -= CommandEvents_BeforeExecute;
                Loaded -= OnLoaded;

               // _dto.CommandTable = null;
            }

            _isDisposed = true;
        }
    }
}
