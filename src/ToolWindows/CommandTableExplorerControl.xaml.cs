using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using CommandTable;

namespace CommandTableInfo.ToolWindows
{
    public partial class CommandTableExplorerControl : UserControl
    {
        private readonly CommandTableExplorerDTO _dto;
        private Task<IEnumerable<Command>> _commands;
        private EnvDTE.CommandEvents _cmdEvents;
        private bool _hasUsedInspectMode;

        public CommandTableExplorerControl(CommandTableExplorerDTO dto)
        {
            _dto = dto;
            _cmdEvents = dto.DTE.Events.CommandEvents;
            Commands = _dto.DteCommands;
            DataContext = this;
            Loaded += OnLoaded;
            Unloaded += OnUnloaded;

            InitializeComponent();
            CommandTreeItem.ItemSelected += CommandTreeItem_ItemSelected;
            _commands = _dto.CommandTable.GetCommands();
        }

        public IEnumerable<EnvDTE.Command> Commands { get; }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            txtFilter.Clear();

            details.Visibility = Visibility.Hidden;
            groupDetails.Visibility = Visibility.Hidden;

            var view = (CollectionView)CollectionViewSource.GetDefaultView(list.ItemsSource);
            view.Filter = UserFilter;
            list.SelectionChanged += OnSelectionChanged;
        }

        private void OnUnloaded(object sender, RoutedEventArgs e)
        {
            cbInspect.IsChecked = false;
            list.SelectionChanged -= OnSelectionChanged;
        }

        private void CommandTreeItem_ItemSelected(object sender, CommandTreeItem e)
        {
            txtPlacementSymbolicGuid.Text = e.Command.SymbolicItemId.SymbolicGuidName.Replace("guidSolutionExplorerMenu", "guidSHLMainMenu");
            txtPlacementSymbolicId.Text = e.Command.SymbolicItemId.SymbolicDWordName;
            txtPlacementGuid.Text = e.Command.ItemId.Guid.ToString();
            txtPlacementId.Text = "0x" + e.Command.ItemId.DWord.ToString("x") + $" ({e.Command.ItemId.DWord})";
            txtPlacementPriority.Text = "0x" + e.Command.Priority.ToString("x") + $" ({e.Command.Priority})";
            txtPlacementType.Text = "Group";

            groupDetails.Visibility = Visibility.Visible;

            if (e.Command is CommandContainer menu)
            {
                txtPlacementType.Text = menu.Type.ToString();
            }
        }

        private void OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var cmd = (EnvDTE.Command)list.SelectedValue;

            txtName.Content = cmd.Name;
            txtGuid.Text = cmd.Guid;
            txtId.Text = "0x" + cmd.ID.ToString("x") + $" ({cmd.ID})";
            txtBindings.Text = string.Join(Environment.NewLine, GetBindings(cmd.Bindings as object[]));

            PopulateGroupsAsync(cmd).ConfigureAwait(false);

            details.Visibility = Visibility.Visible;
        }

        private async Task PopulateGroupsAsync(EnvDTE.Command cmd)
        {
            if (_commands.IsCompleted)
            {
                IEnumerable<Command> commands = await _commands;
                Command command = commands.FirstOrDefault(c => c.ItemId.Guid == new Guid(cmd.Guid) && c.ItemId.DWord == cmd.ID);

                if (command != null)
                {
                    tree.ItemsSource = command.Placements.Select(p => new CommandTreeItem(p));

                    txtPriority.Text = "0x" + command.Priority.ToString("x") + $" ({command.Priority})";
                    txtPackage.Text = command.SourcePackageInfo.PackageName;
                    txtAssembly.Text = command.SourcePackageInfo.Assembly;
                }

                loading.Visibility = Visibility.Collapsed;
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
            if (string.IsNullOrEmpty(txtFilter.Text))
            {
                return true;
            }
            else
            {
                var cmd = (EnvDTE.Command)item;
                return cmd.Name.IndexOf(txtFilter.Text, StringComparison.OrdinalIgnoreCase) >= 0 ||
                       cmd.Guid.IndexOf(txtFilter.Text, StringComparison.OrdinalIgnoreCase) >= 0 ||
                       txtFilter.Text.Equals(cmd.ID);
            }
        }

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            string text = ((TextBox)sender).Text;
            RefreshAsync(text).ConfigureAwait(false);
        }

        private async Task RefreshAsync(string text)
        {
            await Task.Delay(300);

            if (text == txtFilter.Text)
            {
                CollectionViewSource.GetDefaultView(list.ItemsSource).Refresh();
            }
        }

        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {
            var cb = (CheckBox)sender;

            if (cb.IsChecked == true)
            {
                _cmdEvents.BeforeExecute += CommandEvents_BeforeExecute;
                txtFilter.Clear();

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
            if ((Keyboard.IsKeyDown(Key.LeftShift) || Keyboard.IsKeyDown(Key.RightShift)) &&
                (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl)))
            {
                CancelDefault = true;
                EnvDTE.Command cmd = _dto.DteCommands.FirstOrDefault(c => c.Guid == Guid && c.ID == ID);

                if (cmd != null)
                {
                    int index = _dto.DteCommands.IndexOf(cmd);
                    list.SelectedIndex = index;
                    list.ScrollIntoView(list.SelectedItem);
                }
            }
        }
    }
}
