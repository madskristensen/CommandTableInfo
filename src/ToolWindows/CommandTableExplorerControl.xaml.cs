using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using CommandTable;

namespace CommandTableInfo.ToolWindows
{
    public partial class CommandTableExplorerControl : UserControl
    {
        private readonly CommandTableExplorerDTO _dto;
        private Task<IEnumerable<Command>> _commands;

        public CommandTableExplorerControl(CommandTableExplorerDTO dto)
        {
            _dto = dto;
            Commands = _dto.DteCommands;
            DataContext = this;
            Loaded += OnLoaded;

            InitializeComponent();
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
            CommandTreeItem.ItemSelected += CommandTreeItem_ItemSelected;

            _commands = _dto.CommandTable.GetCommands();
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
            txtId.Text =  "0x" + cmd.ID.ToString("x") + $" ({cmd.ID})";
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
            System.Windows.Forms.MessageBox.Show(cb.IsChecked.ToString());
        }
    }
}
