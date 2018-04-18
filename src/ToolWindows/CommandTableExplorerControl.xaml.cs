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
            details.Visibility = Visibility.Hidden;

            var view = (CollectionView)CollectionViewSource.GetDefaultView(list.ItemsSource);
            view.Filter = UserFilter;
            list.SelectionChanged += OnSelectionChanged;

            _commands = _dto.CommandTable.GetCommands();
        }

        private void OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var cmd = (EnvDTE.Command)list.SelectedValue;

            txtName.Text = cmd.Name;
            txtLocalizedName.Text = cmd.LocalizedName;
            txtGuid.Text = cmd.Guid;
            txtId.Text =  "0x" + cmd.ID.ToString("x") + $" ({cmd.ID})";
            lblBindings.Content = string.Join(Environment.NewLine, GetBindings(cmd.Bindings as object[]));

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
                    groupView.ItemsSource = command.Placements;//.Select(p => new CommandItemViewModel(p));
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
    }
}
