using Microsoft.VisualStudio;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;

namespace CommandTableInfo.ToolWindows
{
    public partial class CommandTableExplorerControl : UserControl, IDisposable
    {
        private static readonly IDictionary<Guid, string> KnownCommandSetNames = CreateKnownCommandSetNameMap();
        private static readonly IDictionary<Guid, IDictionary<int, string>> KnownCommandIds = CreateKnownCommandIdMap();
        private readonly CommandTableExplorerDTO _dto;
        private readonly EnvDTE.CommandEvents _cmdEvents;
        private readonly Dictionary<EnvDTE.Command, CommandSearchIndex> _searchIndex;
        private readonly Dictionary<string, EnvDTE.Command> _commandByName;
        private bool _hasUsedInspectMode;
        private CollectionView _view;
        private bool _isDisposed;
        private CancellationTokenSource _refreshCancellationTokenSource;
        private string _cachedFilterText;
        private string _cachedNormalizedFilterText;
        private string _cachedHexFilterText;
        private string _cachedNormalizedBindingFilterText;

        public CommandTableExplorerControl(CommandTableExplorerDTO dto)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
            _dto = dto;
            _cmdEvents = dto.DTE.Events.CommandEvents;
            Commands = _dto.DteCommands;
            _searchIndex = new Dictionary<EnvDTE.Command, CommandSearchIndex>();
            _commandByName = new Dictionary<string, EnvDTE.Command>(StringComparer.OrdinalIgnoreCase);
            _cachedFilterText = string.Empty;
            _cachedNormalizedFilterText = string.Empty;
            _cachedHexFilterText = string.Empty;
            _cachedNormalizedBindingFilterText = string.Empty;

            foreach (EnvDTE.Command command in _dto.DteCommands)
            {
                _searchIndex[command] = CreateSearchIndex(command);

                if (!_commandByName.ContainsKey(command.Name))
                {
                    _commandByName.Add(command.Name, command);
                }
            }

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
                txtGuid.Text = FormatGuidValue(cmd.Guid);
                txtId.Text = FormatIdValue(cmd.Guid, cmd.ID);
                txtGuid.Tag = cmd.Guid;
                txtId.Tag = "0x" + cmd.ID.ToString("x", CultureInfo.InvariantCulture) + " (" + cmd.ID.ToString(CultureInfo.InvariantCulture) + ")";
                txtBindings.Text = string.Join(Environment.NewLine, GetBindings(cmd.Bindings as object[]));

                details.Visibility = Visibility.Visible;
            }

            UpdateCopyButtonsVisibility();

        }

        private static IEnumerable<string> GetBindings(IEnumerable<object> bindings)
        {
            if (bindings == null)
            {
                return Enumerable.Empty<string>();
            }

            IEnumerable<string> result = bindings
                .Where(binding => binding != null)
                .Select(binding => binding.ToString())
                .Where(binding => !string.IsNullOrWhiteSpace(binding))
                .Select(binding => binding.IndexOf("::", StringComparison.Ordinal) >= 0
                    ? binding.Substring(binding.IndexOf("::", StringComparison.Ordinal) + 2)
                    : binding)
                .Distinct(StringComparer.OrdinalIgnoreCase);

            return result;
        }

        private bool UserFilter(object item)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();

            if (string.IsNullOrWhiteSpace(txtFilter.Text))
            {
                return true;
            }
            else
            {
                var cmd = (EnvDTE.Command)item;
                string filterText = txtFilter.Text.Trim();
                EnsureFilterCache(filterText);

                if (!_searchIndex.TryGetValue(cmd, out CommandSearchIndex searchIndex))
                {
                    searchIndex = CreateSearchIndex(cmd);
                    _searchIndex[cmd] = searchIndex;
                }

                if (searchIndex.Name.IndexOf(filterText, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }

                if (searchIndex.Guid.IndexOf(filterText, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }

                if (!string.IsNullOrEmpty(_cachedNormalizedFilterText) &&
                    searchIndex.NormalizedGuid.IndexOf(_cachedNormalizedFilterText, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }

                if (searchIndex.IdAsDecimal.IndexOf(filterText, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    searchIndex.IdAsHex.IndexOf(_cachedHexFilterText, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    searchIndex.IdAsHexPrefixed.IndexOf(filterText, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }

                if (string.IsNullOrEmpty(_cachedNormalizedBindingFilterText))
                {
                    return false;
                }

                return searchIndex.NormalizedBindings.Any(binding =>
                    binding.IndexOf(_cachedNormalizedBindingFilterText, StringComparison.OrdinalIgnoreCase) >= 0);
            }
        }

        private void EnsureFilterCache(string filterText)
        {
            if (string.Equals(_cachedFilterText, filterText, StringComparison.Ordinal))
            {
                return;
            }

            _cachedFilterText = filterText;
            _cachedNormalizedFilterText = NormalizeGuidForSearch(filterText);
            _cachedNormalizedBindingFilterText = NormalizeBindingForSearch(filterText);
            _cachedHexFilterText = filterText.StartsWith("0x", StringComparison.OrdinalIgnoreCase)
                ? filterText.Substring(2)
                : filterText;
        }

        private static CommandSearchIndex CreateSearchIndex(EnvDTE.Command command)
        {
            var bindings = GetBindings(command.Bindings as object[])
                .Select(NormalizeBindingForSearch)
                .Where(binding => !string.IsNullOrEmpty(binding))
                .ToArray();

            string idAsHex = command.ID.ToString("x", CultureInfo.InvariantCulture);

            return new CommandSearchIndex(
                command.Name ?? string.Empty,
                command.Guid ?? string.Empty,
                NormalizeGuidForSearch(command.Guid),
                command.ID.ToString(CultureInfo.InvariantCulture),
                idAsHex,
                "0x" + idAsHex,
                bindings);
        }

        private static string NormalizeGuidForSearch(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }

            return value
                .Replace("-", string.Empty)
                .Replace("{", string.Empty)
                .Replace("}", string.Empty)
                .Trim();
        }

        private static string NormalizeBindingForSearch(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }

            return value.Replace(" ", string.Empty).Trim();
        }

        private static string FormatGuidValue(string guidText)
        {
            if (!Guid.TryParse(guidText, out Guid guid))
            {
                return guidText;
            }

            if (KnownCommandSetNames.TryGetValue(guid, out string knownName))
            {
                return string.Format(CultureInfo.InvariantCulture, "{0} ({1})", knownName, guidText);
            }

            return guidText;
        }

        private static string FormatIdValue(string guidText, int id)
        {
            string numericValue = "0x" + id.ToString("x", CultureInfo.InvariantCulture) + " (" + id.ToString(CultureInfo.InvariantCulture) + ")";

            if (!Guid.TryParse(guidText, out Guid guid))
            {
                return numericValue;
            }

            if (KnownCommandIds.TryGetValue(guid, out IDictionary<int, string> commandIds) &&
                commandIds.TryGetValue(id, out string knownIdName))
            {
                return string.Format(CultureInfo.InvariantCulture, "{0} ({1})", knownIdName, numericValue);
            }

            return numericValue;
        }

        private static IDictionary<Guid, string> CreateKnownCommandSetNameMap()
        {
            var map = new Dictionary<Guid, string>();

            AddKnownCommandSetName(map, VSConstants.GUID_VSStandardCommandSet97, "guidVSStd97");
            AddKnownCommandSetName(map, VSConstants.VSStd2K, "guidVSStd2K");

            foreach (FieldInfo field in typeof(VSConstants).GetFields(BindingFlags.Public | BindingFlags.Static))
            {
                if (field.FieldType != typeof(Guid))
                {
                    continue;
                }

                Guid guid = (Guid)field.GetValue(null);

                if (!map.ContainsKey(guid))
                {
                    map.Add(guid, field.Name);
                }
            }

            return map;
        }

        private static IDictionary<Guid, IDictionary<int, string>> CreateKnownCommandIdMap()
        {
            var map = new Dictionary<Guid, IDictionary<int, string>>();

            AddKnownCommandIds(map, VSConstants.GUID_VSStandardCommandSet97, typeof(VSConstants.VSStd97CmdID));
            AddKnownCommandIds(map, VSConstants.VSStd2K, typeof(VSConstants.VSStd2KCmdID));

            return map;
        }

        private static void AddKnownCommandIds(IDictionary<Guid, IDictionary<int, string>> map, Guid commandSet, Type enumType)
        {
            var commandIds = new Dictionary<int, string>();

            foreach (object value in Enum.GetValues(enumType))
            {
                int commandId = Convert.ToInt32(value, CultureInfo.InvariantCulture);

                if (!commandIds.ContainsKey(commandId))
                {
                    commandIds.Add(commandId, Enum.GetName(enumType, value));
                }
            }

            if (!map.ContainsKey(commandSet))
            {
                map.Add(commandSet, commandIds);
            }
        }

        private static void AddKnownCommandSetName(IDictionary<Guid, string> map, Guid guid, string name)
        {
            if (!map.ContainsKey(guid))
            {
                map.Add(guid, name);
            }
        }

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            string text = ((TextBox)sender).Text;

            _refreshCancellationTokenSource?.Cancel();
            _refreshCancellationTokenSource?.Dispose();
            _refreshCancellationTokenSource = new CancellationTokenSource();

            _ = RefreshAsync(text, _refreshCancellationTokenSource.Token);
        }

        private async Task RefreshAsync(string text, CancellationToken cancellationToken)
        {
            try
            {
                await Task.Delay(300, cancellationToken);
                await Microsoft.VisualStudio.Shell.ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
            }
            catch (OperationCanceledException)
            {
                return;
            }

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

                _commandByName.TryGetValue(text.Trim(), out EnvDTE.Command cmd);

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

            if (cmd == null)
            {
                _dto.DTE.StatusBar.Text = "No command selected";
                return;
            }

            try
            {
                if (cmd.IsAvailable)
                {
                    _dto.DTE.Commands.Raise(cmd.Guid, cmd.ID, null, null);
                    _dto.DTE.StatusBar.Clear();
                }
                else
                {
                    _dto.DTE.StatusBar.Text = $"The command '{cmd.Name}' is not available in the current context";
                }
            }
            catch (Exception ex)
            {
                _dto.DTE.StatusBar.Text = $"Failed to execute '{cmd.Name}': {ex.Message}";
            }
        }

        private void CopyGuid_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
            CopyValueToClipboard(txtGuid.Tag as string ?? txtGuid.Text);
        }

        private void CopyId_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
            CopyValueToClipboard(txtId.Tag as string ?? txtId.Text);
        }

        private void CopyBindings_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
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
            txtGuid.Tag = null;
            txtId.Tag = null;

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
                _refreshCancellationTokenSource?.Cancel();
                _refreshCancellationTokenSource?.Dispose();
                _refreshCancellationTokenSource = null;
            }

            _isDisposed = true;
        }

        private sealed class CommandSearchIndex
        {
            public CommandSearchIndex(string name, string guid, string normalizedGuid, string idAsDecimal, string idAsHex, string idAsHexPrefixed, string[] normalizedBindings)
            {
                Name = name;
                Guid = guid;
                NormalizedGuid = normalizedGuid;
                IdAsDecimal = idAsDecimal;
                IdAsHex = idAsHex;
                IdAsHexPrefixed = idAsHexPrefixed;
                NormalizedBindings = normalizedBindings;
            }

            public string Name { get; }
            public string Guid { get; }
            public string NormalizedGuid { get; }
            public string IdAsDecimal { get; }
            public string IdAsHex { get; }
            public string IdAsHexPrefixed { get; }
            public string[] NormalizedBindings { get; }
        }
    }
}
