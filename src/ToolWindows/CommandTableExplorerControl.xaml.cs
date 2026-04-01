using Microsoft.VisualStudio;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using CommandTableInfo.Services;
using Microsoft.VisualStudio.Shell.Interop;

namespace CommandTableInfo.ToolWindows
{
    public partial class CommandTableExplorerControl : UserControl, IDisposable
    {
        private static CommandTableExplorerControl _contextMenuSourceControl;
        private static CommandTableExplorerControl _latestControl;
        private static readonly IDictionary<Guid, string> KnownCommandSetNames = CreateKnownCommandSetNameMap();
        private static readonly IDictionary<Guid, IDictionary<int, string>> KnownCommandIds = CreateKnownCommandIdMap();
        private static readonly Regex HierarchyGuidIdRegex = new Regex(@"\((?<guid>\{?[0-9a-fA-F-]{36}\}?)\s*:\s*0x(?<id>[0-9a-fA-F]+)", RegexOptions.Compiled);
        private readonly CommandTableExplorerDTO _dto;
        private readonly EnvDTE.CommandEvents _cmdEvents;
        private readonly Dictionary<EnvDTE.Command, CommandSearchIndex> _searchIndex;
        private readonly Dictionary<string, EnvDTE.Command> _commandByName;
        private EnvDTE.Command _selectedCommand;
        private string _hierarchyCopyText;
        private bool _hasUsedInspectMode;
        private CollectionView _view;
        private bool _isDisposed;
        private CancellationTokenSource _refreshCancellationTokenSource;
        private string _cachedFilterText;
        private string _cachedNormalizedFilterText;
        private string _cachedHexFilterText;
        private string _cachedNormalizedBindingFilterText;
        private FrameworkElement _contextMenuPlacementTarget;
        private HierarchyTreeNode _contextMenuHierarchyNode;

        internal CommandTableExplorerControl(CommandTableExplorerDTO dto)
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
            _hierarchyCopyText = string.Empty;
            _latestControl = this;
            _contextMenuSourceControl = this;

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
                _selectedCommand = cmd;
                txtName.Content = cmd.Name;
                txtGuid.Text = FormatGuidValue(cmd.Guid);
                txtId.Text = FormatIdValue(cmd.Guid, cmd.ID);
                txtGuid.Tag = cmd.Guid;
                txtId.Tag = "0x" + cmd.ID.ToString("x", CultureInfo.InvariantCulture) + " (" + cmd.ID.ToString(CultureInfo.InvariantCulture) + ")";
                txtGuidId.Text = string.Format(CultureInfo.InvariantCulture, "{0}:0x{1} ({2})", cmd.Guid, cmd.ID.ToString("x", CultureInfo.InvariantCulture), cmd.ID.ToString(CultureInfo.InvariantCulture));
                txtGuidId.Tag = string.Format(CultureInfo.InvariantCulture, "{0}:0x{1}", cmd.Guid, cmd.ID.ToString("x", CultureInfo.InvariantCulture));
                txtBindings.Text = string.Join(Environment.NewLine, GetBindings(cmd.Bindings as object[]));

                CommandHierarchyInfo hierarchy = _dto.CommandHierarchyService?.GetHierarchy(cmd) ?? CommandHierarchyInfo.Empty;
                txtDisplayName.Text = hierarchy.DisplayName;
                txtButtonText.Text = hierarchy.ButtonText;
                _hierarchyCopyText = hierarchy.HierarchyCopyText;
                treeHierarchy.ItemsSource = BuildHierarchyTree(hierarchy.HierarchyText);

                details.Visibility = Visibility.Visible;
            }

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

        internal static bool CanExecuteNativeContextMenuCommand(int commandId)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();

            if (_contextMenuSourceControl == null)
            {
                _contextMenuSourceControl = _latestControl;
            }

            if (_contextMenuSourceControl == null)
            {
                return false;
            }

            if (commandId == PackageIds.CopyContextValueId)
            {
                return _contextMenuSourceControl.CanExecuteCopyContextValue();
            }

            if (commandId == PackageIds.CopyVsctSymbolsId)
            {
                return _contextMenuSourceControl.CanExecuteCopyVsctSymbols();
            }

            return false;
        }

        internal static void ExecuteNativeContextMenuCommand(int commandId)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();

            if (_contextMenuSourceControl == null)
            {
                return;
            }

            if (commandId == PackageIds.CopyContextValueId)
            {
                _contextMenuSourceControl.ExecuteCopyContextValue();
                return;
            }

            if (commandId == PackageIds.CopyVsctSymbolsId)
            {
                _contextMenuSourceControl.ExecuteCopyVsctSymbols();
            }
        }

        private bool CanExecuteCopyContextValue()
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();

            if (_contextMenuPlacementTarget is TextBox textBox)
            {
                return !string.IsNullOrWhiteSpace(GetTextBoxCopyValue(textBox));
            }

            return !string.IsNullOrWhiteSpace(GetHierarchyContextText());
        }

        private bool HasGuidIdPairOnSelectedHierarchyNode()
        {
            return TryGetGuidAndIdFromHierarchyContext(out _, out _);
        }

        private bool CanExecuteCopyVsctSymbols()
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
            return HasGuidIdPairOnSelectedHierarchyNode();
        }

        private bool TryGetGuidAndIdFromHierarchyContext(out Guid commandSet, out int commandId)
        {
            commandSet = Guid.Empty;
            commandId = 0;

            string nodeText = _contextMenuHierarchyNode?.DisplayText;

            if (string.IsNullOrWhiteSpace(nodeText) && treeHierarchy?.SelectedItem is HierarchyTreeNode selectedNode)
            {
                nodeText = selectedNode.DisplayText;
            }

            if (string.IsNullOrWhiteSpace(nodeText))
            {
                return false;
            }

            Match match = HierarchyGuidIdRegex.Match(nodeText);

            if (!match.Success)
            {
                return false;
            }

            if (!Guid.TryParse(match.Groups["guid"].Value, out commandSet))
            {
                return false;
            }

            return int.TryParse(match.Groups["id"].Value, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out commandId);
        }

        private string GetHierarchyContextText()
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();

            string treeText = _contextMenuHierarchyNode?.DisplayText;

            if (string.IsNullOrWhiteSpace(treeText) && treeHierarchy?.SelectedItem is HierarchyTreeNode selectedNode)
            {
                treeText = selectedNode.DisplayText;
            }

            if (string.IsNullOrWhiteSpace(treeText))
            {
                treeText = _hierarchyCopyText;
            }

            return treeText;
        }

        private EnvDTE.Command GetCurrentCommandSelection()
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
            return _selectedCommand ?? list.SelectedValue as EnvDTE.Command;
        }

        private void ExecuteCopyVsctSymbols()
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();

            if (!CanExecuteCopyVsctSymbols())
            {
                return;
            }

            if (!TryGetGuidAndIdFromHierarchyContext(out Guid commandSet, out int commandId))
            {
                return;
            }

            EnvDTE.Command selectedCommand = GetCurrentCommandSelection();

            string guidSymbolName = GetGuidSymbolName(commandSet, selectedCommand);
            string idSymbolName = GetIdSymbolName(commandSet, commandId, guidSymbolName, selectedCommand);
            string snippet = string.Format(
                CultureInfo.InvariantCulture,
                "<GuidSymbol name=\"{0}\" value=\"{{{1}}}\">{4}  <IDSymbol name=\"{2}\" value=\"0x{3}\" />{4}</GuidSymbol>",
                guidSymbolName,
                commandSet,
                idSymbolName,
                commandId.ToString("x4", CultureInfo.InvariantCulture),
                Environment.NewLine);

            CopyValueToClipboard(snippet);
        }

        private void ExecuteCopyContextValue()
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();

            if (_contextMenuPlacementTarget is TextBox textBox)
            {
                CopyValueToClipboard(GetTextBoxCopyValue(textBox));
                return;
            }

            if (_contextMenuPlacementTarget is TreeView)
            {
                string value = GetHierarchyContextText();
                CopyValueToClipboard(value);
                return;
            }

            if (!string.IsNullOrWhiteSpace(_hierarchyCopyText))
            {
                CopyValueToClipboard(GetHierarchyContextText());
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

        private static string GetGuidSymbolName(EnvDTE.Command command)
        {
            Guid guid;

            if (!Guid.TryParse(command.Guid, out guid))
            {
                return "guidCommandSet";
            }

            if (guid == PackageGuids.guidCommandTablePackageCmdSet)
            {
                return "guidCommandTablePackageCmdSet";
            }

            string knownName;
            if (KnownCommandSetNames.TryGetValue(guid, out knownName) && !string.IsNullOrWhiteSpace(knownName))
            {
                return knownName;
            }

            return "guid" + ToIdentifier(command.Name) + "CmdSet";
        }

        private static string GetGuidSymbolName(Guid commandSet, EnvDTE.Command selectedCommand)
        {
            if (commandSet == PackageGuids.guidCommandTablePackageCmdSet)
            {
                return "guidCommandTablePackageCmdSet";
            }

            if (KnownCommandSetNames.TryGetValue(commandSet, out string knownName) && !string.IsNullOrWhiteSpace(knownName))
            {
                return knownName;
            }

            if (selectedCommand != null && Guid.TryParse(selectedCommand.Guid, out Guid selectedCommandSet) && selectedCommandSet == commandSet)
            {
                return GetGuidSymbolName(selectedCommand);
            }

            return "guidCommandSet";
        }

        private static string GetIdSymbolName(EnvDTE.Command command, string guidSymbolName)
        {
            Guid guid;
            if (Guid.TryParse(command.Guid, out guid))
            {
                IDictionary<int, string> ids;
                string knownId;

                if (KnownCommandIds.TryGetValue(guid, out ids) &&
                    ids.TryGetValue(command.ID, out knownId) &&
                    !string.IsNullOrWhiteSpace(knownId))
                {
                    return knownId;
                }

                if (guid == PackageGuids.guidCommandTablePackageCmdSet && command.ID == PackageIds.ShowToolWindowId)
                {
                    return "ShowToolWindowId";
                }
            }

            string baseName = ToIdentifier(command.Name);

            if (string.IsNullOrWhiteSpace(baseName))
            {
                baseName = guidSymbolName + "Command";
            }

            return baseName.EndsWith("Id", StringComparison.OrdinalIgnoreCase)
                ? baseName
                : baseName + "Id";
        }

        private static string GetIdSymbolName(Guid commandSet, int commandId, string guidSymbolName, EnvDTE.Command selectedCommand)
        {
            if (KnownCommandIds.TryGetValue(commandSet, out IDictionary<int, string> ids) &&
                ids.TryGetValue(commandId, out string knownId) &&
                !string.IsNullOrWhiteSpace(knownId))
            {
                return knownId;
            }

            if (commandSet == PackageGuids.guidCommandTablePackageCmdSet)
            {
                if (commandId == PackageIds.ShowToolWindowId)
                {
                    return "ShowToolWindowId";
                }

                if (commandId == PackageIds.CopyContextValueId)
                {
                    return "CopyContextValueId";
                }

                if (commandId == PackageIds.CopyVsctSymbolsId)
                {
                    return "CopyVsctSymbolsId";
                }
            }

            if (selectedCommand != null && Guid.TryParse(selectedCommand.Guid, out Guid selectedCommandSet) && selectedCommandSet == commandSet && selectedCommand.ID == commandId)
            {
                return GetIdSymbolName(selectedCommand, guidSymbolName);
            }

            return guidSymbolName + "CommandId";
        }

        private static string ToIdentifier(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return "Command";
            }

            string token = value;
            int lastDot = token.LastIndexOf('.');
            if (lastDot >= 0 && lastDot < token.Length - 1)
            {
                token = token.Substring(lastDot + 1);
            }

            var chars = token.Where(char.IsLetterOrDigit).ToArray();
            if (chars.Length == 0)
            {
                return "Command";
            }

            if (!char.IsLetter(chars[0]) && chars[0] != '_')
            {
                return "Command" + new string(chars);
            }

            return new string(chars);
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

        private static string GetTextBoxCopyValue(TextBox textBox)
        {
            string value = textBox.Tag as string;

            if (string.IsNullOrWhiteSpace(value))
            {
                value = textBox.SelectedText;
            }

            if (string.IsNullOrWhiteSpace(value))
            {
                value = textBox.Text;
            }

            return value;
        }

        private void DetailsTextBox_PreviewMouseRightButtonUp(object sender, MouseButtonEventArgs e)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();

            if (sender is TextBox textBox)
            {
                SetContextMenuTarget(textBox, null);
                ShowNativeContextMenu(PackageIds.DetailsContextMenu, e);
                e.Handled = true;
            }
        }

        private void TreeHierarchy_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();

            var source = e.OriginalSource as DependencyObject;
            var item = FindAncestor<TreeViewItem>(source);

            if (item != null)
            {
                item.IsSelected = true;
                item.Focus();
                _contextMenuHierarchyNode = item.DataContext as HierarchyTreeNode;
                SetContextMenuTarget(treeHierarchy, _contextMenuHierarchyNode);
                e.Handled = true;
                return;
            }

            _contextMenuHierarchyNode = treeHierarchy.SelectedItem as HierarchyTreeNode;
            SetContextMenuTarget(treeHierarchy, _contextMenuHierarchyNode);
        }

        private void TreeHierarchy_PreviewMouseRightButtonUp(object sender, MouseButtonEventArgs e)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();

            _contextMenuHierarchyNode = _contextMenuHierarchyNode ?? treeHierarchy.SelectedItem as HierarchyTreeNode;
            SetContextMenuTarget(treeHierarchy, _contextMenuHierarchyNode);
            ShowNativeContextMenu(PackageIds.HierarchyContextMenu, e);
            e.Handled = true;
        }

        private void SetContextMenuTarget(FrameworkElement placementTarget, HierarchyTreeNode hierarchyNode)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();
            _contextMenuSourceControl = this;
            _contextMenuPlacementTarget = placementTarget;
            _contextMenuHierarchyNode = hierarchyNode ?? treeHierarchy?.SelectedItem as HierarchyTreeNode;
        }

        private void ShowNativeContextMenu(int menuId, MouseButtonEventArgs e)
        {
            Microsoft.VisualStudio.Shell.ThreadHelper.ThrowIfNotOnUIThread();

            if (_dto.VsUiShell == null)
            {
                return;
            }

            Point screenPoint = PointToScreen(e.GetPosition(this));
            var location = new[]
            {
                new POINTS
                {
                    x = ToPointValue(screenPoint.X),
                    y = ToPointValue(screenPoint.Y)
                }
            };

            Guid commandGroup = PackageGuids.guidCommandTablePackageCmdSet;
            _dto.VsUiShell.ShowContextMenu(0, ref commandGroup, menuId, location, null);
        }

        private static short ToPointValue(double value)
        {
            if (value < short.MinValue)
            {
                return short.MinValue;
            }

            if (value > short.MaxValue)
            {
                return short.MaxValue;
            }

            return Convert.ToInt16(Math.Round(value));
        }

        private static T FindAncestor<T>(DependencyObject source) where T : DependencyObject
        {
            while (source != null)
            {
                if (source is T match)
                {
                    return match;
                }

                source = System.Windows.Media.VisualTreeHelper.GetParent(source);
            }

            return null;
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

        private void ResetDetails()
        {
            txtName.Content = "loading...";
            txtGuid.Text = "n/a";
            txtId.Text = "n/a";
            txtGuidId.Text = "n/a";
            txtDisplayName.Text = "n/a";
            txtButtonText.Text = "n/a";
            txtBindings.Text = "n/a";
            _hierarchyCopyText = string.Empty;
            treeHierarchy.ItemsSource = BuildHierarchyTree(CommandHierarchyInfo.Empty.HierarchyText);
            txtGuid.Tag = null;
            txtId.Tag = null;
            txtGuidId.Tag = null;
            _selectedCommand = null;

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

                if (ReferenceEquals(_contextMenuSourceControl, this))
                {
                    _contextMenuSourceControl = null;
                }
            }

            _isDisposed = true;
        }

        private static IList<HierarchyTreeNode> BuildHierarchyTree(string hierarchyText)
        {
            var roots = new List<HierarchyTreeNode>();

            if (string.IsNullOrWhiteSpace(hierarchyText))
            {
                return roots;
            }

            string[] lines = hierarchyText
                .Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);

            if (lines.Length == 1 && lines[0].IndexOf("Not available", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                roots.Add(new HierarchyTreeNode(lines[0]));
                return roots;
            }

            foreach (string line in lines)
            {
                AddHierarchyPath(roots, line);
            }

            return roots;
        }

        private static void AddHierarchyPath(IList<HierarchyTreeNode> roots, string path)
        {
            string[] parts = path
                .Split(new[] { " > " }, StringSplitOptions.RemoveEmptyEntries);

            if (parts.Length == 0)
            {
                return;
            }

            IList<HierarchyTreeNode> level = roots;

            foreach (string part in parts)
            {
                HierarchyTreeNode existing = level.FirstOrDefault(node => string.Equals(node.DisplayText, part, StringComparison.OrdinalIgnoreCase));

                if (existing == null)
                {
                    existing = new HierarchyTreeNode(part);
                    level.Add(existing);
                }

                level = existing.Children;
            }
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

    public sealed class HierarchyTreeNode
    {
        public HierarchyTreeNode(string displayText)
        {
            DisplayText = displayText;
            Children = new List<HierarchyTreeNode>();
        }

        public string DisplayText { get; }

        public IList<HierarchyTreeNode> Children { get; }
    }
}
