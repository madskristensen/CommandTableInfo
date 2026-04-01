using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.CommandBars;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace CommandTableInfo.Services
{
    public interface ICommandHierarchyService
    {
        Task<CommandHierarchyInfo> GetHierarchyAsync(Command command, CancellationToken cancellationToken);
    }

    public sealed class CommandHierarchyInfo
    {
        public static CommandHierarchyInfo Empty { get; } = new CommandHierarchyInfo("n/a", "n/a", "Not available in current UI context", string.Empty);

        public CommandHierarchyInfo(string displayName, string buttonText, string hierarchyText, string hierarchyCopyText)
        {
            DisplayName = displayName;
            ButtonText = buttonText;
            HierarchyText = hierarchyText;
            HierarchyCopyText = hierarchyCopyText;
        }

        public string DisplayName { get; }

        public string ButtonText { get; }

        public string HierarchyText { get; }

        public string HierarchyCopyText { get; }
    }

    internal sealed class CommandHierarchyService : ICommandHierarchyService
    {
        private readonly DTE2 _dte;
        private readonly Commands2 _commands;
        private readonly IVsProfferCommands3 _profferCommands;
        private readonly Dictionary<CommandKey, CommandHierarchyInfo> _cache = new Dictionary<CommandKey, CommandHierarchyInfo>();
        private readonly Dictionary<CommandKey, HierarchyAccumulator> _hierarchyLookup = new Dictionary<CommandKey, HierarchyAccumulator>();
        private bool _hierarchyIndexBuilt;

        public CommandHierarchyService(DTE2 dte, IVsProfferCommands3 profferCommands)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            _dte = dte;
            _commands = (Commands2)dte.Commands;
            _profferCommands = profferCommands;
        }

        public async Task<CommandHierarchyInfo> GetHierarchyAsync(Command command, CancellationToken cancellationToken)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
            return GetHierarchyCore(command);
        }

        private CommandHierarchyInfo GetHierarchyCore(Command command)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (command == null)
            {
                return CommandHierarchyInfo.Empty;
            }

            var key = CreateCommandKey(command);

            if (_cache.TryGetValue(key, out CommandHierarchyInfo cached))
            {
                return cached;
            }

            var paths = new List<string>();
            var buttonTexts = new List<string>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var seenButtonTexts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            EnsureHierarchyIndexBuilt();

            if (_hierarchyLookup.TryGetValue(key, out HierarchyAccumulator accumulator))
            {
                foreach (string path in accumulator.Paths)
                {
                    if (seen.Add(path))
                    {
                        paths.Add(path);
                    }
                }

                foreach (string cachedButtonText in accumulator.ButtonTexts)
                {
                    if (seenButtonTexts.Add(cachedButtonText))
                    {
                        buttonTexts.Add(cachedButtonText);
                    }
                }
            }

            Guid.TryParse(command.Guid, out Guid commandSet);
            TryAddFindCommandBarPath(command, commandSet, paths, seen);

            string displayName = GetDisplayName(command);

            if (buttonTexts.Count == 0)
            {
                buttonTexts.Add(displayName);
            }

            string buttonText = string.Join(Environment.NewLine, buttonTexts);

            CommandHierarchyInfo info;

            if (paths.Count == 0)
            {
                info = new CommandHierarchyInfo(displayName, buttonText, CommandHierarchyInfo.Empty.HierarchyText, string.Empty);
            }
            else
            {
                string joined = string.Join(Environment.NewLine, paths);
                info = new CommandHierarchyInfo(displayName, buttonText, joined, joined);
            }

            _cache[key] = info;
            return info;
        }

        private void EnsureHierarchyIndexBuilt()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (_hierarchyIndexBuilt)
            {
                return;
            }

            var commandBars = _dte.CommandBars as CommandBars;

            if (commandBars == null)
            {
                _hierarchyIndexBuilt = true;
                return;
            }

            foreach (CommandBar commandBar in commandBars)
            {
                var root = new List<HierarchyNode>
                {
                    new HierarchyNode(commandBar.Name, null, null)
                };

                TraverseControlsForIndex(commandBar.Controls, root);
            }

            _hierarchyIndexBuilt = true;
        }

        private void TraverseControlsForIndex(CommandBarControls controls, IList<HierarchyNode> path)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (controls == null)
            {
                return;
            }

            for (int index = 1; index <= controls.Count; index++)
            {
                CommandBarControl control;

                try
                {
                    control = controls[index];
                }
                catch
                {
                    continue;
                }

                HierarchyNode node = CreateNode(control);
                var currentPath = new List<HierarchyNode>(path) { node };

                if (node.Guid.HasValue && node.Id.HasValue)
                {
                    AddIndexedPath(node, currentPath);
                }

                if (control is CommandBarPopup popup)
                {
                    TraverseControlsForIndex(popup.CommandBar?.Controls, currentPath);
                }
            }
        }

        private void AddIndexedPath(HierarchyNode node, IEnumerable<HierarchyNode> path)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var key = new CommandKey(node.Guid.Value.ToString("D"), node.Id.Value);

            if (!_hierarchyLookup.TryGetValue(key, out HierarchyAccumulator accumulator))
            {
                accumulator = new HierarchyAccumulator();
                _hierarchyLookup[key] = accumulator;
            }

            accumulator.AddPath(string.Join(" > ", path.Select(FormatNode)));

            if (!string.IsNullOrWhiteSpace(node.Label))
            {
                accumulator.AddButtonText(node.Label);
            }
        }

        private static string GetDisplayName(Command command)
        {
            try
            {
                object localizedName = command.GetType().InvokeMember("LocalizedName", System.Reflection.BindingFlags.GetProperty, null, command, null, CultureInfo.InvariantCulture);
                string value = localizedName as string;

                if (!string.IsNullOrWhiteSpace(value))
                {
                    return value;
                }
            }
            catch
            {
            }

            return command.Name ?? "n/a";
        }

        private static void AddPath(IList<string> paths, ISet<string> seen, IEnumerable<HierarchyNode> nodes)
        {
            string path = string.Join(" > ", nodes.Select(FormatNode));

            if (seen.Add(path))
            {
                paths.Add(path);
            }
        }

        private void TryAddFindCommandBarPath(Command command, Guid commandSet, IList<string> paths, ISet<string> seen)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (_profferCommands == null || commandSet == Guid.Empty)
            {
                return;
            }

            object commandBarObject;
            int hr = _profferCommands.FindCommandBar(null, ref commandSet, (uint)command.ID, out commandBarObject);

            if (ErrorHandler.Succeeded(hr) && commandBarObject is CommandBar commandBar)
            {
                string path = string.Format(CultureInfo.InvariantCulture, "FindCommandBar > {0} ({1}:{2})", commandBar.Name, command.Guid, FormatId(command.ID));

                if (seen.Add(path))
                {
                    paths.Add(path);
                }
            }
        }

        private HierarchyNode CreateNode(CommandBarControl control)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            string caption = (control.Caption ?? string.Empty).Trim();

            if (string.IsNullOrWhiteSpace(caption))
            {
                caption = control.Type.ToString();
            }

            try
            {
                _commands.CommandInfo(control, out string guid, out int id);

                if (Guid.TryParse(guid, out Guid parsedGuid))
                {
                    return new HierarchyNode(caption, parsedGuid, id);
                }
            }
            catch
            {
            }

            return new HierarchyNode(caption, null, null);
        }

        private static string FormatNode(HierarchyNode node)
        {
            if (!node.Guid.HasValue || !node.Id.HasValue)
            {
                return node.Label;
            }

            return string.Format(CultureInfo.InvariantCulture, "{0} ({1}:{2})", node.Label, node.Guid.Value, FormatId(node.Id.Value));
        }

        private static string FormatId(int id)
        {
            return string.Format(CultureInfo.InvariantCulture, "0x{0} ({1})", id.ToString("x", CultureInfo.InvariantCulture), id.ToString(CultureInfo.InvariantCulture));
        }

        private static CommandKey CreateCommandKey(Command command)
        {
            if (command == null)
            {
                return new CommandKey(string.Empty, 0);
            }

            string guidText = command.Guid ?? string.Empty;

            if (Guid.TryParse(guidText, out Guid guid))
            {
                guidText = guid.ToString("D");
            }

            return new CommandKey(guidText, command.ID);
        }

        private struct CommandKey : IEquatable<CommandKey>
        {
            public CommandKey(string guid, int id)
            {
                Guid = guid;
                Id = id;
            }

            public string Guid { get; }

            public int Id { get; }

            public bool Equals(CommandKey other)
            {
                return string.Equals(Guid, other.Guid, StringComparison.OrdinalIgnoreCase) && Id == other.Id;
            }

            public override bool Equals(object obj)
            {
                return obj is CommandKey other && Equals(other);
            }

            public override int GetHashCode()
            {
                unchecked
                {
                    return ((Guid != null ? StringComparer.OrdinalIgnoreCase.GetHashCode(Guid) : 0) * 397) ^ Id;
                }
            }
        }

        private sealed class HierarchyNode
        {
            public HierarchyNode(string label, Guid? guid, int? id)
            {
                Label = label;
                Guid = guid;
                Id = id;
            }

            public string Label { get; }

            public Guid? Guid { get; }

            public int? Id { get; }
        }

        private sealed class HierarchyAccumulator
        {
            private readonly HashSet<string> _seenPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            private readonly HashSet<string> _seenButtonTexts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            public HierarchyAccumulator()
            {
                Paths = new List<string>();
                ButtonTexts = new List<string>();
            }

            public IList<string> Paths { get; }

            public IList<string> ButtonTexts { get; }

            public void AddPath(string path)
            {
                if (string.IsNullOrWhiteSpace(path))
                {
                    return;
                }

                if (_seenPaths.Add(path))
                {
                    Paths.Add(path);
                }
            }

            public void AddButtonText(string text)
            {
                if (string.IsNullOrWhiteSpace(text))
                {
                    return;
                }

                if (_seenButtonTexts.Add(text))
                {
                    ButtonTexts.Add(text);
                }
            }
        }
    }
}
