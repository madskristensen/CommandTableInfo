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

namespace CommandTableInfo.Services
{
    public interface ICommandHierarchyService
    {
        CommandHierarchyInfo GetHierarchy(Command command);
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

        public CommandHierarchyService(DTE2 dte, IVsProfferCommands3 profferCommands)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            _dte = dte;
            _commands = (Commands2)dte.Commands;
            _profferCommands = profferCommands;
        }

        public CommandHierarchyInfo GetHierarchy(Command command)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (command == null)
            {
                return CommandHierarchyInfo.Empty;
            }

            var key = new CommandKey(command.Guid ?? string.Empty, command.ID);

            if (_cache.TryGetValue(key, out CommandHierarchyInfo cached))
            {
                return cached;
            }

            var paths = new List<string>();
            var buttonTexts = new List<string>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var seenButtonTexts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            Guid.TryParse(command.Guid, out Guid commandSet);
            EnumerateCommandBars(command, commandSet, paths, seen, buttonTexts, seenButtonTexts);
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

        private void EnumerateCommandBars(Command command, Guid commandSet, IList<string> paths, ISet<string> seen, IList<string> buttonTexts, ISet<string> seenButtonTexts)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var commandBars = _dte.CommandBars as CommandBars;

            if (commandBars == null)
            {
                return;
            }

            foreach (CommandBar commandBar in commandBars)
            {
                var root = new List<HierarchyNode>
                {
                    new HierarchyNode(commandBar.Name, null, null)
                };

                TraverseControls(commandBar.Controls, root, commandSet, command.ID, paths, seen, buttonTexts, seenButtonTexts);
            }
        }

        private void TraverseControls(CommandBarControls controls, IList<HierarchyNode> path, Guid commandSet, int commandId, IList<string> paths, ISet<string> seen, IList<string> buttonTexts, ISet<string> seenButtonTexts)
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

                if (node.Guid.HasValue &&
                    node.Guid.Value == commandSet &&
                    node.Id.HasValue &&
                    node.Id.Value == commandId)
                {
                    AddPath(paths, seen, currentPath);

                    if (!string.IsNullOrWhiteSpace(node.Label) && seenButtonTexts.Add(node.Label))
                    {
                        buttonTexts.Add(node.Label);
                    }
                }

                if (control is CommandBarPopup popup)
                {
                    TraverseControls(popup.CommandBar?.Controls, currentPath, commandSet, commandId, paths, seen, buttonTexts, seenButtonTexts);
                }
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
    }
}
