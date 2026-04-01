using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.CommandBars;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Threading;
using Microsoft.Win32;

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

        /// <summary>
        /// Eagerly pre-builds the hierarchy index and owner package map in the background
        /// so that the first <see cref="GetHierarchyAsync"/> call does not block.
        /// </summary>
        Task PreBuildAsync(CancellationToken cancellationToken);
    }

    public sealed class CommandHierarchyInfo
    {
        public static CommandHierarchyInfo Empty { get; } = new CommandHierarchyInfo("n/a", "n/a", "Not available in current UI context", string.Empty, null);

        public CommandHierarchyInfo(string displayName, string buttonText, string hierarchyText, string hierarchyCopyText, string ownerPackageName)
        {
            DisplayName = displayName;
            ButtonText = buttonText;
            HierarchyText = hierarchyText;
            HierarchyCopyText = hierarchyCopyText;
            OwnerPackageName = ownerPackageName;
        }

        public string DisplayName { get; }

        public string ButtonText { get; }

        public string HierarchyText { get; }

        public string HierarchyCopyText { get; }

        public string OwnerPackageName { get; }
    }

    internal sealed class CommandHierarchyService : ICommandHierarchyService
    {
        private static readonly Dictionary<Guid, string> WellKnownCommandSetOwners = CreateWellKnownCommandSetOwnerMap();

        private readonly DTE2 _dte;
        private readonly Commands2 _commands;
        private readonly IVsProfferCommands3 _profferCommands;
        private readonly IVsShell _vsShell;
        private readonly Dictionary<CommandKey, CommandHierarchyInfo> _cache = new Dictionary<CommandKey, CommandHierarchyInfo>();
        private readonly Dictionary<CommandKey, HierarchyAccumulator> _hierarchyLookup = new Dictionary<CommandKey, HierarchyAccumulator>();
        private Dictionary<Guid, string> _commandSetOwnerMap;
        private bool _hierarchyIndexBuilt;

        public CommandHierarchyService(DTE2 dte, IVsProfferCommands3 profferCommands, IVsShell vsShell)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            _dte = dte;
            _commands = (Commands2)dte.Commands;
            _profferCommands = profferCommands;
            _vsShell = vsShell;
        }

        public async Task PreBuildAsync(CancellationToken cancellationToken)
        {
            // Step 1: Traverse CommandBars on the UI thread (COM requirement)
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
            EnsureHierarchyIndexBuilt();

            // Capture the registry root path on the UI thread (VSRegistry uses COM internally)
            string registryRootPath = GetRegistryRootPath();

            // Step 2: Build the owner map on a background thread (registry I/O is thread-safe)
            await TaskScheduler.Default;
            cancellationToken.ThrowIfCancellationRequested();
            BuildCommandSetOwnerMap(registryRootPath);
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

            string ownerPackageName = ResolveOwnerPackageName(commandSet);

            CommandHierarchyInfo info;

            if (paths.Count == 0)
            {
                info = new CommandHierarchyInfo(displayName, buttonText, CommandHierarchyInfo.Empty.HierarchyText, string.Empty, ownerPackageName);
            }
            else
            {
                string joined = string.Join(Environment.NewLine, paths);
                info = new CommandHierarchyInfo(displayName, buttonText, joined, joined, ownerPackageName);
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
                var path = new List<HierarchyNode>
                {
                    new HierarchyNode(commandBar.Name, null, null)
                };

                TraverseControlsForIndex(commandBar.Controls, path);
                path.RemoveAt(path.Count - 1);
            }

            _hierarchyIndexBuilt = true;
        }

        private void TraverseControlsForIndex(CommandBarControls controls, List<HierarchyNode> path)
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
                path.Add(node);

                if (node.Guid.HasValue && node.Id.HasValue)
                {
                    AddIndexedPath(node, path);
                }

                if (control is CommandBarPopup popup)
                {
                    TraverseControlsForIndex(popup.CommandBar?.Controls, path);
                }

                path.RemoveAt(path.Count - 1);
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

        private string ResolveOwnerPackageName(Guid commandSet)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (commandSet == Guid.Empty)
            {
                return null;
            }

            EnsureCommandSetOwnerMap();

            if (_commandSetOwnerMap.TryGetValue(commandSet, out string ownerName))
            {
                return ownerName;
            }

            if (WellKnownCommandSetOwners.TryGetValue(commandSet, out string wellKnownOwner))
            {
                return wellKnownOwner;
            }

            return null;
        }

        private static string GetRegistryRootPath()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                using (RegistryKey rootKey = VSRegistry.RegistryRoot(__VsLocalRegistryType.RegType_Configuration, writable: false))
                {
                    return rootKey?.Name;
                }
            }
            catch
            {
                return null;
            }
        }

        private void BuildCommandSetOwnerMap(string registryRootPath)
        {
            if (_commandSetOwnerMap != null)
            {
                return;
            }

            var map = new Dictionary<Guid, string>();

            if (!string.IsNullOrEmpty(registryRootPath))
            {
                try
                {
                    BuildCommandSetOwnerMapFromRegistry(registryRootPath, map);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.Write(ex);
                }
            }

            _commandSetOwnerMap = map;
        }

        private static void BuildCommandSetOwnerMapFromRegistry(string registryRootPath, Dictionary<Guid, string> ownerMap)
        {
            // Parse hive and subkey from the full registry path
            int firstBackslash = registryRootPath.IndexOf('\\');

            if (firstBackslash < 0)
            {
                return;
            }

            string hiveName = registryRootPath.Substring(0, firstBackslash);
            string subKeyPath = registryRootPath.Substring(firstBackslash + 1);
            RegistryKey hive;

            switch (hiveName)
            {
                case "HKEY_CURRENT_USER":
                    hive = Registry.CurrentUser;
                    break;
                case "HKEY_LOCAL_MACHINE":
                    hive = Registry.LocalMachine;
                    break;
                default:
                    hive = Registry.CurrentUser;
                    break;
            }

            using (RegistryKey rootKey = hive.OpenSubKey(subKeyPath, writable: false))
            {
                if (rootKey == null)
                {
                    return;
                }

                var packageNames = new Dictionary<Guid, string>();

                using (RegistryKey packagesKey = rootKey.OpenSubKey("Packages"))
                {
                    if (packagesKey != null)
                    {
                        foreach (string pkgGuidStr in packagesKey.GetSubKeyNames())
                        {
                            if (!Guid.TryParse(pkgGuidStr, out Guid pkgGuid))
                            {
                                continue;
                            }

                            using (RegistryKey pkgKey = packagesKey.OpenSubKey(pkgGuidStr))
                            {
                                if (pkgKey == null)
                                {
                                    continue;
                                }

                                string name = pkgKey.GetValue("ProductName") as string
                                           ?? pkgKey.GetValue("") as string;

                                if (!string.IsNullOrWhiteSpace(name) && !name.StartsWith("#", StringComparison.Ordinal))
                                {
                                    packageNames[pkgGuid] = name;
                                }
                                else
                                {
                                    string className = pkgKey.GetValue("Class") as string;

                                    if (!string.IsNullOrWhiteSpace(className))
                                    {
                                        int lastDot = className.LastIndexOf('.');
                                        string shortName = lastDot >= 0 && lastDot < className.Length - 1
                                            ? className.Substring(lastDot + 1)
                                            : className;
                                        packageNames[pkgGuid] = shortName;
                                    }
                                }
                            }
                        }
                    }
                }

                using (RegistryKey menusKey = rootKey.OpenSubKey("Menus"))
                {
                    if (menusKey != null)
                    {
                        foreach (string valueName in menusKey.GetValueNames())
                        {
                            if (!Guid.TryParse(valueName, out Guid pkgGuid))
                            {
                                continue;
                            }

                            if (!packageNames.TryGetValue(pkgGuid, out string pkgName))
                            {
                                continue;
                            }

                            string menuData = menusKey.GetValue(valueName) as string;

                            if (string.IsNullOrWhiteSpace(menuData))
                            {
                                continue;
                            }

                            // Menus value format: ",resId,version" or "satellite,resId,version"
                            // The package GUID itself is the value name; associate it as a potential command set owner
                            if (!ownerMap.ContainsKey(pkgGuid))
                            {
                                ownerMap[pkgGuid] = pkgName;
                            }
                        }
                    }
                }

                // Also scan for explicit command set GUID registrations under Packages\{pkg}\CmdSets or similar
                using (RegistryKey packagesKey = rootKey.OpenSubKey("Packages"))
                {
                    if (packagesKey != null)
                    {
                        foreach (string pkgGuidStr in packagesKey.GetSubKeyNames())
                        {
                            if (!Guid.TryParse(pkgGuidStr, out Guid pkgGuid) || !packageNames.ContainsKey(pkgGuid))
                            {
                                continue;
                            }

                            // Many packages register their command set GUID equal to or near their package GUID.
                            // We already map pkgGuid above via Menus. Now also check for explicit entries.
                            string pkgName = packageNames[pkgGuid];

                            if (!ownerMap.ContainsKey(pkgGuid))
                            {
                                ownerMap[pkgGuid] = pkgName;
                            }
                        }
                    }
                }
            }
        }

        private void EnsureCommandSetOwnerMap()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (_commandSetOwnerMap != null)
            {
                return;
            }

            string registryRootPath = GetRegistryRootPath();
            BuildCommandSetOwnerMap(registryRootPath);
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

        private static Dictionary<Guid, string> CreateWellKnownCommandSetOwnerMap()
        {
            return new Dictionary<Guid, string>
            {
                // Standard command sets (shell)
                { VSConstants.GUID_VSStandardCommandSet97, "Visual Studio (Standard97)" },
                { VSConstants.VSStd2K, "Visual Studio (Standard2K)" },
                { VSConstants.VsStd2010, "Visual Studio (Standard2010)" },
                { VSConstants.VsStd11, "Visual Studio (Standard11)" },
                { VSConstants.VsStd12, "Visual Studio (Standard12)" },
                { VSConstants.VsStd14, "Visual Studio (Standard14)" },
                { VSConstants.VsStd15, "Visual Studio (Standard15)" },
                // Shell main menu (guidSHLMainMenu - menus, groups, toolbars)
                { VSConstants.CMDSETID.ShellMainMenu_guid, "Visual Studio (Shell Main Menu)" },
                // UI hierarchy commands
                { VSConstants.GUID_VsUIHierarchyWindowCmds, "Visual Studio (UI Hierarchy)" },
                // Environment package
                { VSConstants.CLSID_VsEnvironmentPackage, "Visual Studio (Environment)" },
                // Text editor
                { VSConstants.GUID_TextEditorFactory, "Visual Studio (Text Editor)" },
                // HTML editor
                { VSConstants.CLSID_HtmedPackage, "Visual Studio (HTML Editor)" },
                // Document outline
                { VSConstants.CLSID_VsDocOutlinePackage, "Visual Studio (Document Outline)" },
                // Task list
                { VSConstants.CLSID_VsTaskListPackage, "Visual Studio (Task List)" },
                // App command routing
                { VSConstants.GUID_AppCommand, "Visual Studio (App Command)" },
                // Solution Explorer pivot list
                { VSConstants.CMDSETID.SolutionExplorerPivotList_guid, "Visual Studio (Solution Explorer)" },
                // C# language commands
                { VSConstants.CMDSETID.CSharpGroup_guid, "Visual Studio (C#)" },
                // Debugger command sets (from VsDebugGuids.h)
                { new Guid("c9dd4a58-47fb-11d2-83e7-00c04f9902c1"), "Visual Studio (Debugger)" },       // guidVSDebugGroup
                { new Guid("c9dd4a59-47fb-11d2-83e7-00c04f9902c1"), "Visual Studio (Debugger)" },       // guidVSDebugCommand
                // Shared commands (from sharedids.h)
                { new Guid("8328592b-227c-11d3-b870-00c04f79f802"), "Visual Studio (Shared Commands)" }, // guidSharedCmd
                // Common IDE package commands (from vsshlids.h)
                { new Guid("6767e06b-5789-472b-8ed7-1f2073716e8c"), "Visual Studio (Common IDE)" },      // guidCommonIDEPackageCmd
                // Shared menu group (from vsshlids.h)
                { new Guid("234a7fc1-cfe9-4335-9e82-061f86e402c1"), "Visual Studio (Shared Menu)" },     // guidSharedMenuGroup
                // Class View commands (from vsshlids.h)
                { new Guid("fb61dcfe-c9cb-4964-8426-c2d38334078c"), "Visual Studio (Class View)" },      // guidClassViewMenu
                // Data commands (from vsshlids.h)
                { new Guid("501822e1-b5af-11d0-b4dc-00a0c91506ef"), "Visual Studio (Data)" },            // guidDataCmdId
                { new Guid("4614107f-217d-4bbf-9dfe-b9e165c65572"), "Visual Studio (Data)" },            // guidVSData
                // Server Explorer commands (from vsshlids.h)
                { new Guid("74d21310-2aee-11d1-8bfb-00a0c90f26f7"), "Visual Studio (Server Explorer)" }, // guid_SE_MenuGroup
                { new Guid("74d21311-2aee-11d1-8bfb-00a0c90f26f7"), "Visual Studio (Server Explorer)" }, // guid_SE_CommandID
                // SQL Object Explorer (from vsshlids.h)
                { new Guid("03f46784-2f90-4122-91ec-72ff9e11d9a3"), "Visual Studio (SQL)" },             // guidSqlObjectExplorerCmdSet
                // Web Browser commands (from wbids.h)
                { new Guid("e8b06f44-6d01-11d2-aa7d-00c04f990343"), "Visual Studio (Web Browser)" },     // guidWBPkgCmd
                { new Guid("e8b06f42-6d01-11d2-aa7d-00c04f990343"), "Visual Studio (Web Browser)" },     // guidWBGrp
                // Editor keybinding emulation command sets (from vsshlids.h)
                { new Guid("9a95f3af-f86a-4aa2-80e6-012bf65dbbc3"), "Visual Studio (Emacs Emulation)" }, // guidEmacsCommandGroup
                { new Guid("7a500d8a-8258-46c3-8965-6ac53ed6b4e7"), "Visual Studio (Brief Emulation)" }, // guidBriefCommandGroup
                // Extension Manager (from vsshlids.h)
                { new Guid("e7576c05-1874-450c-9e98-cf3a0897a069"), "Visual Studio (Extension Manager)" }, // guidExtensionManagerPkg
                // Debug target handlers
                { VSConstants.DebugTargetHandler.guidDebugTargetHandlerCmdSet, "Visual Studio (Debug Target)" },
                { VSConstants.AppPackageDebugTargets.guidAppPackageDebugTargetCmdSet, "Visual Studio (App Package Debug)" },
                // Razor (from RazorGuids.h)
                { new Guid("5289d302-2432-4761-8c45-051c64bd00c4"), "Visual Studio (Razor)" },           // guidRazorCmdSet
                // Team Explorer (from vsshlids.h)
                { new Guid("3f5a3e02-af62-4c13-8d8a-a568ecae238b"), "Visual Studio (Team Explorer)" },   // guidTeamExplorerSharedCmdSet
                // Designer package (from vsshlids.h)
                { new Guid("8d8529d3-625d-4496-8354-3dad630ecc1b"), "Visual Studio (Designer)" },        // guid_VSDesignerPackage
                // VDT Flavor commands (from vsshlids.h)
                { new Guid("462b036f-7349-4835-9e21-bec60e989b9c"), "Visual Studio (VDT Flavor)" },      // guidVDTFlavorCmdSet
                // Reference Manager (from vsshlids.h)
                { new Guid("7b069159-ff02-4752-93e8-96b3cadf441a"), "Visual Studio (Reference Manager)" }, // guidReferenceManagerProvidersPackageCmdSet
                // Universal Projects (from ShellCmdDef.vsct)
                { new Guid("04b4dc54-4183-44e7-b353-61424e7d2dab"), "Visual Studio (Universal Projects)" }, // guidUniversalProjectsCmdSet
                // Project retargeting (from ShellCmdDef.vsct)
                { new Guid("284b5e45-1ff9-40b1-9ccf-92319a47c39b"), "Visual Studio (Project Retargeting)" }, // guidTrackProjectRetargetingCmdSet
            };
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
