using System.Collections.Generic;
using EnvDTE80;
using CommandTableInfo.Services;
using Microsoft.VisualStudio.Shell.Interop;

namespace CommandTableInfo.ToolWindows
{
    internal class CommandTableExplorerDTO
    {
        public IList<EnvDTE.Command> DteCommands { get; set; }
        public DTE2 DTE { get; set; }
        internal IVsUIShell VsUiShell { get; set; }
        internal ICommandHierarchyService CommandHierarchyService { get; set; }
    }
}
