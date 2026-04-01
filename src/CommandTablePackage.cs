using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using CommandTableInfo.Commands;
using CommandTableInfo.Services;
using CommandTableInfo.ToolWindows;
using EnvDTE;
using EnvDTE80;
using Microsoft;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace CommandTableInfo
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration(Vsix.Name, Vsix.Description, Vsix.Version)]
    [Guid(PackageGuids.guidCommandTablePackageString)]
    [ProvideToolWindow(typeof(CommandTableWindow), Style = VsDockStyle.Tabbed, Window = "DocumentWell", Orientation = ToolWindowOrientation.none)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    public sealed class CommandTablePackage : AsyncPackage
    {
        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
            await ShowToolWindow.InitializeAsync(this);
            await ContextMenuCommands.InitializeAsync(this);
        }

        public override IVsAsyncToolWindowFactory GetAsyncToolWindowFactory(Guid toolWindowType)
        {
            if (toolWindowType.Equals(new Guid(CommandTableWindow.WindowGuidString)))
            {
                return this;
            }

            return null;
        }

        protected override string GetToolWindowTitle(Type toolWindowType, int id)
        {
            if (toolWindowType == typeof(CommandTableWindow))
            {
                return CommandTableWindow.Title;
            }

            return base.GetToolWindowTitle(toolWindowType, id);
        }

        protected override async Task<object> InitializeToolWindowAsync(Type toolWindowType, int id, CancellationToken cancellationToken)
        {
            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            var dte = await GetServiceAsync(typeof(DTE)) as DTE2;
            var profferCommands = await GetServiceAsync(typeof(SVsProfferCommands)) as IVsProfferCommands3;
            var uiShell = await GetServiceAsync(typeof(SVsUIShell)) as IVsUIShell;
            var vsShell = await GetServiceAsync(typeof(SVsShell)) as IVsShell;
            Assumes.Present(dte);

            var dto = new CommandTableExplorerDTO();
            var dteCommands = new List<Command>();

            foreach (Command command in dte.Commands)
            {
                if (command != null && !string.IsNullOrEmpty(command.Name))
                {
                    dteCommands.Add(command);
                }
            }

            dto.DTE = dte;
            dto.VsUiShell = uiShell;
            dto.CommandHierarchyService = new CommandHierarchyService(dte, profferCommands, vsShell);
            dto.DteCommands = dteCommands
                .OrderBy(c => c.Name, StringComparer.OrdinalIgnoreCase)
                .ToList();

            return dto;
        }
    }
}
