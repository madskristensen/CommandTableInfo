using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using CommandTable;
using CommandTableInfo.ToolWindows;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace CommandTableInfo
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration("#110", "#112", Vsix.Version, IconResourceID = 400)]
    [Guid(PackageGuids.guidCommandTablePackageString)]
    [ProvideToolWindow(typeof(CommandTableWindow), Style = VsDockStyle.Tabbed, Window = "DocumentWell", Orientation = ToolWindowOrientation.none)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    public sealed class CommandTablePackage : AsyncPackage
    {
        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {   
            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
            await ShowToolWindow.InitializeAsync(this);
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
            var dte = await GetServiceAsync(typeof(DTE)) as DTE2;
            var dto = new CommandTableExplorerDTO();
            var dteCommands = new List<EnvDTE.Command>();

            foreach (EnvDTE.Command command in dte.Commands)
            {
                if (!string.IsNullOrEmpty(command.Name))
                    dteCommands.Add(command);
            }

            dto.DTE = dte;
            dto.DteCommands = dteCommands.OrderBy(c => c.Name).ToList();

            var factory = new CommandTableFactory();
            ICommandTable table = factory.CreateCommandTableFromHost(this, HostLoadType.FromRegisteredMenuDlls);
            dto.CommandTable = await table.GetCommands();

            return dto;
        }
    }
}
