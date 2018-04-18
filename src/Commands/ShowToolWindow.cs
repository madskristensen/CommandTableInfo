using System;
using System.ComponentModel.Design;
using CommandTableInfo.ToolWindows;
using Microsoft.VisualStudio.Shell;
using Task = System.Threading.Tasks.Task;

namespace CommandTableInfo
{
    internal sealed class ShowToolWindow
    {
        public static async Task InitializeAsync(AsyncPackage package)
        {
            var commandService = (IMenuCommandService)await package.GetServiceAsync(typeof(IMenuCommandService));

            var menuCommandID = new CommandID(PackageGuids.guidCommandTablePackageCmdSet, PackageIds.ShowToolWindowId);
            var menuItem = new MenuCommand((sender, e) => Execute(package, sender, e), menuCommandID);
            commandService.AddCommand(menuItem);
        }

        private static void Execute(AsyncPackage package, object sender, EventArgs e)
        {
            package.JoinableTaskFactory.RunAsync(async () =>
            {
                ToolWindowPane window = await package.ShowToolWindowAsync(
                    typeof(CommandTableWindow),
                    0,
                    create: true,
                    cancellationToken: package.DisposalToken);
            });
        }
    }
}
