using CommandTableInfo.ToolWindows;

using Microsoft;
using Microsoft.VisualStudio.Shell;

using System;
using System.ComponentModel.Design;

namespace CommandTableInfo.Commands
{
    internal static class ContextMenuCommands
    {
        public static async System.Threading.Tasks.Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

            var commandService = (IMenuCommandService)await package.GetServiceAsync(typeof(IMenuCommandService));
            Assumes.Present(commandService);

            var copyCommand = new OleMenuCommand(ExecuteCopyContextValue, new CommandID(PackageGuids.guidCommandTablePackageCmdSet, PackageIds.CopyContextValueId));
            copyCommand.BeforeQueryStatus += OnCopyContextValueBeforeQueryStatus;
            commandService.AddCommand(copyCommand);

            var copyVsctSymbolsCommand = new OleMenuCommand(ExecuteCopyVsctSymbols, new CommandID(PackageGuids.guidCommandTablePackageCmdSet, PackageIds.CopyVsctSymbolsId));
            copyVsctSymbolsCommand.BeforeQueryStatus += OnCopyVsctSymbolsBeforeQueryStatus;
            commandService.AddCommand(copyVsctSymbolsCommand);
        }

        private static void ExecuteCopyContextValue(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            CommandTableExplorerControl.ExecuteNativeContextMenuCommand(PackageIds.CopyContextValueId);
        }

        private static void ExecuteCopyVsctSymbols(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            CommandTableExplorerControl.ExecuteNativeContextMenuCommand(PackageIds.CopyVsctSymbolsId);
        }

        private static void OnCopyContextValueBeforeQueryStatus(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (sender is OleMenuCommand command)
            {
                command.Enabled = CommandTableExplorerControl.CanExecuteNativeContextMenuCommand(PackageIds.CopyContextValueId);
                command.Visible = true;
            }
        }

        private static void OnCopyVsctSymbolsBeforeQueryStatus(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (sender is OleMenuCommand command)
            {
                command.Enabled = CommandTableExplorerControl.CanExecuteNativeContextMenuCommand(PackageIds.CopyVsctSymbolsId);
                command.Visible = true;
            }
        }
    }
}
