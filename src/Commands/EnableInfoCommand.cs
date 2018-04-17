using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Linq;
using System.Windows.Input;
using CommandTable;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Task = System.Threading.Tasks.Task;

namespace CommandTableInfo
{
    internal sealed class EnableInfoCommand
    {
        private readonly AsyncPackage _package;
        private CommandEvents _cmdEvents;

        private EnableInfoCommand(AsyncPackage package, OleMenuCommandService commandService, DTE2 dte)
        {
            _package = package ?? throw new ArgumentNullException(nameof(package));
            _cmdEvents = dte.Events.CommandEvents;
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(PackageGuids.guidCommandTablePackageCmdSet, PackageIds.EnableInfoCommandId);
            var menuItem = new OleMenuCommand(Execute, menuCommandID);
            menuItem.BeforeQueryStatus += BeforeQueryStatus;
            commandService.AddCommand(menuItem);
        }

        public static EnableInfoCommand Instance
        {
            get;
            private set;
        }

        public bool IsEnabled
        {
            get;
            private set;
        }

        public ICommandTable CommandTable
        {
            get;
            private set;
        }

        private IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return _package;
            }
        }

        public static async Task InitializeAsync(AsyncPackage package)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var commandService = await package.GetServiceAsync((typeof(IMenuCommandService))) as OleMenuCommandService;
            var dte = await package.GetServiceAsync(typeof(DTE)) as DTE2;

            Instance = new EnableInfoCommand(package, commandService, dte);
        }

        private void BeforeQueryStatus(object sender, EventArgs e)
        {
            var button = (OleMenuCommand)sender;
            button.Checked = IsEnabled;
        }

        private void Execute(object sender, EventArgs e)
        {
            IsEnabled = !IsEnabled;

            if (IsEnabled)
            {
                _cmdEvents.BeforeExecute += CommandEvents_BeforeExecute;

                if (CommandTable == null)
                {
                    var factory = new CommandTableFactory();
                    CommandTable = factory.CreateCommandTableFromHost(_package, HostLoadType.FromRegisteredMenuDlls);
                }
            }
            else
            {
                _cmdEvents.BeforeExecute -= CommandEvents_BeforeExecute;
            }
        }

        private void CommandEvents_BeforeExecute(string Guid, int ID, object CustomIn, object CustomOut, ref bool CancelDefault)
        {
            if (IsEnabled && CommandTable != null)
            {
                if ((Keyboard.IsKeyDown(Key.LeftShift) || Keyboard.IsKeyDown(Key.RightShift)) &&
                    (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl)))
                {
                    CancelDefault = true;
                    ThreadHelper.JoinableTaskFactory.Run(() => ShowInfoAsync(new Guid(Guid), ID));
                }
            }
        }

        private async Task ShowInfoAsync(Guid guid, int id)
        {
            IEnumerable<CommandTable.Command> commands = await CommandTable.GetCommands();
            CommandTable.Command command = commands.FirstOrDefault(c => c.ItemId.Guid == guid && c.ItemId.DWord == id);
            System.Diagnostics.Debug.Write(command);
        }
    }
}
