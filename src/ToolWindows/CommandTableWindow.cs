using System;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.Imaging;
using Microsoft.VisualStudio.Shell;

namespace CommandTableInfo.ToolWindows
{
    [Guid(WindowGuidString)]
    public class CommandTableWindow : ToolWindowPane
    {
        public const string WindowGuidString = "7c9c901a-f0d0-45d0-9ebd-7b2748a4b49a";
        public const string Title = "Command Explorer";

        public CommandTableWindow()
            : this(null)
        { }

        public CommandTableWindow(CommandTableExplorerDTO state)
            : base()
        {
            Caption = Title;
            BitmapImageMoniker = KnownMonikers.CommandUIOption;

            var elm = new CommandTableExplorerControl(state);
            Content = elm;
        }
    }
}
