using System.Collections.Generic;
using CommandTable;
using EnvDTE80;

namespace CommandTableInfo.ToolWindows
{
    public class CommandTableExplorerDTO
    {
        public IList<EnvDTE.Command> DteCommands { get; set; }
        public ICommandTable CommandTable { get; set; }
        public DTE2 DTE { get; set; }
    }
}
