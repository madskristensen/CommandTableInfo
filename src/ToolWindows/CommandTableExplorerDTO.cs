using System.Collections.Generic;
using CommandTable;

namespace CommandTableInfo.ToolWindows
{
    public class CommandTableExplorerDTO
    {
        public IEnumerable<EnvDTE.Command> DteCommands { get; set; }
        public ICommandTable CommandTable { get; set; }
    }
}
