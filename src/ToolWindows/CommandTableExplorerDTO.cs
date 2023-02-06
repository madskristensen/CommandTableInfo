using System.Collections.Generic;
using EnvDTE80;

namespace CommandTableInfo.ToolWindows
{
    public class CommandTableExplorerDTO
    {
        public IList<EnvDTE.Command> DteCommands { get; set; }
        public DTE2 DTE { get; set; }
    }
}
