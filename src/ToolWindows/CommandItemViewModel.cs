using System;
using System.Collections.Generic;
using System.Linq;
using CommandTable;

namespace CommandTableInfo.ToolWindows
{
    public class CommandItemViewModel
    {

        public CommandItemViewModel(CommandGroup group)
        {
            SymbolicId = group.SymbolicItemId.SymbolicDWordName;
            SymbolicName = group.SymbolicItemId.SymbolicGuidName;
            Guid = group.ItemId.Guid;
            ID = group.ItemId.DWord;
            Priority = group.Priority;

            Menu = new CommandItemViewModel(group.Placements.FirstOrDefault());
        }

        public CommandItemViewModel(CommandContainer item)
        {
            if (item == null)
                return;

            SymbolicId = item.SymbolicItemId.SymbolicDWordName;
            SymbolicName = item.SymbolicItemId.SymbolicGuidName;
            Guid = item.ItemId.Guid;
            ID = item.ItemId.DWord;
            Priority = item.Priority;
        }

        public string SymbolicId { get; set; }
        public string SymbolicName { get; set; }
        public Guid Guid { get; set; }
        public uint ID { get; set; }
        public uint Priority { get; set; }
        public CommandItemViewModel Menu { get; set; }

    }
}
