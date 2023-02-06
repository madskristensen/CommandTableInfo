//using System;
//using System.Linq;
//using System.Windows;
//using System.Windows.Controls;

//namespace CommandTableInfo.ToolWindows
//{
//    public class CommandTreeItem : TreeViewItem
//    {
//        public CommandTreeItem(CommandGroup item)
//        {
//            Command = item;
//            Header = item.SymbolicItemId.SymbolicDWordName;

//            if (item.Placements.Any())
//            {
//                var placements = new TreeViewItem { Header = "Placements" };
//                placements.ItemsSource = item.Placements.Select(p => new CommandTreeItem(p));

//                Items.Add(placements);
//            }

//            if (item.Items.Any())
//            {
//                var items = new TreeViewItem { Header = "Commands" };
//                items.ItemsSource = item.Items.Select(p => new CommandTreeItem(p));

//                Items.Add(items);
//            }
//        }

//        public CommandTreeItem(CommandContainer item)
//        {
//            Command = item;
//            Header = item.SymbolicItemId.SymbolicDWordName;

//            if (item.Placements.Any())
//            {
//                var placements = new TreeViewItem { Header = "Placements" };
//                placements.ItemsSource = item.Placements.Select(p => new CommandTreeItem(p));

//                Items.Add(placements);
//            }
//        }

//        public CommandTreeItem(CommandItem item)
//        {
//            Command = item;
//            Header = item.SymbolicItemId.SymbolicDWordName;
//        }

//        public CommandItem Command { get; set; }

//        public static event EventHandler<CommandTreeItem> ItemSelected;

//        protected override void OnSelected(RoutedEventArgs e)
//        {
//            if (IsSelected)
//            {
//                //SetResourceReference(TextElement.ForegroundProperty, EnvironmentColors.SystemHighlightTextBrushKey);
//                ItemSelected?.Invoke(this, this);
//            }
//        }
//    }
//}
