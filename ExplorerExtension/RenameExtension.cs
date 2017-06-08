using SharpShell.Attributes;
using SharpShell.SharpContextMenu;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static ExplorerExtension.Properties.Resources;

namespace ExplorerExtension
{
    [ComVisible(true)]
    [COMServerAssociation(AssociationType.AllFiles)]
    [COMServerAssociation(AssociationType.Directory)]
    public class RenameExtension : SharpContextMenu
    {
        protected override bool CanShowMenu()
        {
            return true;
        }

        protected override ContextMenuStrip CreateMenu()
        {
            var menu = new ContextMenuStrip();
            var parent = new ToolStripMenuItem() { Text = "Archive", Image = main_menu_icon };

            AddAction(parent, "Prefix [yyyyMMdd]", (sender, args) => PrefixDatetime("[yyyyMMdd]"));

            menu.Items.Add(parent);

            return menu;
        }

        private void PrefixDatetime(string format)
        {
            var count = 0;
            foreach (var filePath in SelectedItemPaths)
            {
                var fi = new FileInfo(filePath);
                if (fi.Exists)
                {
                    var date = fi.LastWriteTime;
                    if (fi.Extension == ".msg") date = GetMsgDateTime(fi) ?? date;

                    fi.MoveTo(Path.Combine(fi.DirectoryName, $"{date.ToString(format)}{fi.Name}"));
                    count++;
                    continue;
                }

                var di = new DirectoryInfo(filePath);
                if (di.Exists)
                {
                    di.MoveTo(Path.Combine(di.Parent.FullName, $"{di.LastWriteTime.ToString(format)}{di.Name}"));
                    count++;
                    continue;
                }
            }

            if (count == 0)
            {
                MessageBox.Show($"No file renamed. The files intended to be renamed are:\r\n{string.Join("\r\n", SelectedItemPaths)}");
            }
        }

        private DateTime? GetMsgDateTime(FileInfo fi)
        {
            using (var msg = new MsgReader.Outlook.Storage.Message(fi.FullName))
            {
                var sentOn = msg.SentOn;
                return sentOn;
            }
        }

        private static void AddAction(ToolStripMenuItem parent, string text, EventHandler action, System.Drawing.Image image = null)
        {
            var itemCountLines = new ToolStripMenuItem { Text = text, Image = image };
            itemCountLines.Click += action;
            parent.DropDownItems.Add(itemCountLines);
        }
    }
}
