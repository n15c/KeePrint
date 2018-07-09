using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using KeePass.Plugins;
using System.Windows.Forms;
using KeePassLib;

namespace KeePrint
{
    public sealed class KeePrintExt : Plugin
    {
        private IPluginHost m_host = null;

        public override bool Initialize(IPluginHost host)
        {
            m_host = host;
            //EntryContextMenu tsMenu =
            ToolStripItemCollection menuItemCollection = m_host.MainWindow.EntryContextMenu.Items;
            //System.Drawing.Image image = new System.Drawing.Bitmap("tmp");

            ToolStripMenuItem cmenuItem = new ToolStripMenuItem();
            cmenuItem.Text = "Zugangsdaten drucken";
            cmenuItem.Click += CmenuItem_Click;
            menuItemCollection.Add(cmenuItem);

            return true;
        }

        private void CmenuItem_Click(object sender, EventArgs e)
        {
            string username;
            string password;
            string comment;
            PwEntry[] pwEntry = m_host.MainWindow.GetSelectedEntries();
            PwEntry entry = pwEntry[0];
            password = entry.Strings.ReadSafe("Password");
            username = entry.Strings.ReadSafe("UserName");
            comment = entry.Strings.ReadSafe("Notes");

            PrintPaper print = new PrintPaper(username, password, comment, "template.docx");
            print.PrintDoc();
        }
    }
}
