using System;
using System.Linq;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;

namespace MaycroftOL
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Form1 SetupForm = new Form1();
            SetupForm.ShowDialog();
            //SetupForm.Show();
            if (SetupForm.Controls.Find("tbName", true).Count() > 0)
                SetupForm.Controls.Find("tbName", true)[0].Focus();
        }

        private void btnCopyHTML_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var ol = Globals.ThisAddIn.Application;
                var NewMail = (Microsoft.Office.Interop.Outlook.MailItem)(ol.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem));
                NewMail.Display(NewMail);
                var htmlbody = NewMail.HTMLBody;
                NewMail.Delete();
                htmlbody.Trim();
                Clipboard.SetText(htmlbody.ToString());
                MessageBox.Show("Signature HTML code copied to clipboard.");
            }
            catch(Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex.Message);
                MessageBox.Show("Failed to copy HTML code " + Environment.NewLine + ex.Message);
            }
        }
    }
}
