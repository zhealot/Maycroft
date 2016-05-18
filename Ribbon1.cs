using System;
using System.Linq;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using System.IO;
using Word = Microsoft.Office.Interop.Word;

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
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex.Message);
                MessageBox.Show("Failed to copy HTML code " + Environment.NewLine + ex.Message);
            }
        }

        private void btnCopyReHTML_Click(object sender, RibbonControlEventArgs e)
        {
            //open template document
            Word.Application appWord = new Word.Application();
            appWord.Visible = false;
            try
            {
                object oFilePath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase) + "\\Template.docx";
                Word.Document docTemplate = appWord.Documents.Open(ref oFilePath, ReadOnly: true, Visible: false);
                if (docTemplate.SelectContentControlsByTitle("name").Count > 0)
                {
                    string sSigName = docTemplate.SelectContentControlsByTitle("name")[1].Range.Text.Trim();
                    docTemplate.Close(SaveChanges: false);
                    appWord.Quit(SaveChanges: false);
                    string sAppDataDir = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\microsoft\\signatures";
                    string sSigFile = sAppDataDir + "\\" + sSigName + "_reply.htm";
                    if (File.Exists(sSigFile))
                    {
                        StreamReader sr = new StreamReader(sSigFile, System.Text.Encoding.Default);
                        Clipboard.SetText(sr.ReadToEnd());
                        MessageBox.Show("Reply signature HTML code copied to clipboard");
                    }
                    else
                    {
                        MessageBox.Show("Signature file not found");
                    }
                }
                else
                {
                    MessageBox.Show("Signature template not found");
                    docTemplate.Close(SaveChanges: false);
                    appWord.Quit(SaveChanges: false);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
    }
}
