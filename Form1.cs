using System;
using System.ComponentModel;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;

namespace MaycroftOL
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            var CurUsr = Globals.ThisAddIn.Application.Session.CurrentUser;
            tbSkype.Text = CurUsr.Address;
            if (CurUsr.AddressEntry.Type == "EX")
            {
                tbSkype.Text = CurUsr.AddressEntry.GetExchangeUser().PrimarySmtpAddress;
            }
            tbName.Text = CurUsr.Name;
            cbAddress.Items.Clear();
            cbPOBox.Items.Clear();
            cbAddress.Items.Add("14A Gregory Street, Naenae");
            cbAddress.Items.Add("test address 2");
            cbPOBox.Items.Add("PO Box 30583 Lower Hutt");
            cbPOBox.Items.Add("Po box 2222");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Hide();
            //open template document
            Word.Application WdTemplate = new Word.Application();
            WdTemplate.Visible = false;
            var oSignatureObject = WdTemplate.EmailOptions.EmailSignature;
            var oSignatureEntry = oSignatureObject.EmailSignatureEntries;
            try
            {
                object oFilePath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase) + "\\Template.docx";
                Word.Document TemplateDocu = WdTemplate.Documents.Open(ref oFilePath, ReadOnly: true, Visible: false);
                //### replace key words here
                string SigName = tbName.Text;
                SetText(TemplateDocu, "[name]", SigName.ToUpper());
                SetText(TemplateDocu, "[position]", tbTitle.Text.ToUpper());
                SetText(TemplateDocu, "[Address:]", cbAddress.Text);
                SetText(TemplateDocu, "[Postal Address:]", cbPOBox.Text);
                SetText(TemplateDocu, "[DDI:]", tbPhone.Text);
                SetText(TemplateDocu, "[Mob:]", tbMobile.Text);
                SetText(TemplateDocu, "[Fax:]", tbFax.Text);
                SetText(TemplateDocu, "[Skype:]", tbSkype.Text);
                foreach(Word.ContentControl cc in TemplateDocu.ContentControls)
                {
                    if (cc.LockContentControl)
                        cc.LockContentControl = false;
                    cc.Delete();
                }
                var SigEntry = oSignatureEntry.Add(SigName, TemplateDocu.Content);
                oSignatureObject.NewMessageSignature = SigName;
                //no last row & column for reply & forward mail
                if (TemplateDocu.Content.Tables.Count > 0)
                {
                    var tb = TemplateDocu.Content.Tables[1];
                    int iLastRow = tb.Rows.Count;
                    for(int i = tb.Range.Cells.Count; i > 1; i--)
                    {
                        if (tb.Range.Cells[i].RowIndex == iLastRow)
                        {
                            tb.Range.Cells[i].Delete();
                        }
                    }
                    int iLastColumn = tb.Columns.Count;
                    for(int j = tb.Range.Cells.Count; j > 1; j--)
                    {
                        if (tb.Range.Cells[j].ColumnIndex == iLastColumn)
                        {
                            tb.Range.Cells[j].Delete();
                        }
                    }
                }
                oSignatureEntry.Add(SigName + "_reply", TemplateDocu.Content);
                oSignatureObject.ReplyMessageSignature = SigName + "_reply";
                //set default compose & reply mail font
                WdTemplate.Application.EmailOptions.ComposeStyle.Font.Name = "Arial";
                WdTemplate.Application.EmailOptions.ComposeStyle.Font.Size = 13;
                WdTemplate.Application.EmailOptions.ComposeStyle.Font.Color = Word.WdColor.wdColorGray80; //HEX:333333 - Oct 3355443 - RGB(51,51,51)
                WdTemplate.Application.EmailOptions.ComposeStyle.Font.Bold = 0;
                WdTemplate.Application.EmailOptions.ReplyStyle.Font.Name = "Arial";
                WdTemplate.Application.EmailOptions.ReplyStyle.Font.Size = 13;
                WdTemplate.Application.EmailOptions.ReplyStyle.Font.Color = Word.WdColor.wdColorGray80;
                WdTemplate.Application.EmailOptions.ReplyStyle.Font.Bold = 0;

                TemplateDocu.Close(SaveChanges: false);
                MessageBox.Show(this,"Signature \"" + SigName + "\" created successfully!");
            }
            catch(Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex.Message);
                MessageBox.Show("Failed to create signature: " + Environment.NewLine + ex.Message);
            }
            WdTemplate.Quit(SaveChanges: false);
            GC.Collect();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Hide();
        }

        private void tbName_TextChanged(object sender, EventArgs e)
        {
            Regex InvalidFileName = new Regex("[" + Regex.Escape(new string(Path.GetInvalidFileNameChars())) + "]");
            if (InvalidFileName.IsMatch(tbName.Text))
            {
                tbName.Focus();
                MessageBox.Show("Please do not use following characters:<>:/\\|?\"*");
            }
        }

        private void tbEmail_Leave(object sender, EventArgs e)
        {
            if (!IsEmail(tbSkype.Text))
            {
                tbSkype.Focus();
                MessageBox.Show("Email address no valid.");
            }
        }

        bool IsEmail(string s)
        {
            try
            {
                var add = new System.Net.Mail.MailAddress(s);
                return add.Address == s;
            }
            catch { return false; }
        }

        int SetText(Word.Document doc, string foo, string bar)
        {
            Word.ContentControl cc;
            if (doc.SelectContentControlsByTitle(foo).Count > 0)
            {
                cc = doc.SelectContentControlsByTitle(foo)[1];
                if (cc.LockContents)
                    cc.LockContents = false;
                cc.Range.Text = bar + (bar == "" ? "" : " ");
                string sTmp = foo.Replace("[", "").Replace("]", "").Trim();
                if (bar.Trim() == "" && doc.SelectContentControlsByTitle(sTmp).Count > 0)
                {
                    cc = doc.SelectContentControlsByTitle(sTmp)[1];
                    if (cc.LockContents)
                        cc.LockContents = false;
                    cc.Range.Text = "";
                }
                return 1;
            }
            else return -1;
        }
    }
}
