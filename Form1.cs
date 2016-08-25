using System;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace MaycroftOL
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            cbAddress.Items.Clear();
            cbPOBox.Items.Clear();
            cbAddress.Items.Add("14A Gregory Street, Naenae");
            cbAddress.Items.Add("6 Roy St, Palmerston North");
            cbPOBox.Items.Add("PO Box 30583 Lower Hutt");
            Word.Application WdTemplate = new Word.Application();
            WdTemplate.Visible = false;
            try
            {
                object oFiePath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase) + "\\Template.docx";
                Word.Document doc = WdTemplate.Documents.Open(ref oFiePath, ReadOnly: true, Visible: false);
                tbName.Text = ValueInTemplate(doc, "name");
                tbTitle.Text = ValueInTemplate(doc, "title");
                cbAddress.Text = ValueInTemplate(doc, "Address");
                cbPOBox.Text = ValueInTemplate(doc, "Postal Address");
                tbMobile.Text = ValueInTemplate(doc, "Mob");
                tbPhone.Text = ValueInTemplate(doc, "DDI");
                tbFax.Text = ValueInTemplate(doc, "Fax");
                tbSkype.Text = ValueInTemplate(doc, "Skype");
                tbProject.Text = "";
                if (doc.Hyperlinks.Count > 0)
                {
                    for (int i = 1; i <= doc.Hyperlinks.Count; i++)
                    {
                        if (doc.Hyperlinks[i].TextToDisplay.ToUpper().IndexOf("VIEW FEATURE PROJECT") >= 0)
                        {
                            tbProject.Text = doc.Hyperlinks[i].Address.Trim();
                            break;
                        }
                    }
                }
                doc.Close(SaveChanges: false);
                WdTemplate.Quit();
            }
            catch(Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex.Message);
            }
            cbProject.Checked = true;
            //var CurUsr = Globals.ThisAddIn.Application.Session.CurrentUser;
            //tbSkype.Text = CurUsr.Address;
            //tbName.Text = CurUsr.Name;
            //if (CurUsr.AddressEntry.Type == "EX")
            //{
            //    tbSkype.Text = CurUsr.AddressEntry.GetExchangeUser().PrimarySmtpAddress;
            //}
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
                Word.Document TemplateDocu = WdTemplate.Documents.Open(ref oFilePath, ReadOnly: false, Visible: false);
                //### replace key words here
                string SigName = tbName.Text;
                SetText(TemplateDocu, "name", SigName.ToUpper());
                SetText(TemplateDocu, "title", tbTitle.Text.ToUpper());
                SetText(TemplateDocu, "Address", cbAddress.Text);
                SetText(TemplateDocu, "Postal Address", cbPOBox.Text);
                SetText(TemplateDocu, "DDI", tbPhone.Text);
                SetText(TemplateDocu, "Mob", tbMobile.Text);
                SetText(TemplateDocu, "Fax", tbFax.Text);
                SetText(TemplateDocu, "Skype", tbSkype.Text);
                //replace 'link to project' hyperlink
                if (TemplateDocu.Hyperlinks.Count > 0)
                {
                    for (int i = 1; i <= TemplateDocu.Hyperlinks.Count; i++)
                    {
                        if (TemplateDocu.Hyperlinks[i].TextToDisplay.ToUpper().Trim().IndexOf("VIEW FEATURE PROJECT") >= 0)
                        {
                            if (cbProject.Checked && tbProject.Text.Trim().Length > 0)
                            {
                                TemplateDocu.Hyperlinks[i].Address = tbProject.Text;
                            }
                            else 
                            {
                                TemplateDocu.Hyperlinks[i].TextToDisplay = " ";
                            }
                            break;
                        }
                    }
                }
                //save changes back to template
                TemplateDocu.Save();

                foreach (Word.ContentControl cc in TemplateDocu.ContentControls)
                {
                    if (cc.LockContentControl)
                        cc.LockContentControl = false;
                    cc.Delete();
                }
                var SigEntry = oSignatureEntry.Add(SigName, TemplateDocu.Tables[1].Range);
                oSignatureObject.NewMessageSignature = SigName;
                //no last two rows & column for reply & forward mail
                if (TemplateDocu.Content.Tables.Count > 0)
                {
                    var tb = TemplateDocu.Content.Tables[1];
                    int iLastRow = tb.Rows.Count;
                    for(int i = tb.Range.Cells.Count; i > 1; i--)
                    {
                        if (tb.Range.Cells[i].RowIndex >= iLastRow - 1)
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
                oSignatureEntry.Add(SigName + "_reply", TemplateDocu.Tables[1].Range);
                oSignatureObject.ReplyMessageSignature = SigName + "_reply";
                //set default compose & reply mail font
                WdTemplate.Application.EmailOptions.ComposeStyle.Font.Name = "Arial";
                WdTemplate.Application.EmailOptions.ComposeStyle.Font.Size = 11;
                WdTemplate.Application.EmailOptions.ComposeStyle.Font.Color = Word.WdColor.wdColorGray80; //HEX:333333 - Oct 3355443 - RGB(51,51,51)
                WdTemplate.Application.EmailOptions.ComposeStyle.Font.Bold = 0;
                WdTemplate.Application.EmailOptions.ReplyStyle.Font.Name = "Arial";
                WdTemplate.Application.EmailOptions.ReplyStyle.Font.Size = 11;
                WdTemplate.Application.EmailOptions.ReplyStyle.Font.Color = Word.WdColor.wdColorGray80;
                WdTemplate.Application.EmailOptions.ReplyStyle.Font.Bold = 0;

                TemplateDocu.Close(SaveChanges:false);
                MessageBox.Show(this,"Signature \"" + SigName + "\" created successfully!");
            }
            catch(Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex.Message);
                if (WdTemplate.Documents.Count > 0)
                {
                    WdTemplate.Documents[1].Close(SaveChanges: false);
                }
                MessageBox.Show("Failed to create signature: " + Environment.NewLine + ex.Message);
            }
            WdTemplate.Quit(SaveChanges: false);
            WdTemplate = null;
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
            //if (!IsEmail(tbSkype.Text))
            //{
            //    tbSkype.Focus();
            //    MessageBox.Show("Email address no valid.");
            //}
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

        int SetText(Word.Document doc, string CCTitle, string CCText)
        {
            Word.ContentControl cc;
            if (doc.SelectContentControlsByTitle(CCTitle).Count > 0 )
            {
                cc = doc.SelectContentControlsByTitle(CCTitle)[1];
                if (CCText.Trim() == string.Empty)
                {
                    cc.Range.Text = "";
                    cc.Range.Font.Hidden = 1;                    
                }
                else
                {
                    if (cc.LockContents)
                        cc.LockContents = false;
                    cc.Range.Font.Hidden = 0;
                    cc.Range.Text = CCText.Trim() + " ";
                }
            }
            if (doc.SelectContentControlsByTitle(CCTitle + "_T").Count > 0)
            {
                cc = doc.SelectContentControlsByTitle(CCTitle + "_T")[1];
                if (cc.LockContents)
                    cc.LockContents = false;
                //hide cc's text if there's no content for this cc
                string ccTmp = string.Empty;
                if (CCText.Trim() == "")
                {
                    cc.Range.Text = "";
                    cc.Range.Font.Hidden = 1;
                }
                else
                {
                    if (CCTitle == "name" || CCTitle == "title")
                    {
                        ccTmp = CCTitle;
                    }
                    else
                    {
                        ccTmp = CCTitle + ": ";
                    }
                    cc.Range.Font.Hidden = 0;
                    cc.Range.Text = ccTmp;
                }
                return 1;
            }
            return -1;
        }

        string ValueInTemplate(Word.Document doc, string CCTitle)
        {
            if (doc.SelectContentControlsByTitle(CCTitle).Count > 0)
            {
                Word.ContentControl cc;
                cc = doc.SelectContentControlsByTitle(CCTitle)[1];
                if (cc.Range.Text == null)
                {
                    return "";
                }
                else
                {
                    return cc.Range.Text.Trim();
                }
            }
            else return "";
        }

        private void cbProject_CheckedChanged(object sender, EventArgs e)
        {
            tbProject.Enabled = cbProject.Checked;
        }
    }
}
