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
            tbEmail.Text = CurUsr.Address;
            if (CurUsr.AddressEntry.Type == "EX")
            {
                tbEmail.Text = CurUsr.AddressEntry.GetExchangeUser().PrimarySmtpAddress;
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
                var Range = TemplateDocu.Content;
                SetText(Range, "[name]", SigName.ToUpper());
                SetText(Range, "[position]", tbTitle.Text.ToUpper());
                SetText(Range, "[address]", cbAddress.Text);
                SetText(Range, "[PO Box]", cbPOBox.Text);
                SetText(Range, "[phone]", tbPhone.Text);
                SetText(Range, "[mobile]", tbMobile.Text);
                SetText(Range, "[fax]", tbFax.Text);
                SetText(Range, "[email]", tbEmail.Text);
                //if (!chkLinkedIn.Checked)
                //{
                //    var rg = TemplateDocu.Content;
                //    rg.Find.ClearAllFuzzyOptions();
                //    rg.Find.ClearFormatting();
                //    rg.Find.Wrap = Word.WdFindWrap.wdFindStop;
                //    rg.Find.Text = "Follow us on LinkedIn";
                //    rg.Find.Execute();
                //    if (rg.Find.Found)
                //    {
                //        //rg.SetRange(rg.Start - 1, rg.End + 2);
                //        rg.Delete();
                //    }
                //}
                var SigEntry = oSignatureEntry.Add(SigName, TemplateDocu.Content);
                oSignatureObject.NewMessageSignature = SigName;
                //no last row for reply & forward mail
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
                }
                oSignatureEntry.Add(SigName + "_reply", TemplateDocu.Content);
                oSignatureObject.ReplyMessageSignature = SigName + "_reply";
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
            if (!IsEmail(tbEmail.Text))
            {
                tbEmail.Focus();
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

        int SetText(Word.Range rg, string foo, string bar)
        {
            rg.Find.ClearAllFuzzyOptions();
            rg.Find.ClearFormatting();
            rg.Find.Wrap = Word.WdFindWrap.wdFindStop;
            rg.Find.Forward = true;
            rg.Find.MatchWholeWord = true;
            rg.Find.MatchCase = true;
            rg.Find.Text = foo;
            rg.Find.Replacement.Text = bar;
            rg.Find.Execute(Replace: Word.WdReplace.wdReplaceAll);
            if (rg.Find.Found)
            {
                return 1;
            }
            else
            {
                return -1;
            }
        }
    }
}
