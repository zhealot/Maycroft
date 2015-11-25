﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
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
                string SigName = "AAAA";
                System.Diagnostics.Debug.WriteLine(TemplateDocu.Name + " : " + TemplateDocu.Paragraphs[1].Range.Text);
                var SigEntry = oSignatureEntry.Add(SigName, TemplateDocu.Content);
                oSignatureObject.NewMessageSignature =SigName;
                oSignatureObject.ReplyMessageSignature =SigName;
                TemplateDocu.Close(SaveChanges: false);
            }
            catch(Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex.Message);
            }
            WdTemplate.Quit();
            //if (WdTemplate != null)
            //    System.Runtime.InteropServices.Marshal.ReleaseComObject(WdTemplate);
            GC.Collect();
        }
    }
}
