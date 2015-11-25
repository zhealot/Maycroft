using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

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
            SetupForm.Show();
            SetupForm.Controls.Find("tbName", true)[0].Focus();
        }
    }
}
