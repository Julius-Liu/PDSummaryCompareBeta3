using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using System.Data.OleDb;

namespace PDCompare_Beta3
{
    public partial class Main : Form
    {
        private string customDir = "c:\\";

        public Main()
        {
            InitializeComponent();            
        }

        private void btnBrowseOld_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = customDir;
            openFileDialog1.Filter = "Excel files|*.xlsx;*.xls";
            openFileDialog1.FilterIndex = 1;
            //openFileDialog1.RestoreDirectory = true;
            openFileDialog1.FileName = "";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                tbOld.Text = openFileDialog1.FileName;
                string[] split = openFileDialog1.FileName.Split('\\');
                customDir = openFileDialog1.FileName.Substring(0, openFileDialog1.FileName.Length - split[split.Length - 1].Length);
            }
        }

        private void btnBrowseNew_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = customDir;
            openFileDialog1.Filter = "Excel files|*.xlsx;*.xls";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.FileName = "";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                tbNew.Text = openFileDialog1.FileName;
            }
        }

        private void btnGo_Click(object sender, EventArgs e)
        {
            Progress progress1 = new Progress(tbOld.Text, tbNew.Text, tbResult.Text);
            progress1.Show();            
        }

        private void createTestResultFileToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            TestResultExcelForm testResultExcelForm1 = new TestResultExcelForm();
            testResultExcelForm1.Show();
        }

        private void testFileNameToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string[] split = tbNew.Text.Split('\\');
            int yyyLen = split[split.Length - 1].Length;

            string resultFileName = tbNew.Text.Substring(0, tbNew.Text.Length - yyyLen) + tbResult.Text + ".xlsx";

            MessageBox.Show(resultFileName);
        }

        private void testProgressToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Progress progress1 = new Progress(tbOld.Text, tbNew.Text, tbResult.Text);
            progress1.Show();
        }

        private void tbResult_KeyPress(object sender, KeyPressEventArgs e)
        {
            // When you press Enter in Result File Name input box,
            // it'll have the same behavior as Click GO
            if (e.KeyChar == (char)13)
            {
                e.Handled = true;
                SendKeys.Send("{TAB} + {HOME}");
            }
        }
    }
}
