using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;

namespace PDCompare_Beta3
{
    public partial class TestResultExcelForm : Form
    {
        public TestResultExcelForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //create  
            string FileName = @"C:\Users\jliu15\Desktop\0607\resultFileTest.xlsx";
            object Nothing = System.Reflection.Missing.Value;
            var app = new Excel.Application();
            app.Visible = false;
            Excel.Workbook workBook = app.Workbooks.Add(Nothing);
            Excel.Worksheet wsAdded = (Excel.Worksheet)workBook.Sheets[1];
            wsAdded.Name = "Added";
            wsAdded.Tab.Color = System.Drawing.Color.Green;
            //headline  
            
            wsAdded.Cells[1, 1] = "FileName";
            
            wsAdded.Cells[1, 2] = "FindString";
            wsAdded.Cells[1, 3] = "ReplaceString";

            wsAdded.SaveAs(FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing);
            workBook.Close(false, Type.Missing, Type.Missing);
            app.Quit();

            MessageBox.Show("Finished!", "System Message", MessageBoxButtons.OK, MessageBoxIcon.Information);

            Application.Exit();
        }
    }
}
