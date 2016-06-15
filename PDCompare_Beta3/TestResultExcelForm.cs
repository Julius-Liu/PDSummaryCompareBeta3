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
using System.IO;

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
            string FileName = @"C:\Users\jliu15\Desktop\0607\resultFileTest.xlsx";

            // if exists, delete it.
            if (File.Exists(Path.GetFullPath(FileName)))
            {
                File.Delete(Path.GetFullPath(FileName));
            }

            object Nothing = System.Reflection.Missing.Value;
            var app = new Excel.Application();
            app.Visible = false;
            Excel.Workbook workBook = app.Workbooks.Add(Nothing);
            Excel.Worksheet wsAdded = (Excel.Worksheet)workBook.Sheets[1];
            wsAdded.Name = "Added";
            wsAdded.Tab.Color = System.Drawing.Color.Green;
            //headline  

            workBook.Sheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);  // Changed
            Excel.Worksheet wsChanged = (Excel.Worksheet)workBook.Sheets[1];
            wsChanged.Name = "Changed";
            wsChanged.Tab.Color = System.Drawing.Color.Blue;

            wsAdded.Cells[1, 1] = "Added:";
            
            wsAdded.Cells[1, 2] = "FindString";
            wsAdded.Cells[1, 3] = "ReplaceString";
            wsAdded.Cells[2, 1] = "Corinth Classroom";

            Excel.Range range = (Excel.Range)wsAdded.get_Range("A1");
            range.Font.Size = 18;
            range.Font.Bold = true;
            range.Font.Underline = true;

            Excel.Range rangeAutoFit = (Excel.Range)wsAdded.get_Range("A1", "C1");
            range.EntireColumn.AutoFit();

            wsAdded.SaveAs(FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing);

            //workBook.
            //workBook.Close(false, Type.Missing, Type.Missing);
            
            //app.Quit();
            
            MessageBox.Show("Finished!", "System Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
            app.Visible = true;
            Application.Exit();
        }
    }
}
