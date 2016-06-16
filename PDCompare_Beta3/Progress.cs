using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using System.IO;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;

namespace PDCompare_Beta3
{
    public partial class Progress : Form
    {
        //private string pdOld;
        //private string pdNew;
        //private string resultFile;
        //private int recordCount;

        public Progress(string pdOld, string pdNew, string resultFile)
        {
            InitializeComponent();

            //this.pdOld = pdOld;
            //this.pdNew = pdNew;
            //this.resultFile = resultFile;

            string[] parameter = new string[]{pdOld, pdNew, resultFile};

            //object obj1 = (object)parameter;
            //string[] tempArray = null;
            //List<string> myList = new List<string>(
            //object[] tempArray = (object[])obj1;
            //MessageBox.Show(tempArray[0].ToString() + " "+tempArray[1].ToString() + " "+ tempArray[2]);

            Thread myThread = new Thread(PDSummaryCompare);
            myThread.IsBackground = true;
            myThread.Start(parameter);
        }

        private delegate void PDSummaryCompareDelegate(object parameter);

        /// <summary>  
        /// 进行循环  
        /// </summary>  
        /// <param name="number"></param>  
        private void PDSummaryCompare(object parameter)
        {
            if (progressBar1.InvokeRequired)
            {
                PDSummaryCompareDelegate myDelegate = PDSummaryCompare;
                progressBar1.Invoke(myDelegate, parameter);
            }
            else // here starts real PDSummaryCompare logic
            {
                // Divide parameter to three strings for use
                object[] objectArray = (object[])parameter;
                string pdOld = objectArray[0].ToString();
                string pdNew = objectArray[1].ToString();
                string resultFile = objectArray[2].ToString();

                string strConOld = " Provider = Microsoft.Jet.OLEDB.4.0 ; Data Source = " + pdOld + ";" + "Extended Properties=Excel 8.0";
                string strConNew = " Provider = Microsoft.Jet.OLEDB.4.0 ; Data Source = " + pdNew + ";" + "Extended Properties=Excel 8.0";
                //MessageBox.Show(pdOld);
                OleDbConnection myConnOld = new OleDbConnection(strConOld);
                myConnOld.Open();
                OleDbConnection myConnNew = new OleDbConnection(strConNew);
                myConnNew.Open();

                string strCom = "select * from [workSheet1$]";

                OleDbDataAdapter myCommandOld = new OleDbDataAdapter(strCom, myConnOld);
                DataSet dsOld = new DataSet();
                myCommandOld.Fill(dsOld, "ExcelInfo");
                this.Cursor = Cursors.WaitCursor;
                int rowCountOld = dsOld.Tables[0].Rows.Count;

                OleDbDataAdapter myCommandNew = new OleDbDataAdapter(strCom, myConnNew);
                DataSet dsNew = new DataSet();
                myCommandNew.Fill(dsNew, "ExcelInfo");
                this.Cursor = Cursors.WaitCursor;
                int rowCountNew = dsNew.Tables[0].Rows.Count;

                string[] split = pdNew.Split('\\');

                string resultFileName = pdNew.Substring(0, pdNew.Length - split[split.Length - 1].Length) + resultFile + ".xlsx";
                //MessageBox.Show(resultFileName);

                //string resultFileName = @"C:\Users\jliu15\Desktop\0607\resultFileTest.xlsx" + @"ttt";

                if (File.Exists(Path.GetFullPath(resultFileName)))
                {
                    File.Delete(Path.GetFullPath(resultFileName));
                }

                object Nothing = System.Reflection.Missing.Value;
                var app = new Excel.Application();
                app.Visible = false;

                Excel.Workbook workBook = app.Workbooks.Add(Nothing);                   // Sort Order
                Excel.Worksheet wsSortOrder = (Excel.Worksheet)workBook.Sheets[1];
                wsSortOrder.Name = "Sort Order";
                wsSortOrder.Tab.Color = System.Drawing.Color.Yellow;

                workBook.Sheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);    // Localization
                Excel.Worksheet wsLocalization = (Excel.Worksheet)workBook.Sheets[1];
                wsLocalization.Name = "Localization";
                wsLocalization.Tab.Color = System.Drawing.Color.Yellow;

                workBook.Sheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);     // Option Code
                Excel.Worksheet wsOptionCode = (Excel.Worksheet)workBook.Sheets[1];
                wsOptionCode.Name = "Option Code";
                wsOptionCode.Tab.Color = System.Drawing.Color.Yellow;

                workBook.Sheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);        // Removed
                Excel.Worksheet wsRemoved = (Excel.Worksheet)workBook.Sheets[1];
                wsRemoved.Name = "Removed";
                wsRemoved.Tab.Color = System.Drawing.Color.Red;

                workBook.Sheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);  // Changed
                Excel.Worksheet wsChanged = (Excel.Worksheet)workBook.Sheets[1];
                wsChanged.Name = "Changed";
                wsChanged.Tab.Color = System.Drawing.Color.Blue;

                workBook.Sheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);    // Added
                Excel.Worksheet wsAdded = (Excel.Worksheet)workBook.Sheets[1];
                wsAdded.Name = "Added";
                wsAdded.Tab.Color = System.Drawing.Color.Green;

                // headline  

                wsAdded.Cells[1, 1] = "Added:";
                wsChanged.Cells[1, 1] = "Changed:";
                wsRemoved.Cells[1, 1] = "Removed:";
                wsOptionCode.Cells[1, 1] = "Option Code:";
                wsLocalization.Cells[1, 1] = "Localization";
                wsSortOrder.Cells[1, 1] = "Sort Order";

                #region font settings

                Excel.Range rangeWsAdded = (Excel.Range)wsAdded.get_Range("A1");
                rangeWsAdded.Font.Size = 18;
                rangeWsAdded.Font.Bold = true;
                rangeWsAdded.Font.Underline = true;

                Excel.Range rangeAutoFit = (Excel.Range)wsAdded.get_Range("A1", "C1");
                rangeWsAdded.EntireColumn.AutoFit();

                // TBD

                #endregion

                //wsAdded.SaveAs(FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing);

                //workBook.
                //workBook.Close(false, Type.Missing, Type.Missing);

                //app.Quit();

                //MessageBox.Show("Finished!", "System Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //app.Visible = true;
                //Application.Exit();

                bool find = false;

                int cursorAdded = 2;
                int cursorChanged = 2;
                int cursorRemoved = 2;
                int cursorOptionCode = 2;
                int cursorLocalization = 2;
                int cursorSortOrder = 2;

                //progressBar1.Maximum = rowCountNew - 1;
                //this.recordCount = rowCountNew;

                for (int i = 1; i < rowCountNew; i++)
                {
                    find = false;

                    string componentNameNew = dsNew.Tables[0].Rows[i][0].ToString();

                    for (int j = 1; j < rowCountOld; j++)
                    {
                        string componentNameOld = dsOld.Tables[0].Rows[j][0].ToString();

                        if (componentNameNew == componentNameOld)
                        {
                            find = true;

                            // Part Number is different
                            if (dsNew.Tables[0].Rows[i][1].ToString() != dsOld.Tables[0].Rows[j][1].ToString())
                            {
                                wsChanged.Cells[cursorChanged, 1] = componentNameNew;
                                wsChanged.Cells[cursorChanged, 2] = dsOld.Tables[0].Rows[j][18].ToString() + "," + dsOld.Tables[0].Rows[j][19].ToString() + "," + dsOld.Tables[0].Rows[j][20].ToString()
                                    + " --> " + dsNew.Tables[0].Rows[i][18].ToString() + "," + dsNew.Tables[0].Rows[i][19].ToString() + "," + dsNew.Tables[0].Rows[i][20].ToString();
                                wsChanged.Cells[cursorChanged++, 3] = dsNew.Tables[0].Rows[i][1].ToString() + " --> " + dsOld.Tables[0].Rows[j][1].ToString();
                            }
                            // Option Code is different
                            if (dsNew.Tables[0].Rows[i][3].ToString() != dsOld.Tables[0].Rows[j][3].ToString())
                            {
                                wsOptionCode.Cells[cursorOptionCode, 1] = componentNameNew;
                                wsOptionCode.Cells[cursorOptionCode, 2] = dsNew.Tables[0].Rows[i][1].ToString();    // Part Number
                                wsOptionCode.Cells[cursorOptionCode++, 3] = dsOld.Tables[0].Rows[j][3].ToString() + " --> " + dsNew.Tables[0].Rows[i][3].ToString(); // oldOptionCode --> newOptionCode
                            }
                            // Localization
                            if (dsNew.Tables[0].Rows[i][4].ToString() != dsOld.Tables[0].Rows[j][4].ToString())
                            {
                                wsLocalization.Cells[cursorLocalization, 1] = componentNameNew;
                                wsLocalization.Cells[cursorLocalization, 2] = dsNew.Tables[0].Rows[i][1].ToString();    // Part Number
                                wsLocalization.Cells[cursorLocalization++, 3] = dsOld.Tables[0].Rows[j][4].ToString() + " --> " + dsNew.Tables[0].Rows[i][4].ToString(); // oldLocalization --> newLocalization
                            }
                            // Sort Order
                            if (dsNew.Tables[0].Rows[i][7].ToString() != dsOld.Tables[0].Rows[j][7].ToString())
                            {
                                wsSortOrder.Cells[cursorSortOrder, 1] = componentNameNew;
                                wsSortOrder.Cells[cursorSortOrder, 2] = dsNew.Tables[0].Rows[i][1].ToString();    // Part Number
                                wsSortOrder.Cells[cursorSortOrder++, 3] = dsOld.Tables[0].Rows[j][7].ToString() + " --> " + dsNew.Tables[0].Rows[i][7].ToString(); // oldSortOrder --> newSortOrder
                            }
                            break;
                        }
                    }

                    if (find == false)  // Add to Added
                    {
                        wsAdded.Cells[cursorAdded, 1] = componentNameNew;
                        wsAdded.Cells[cursorAdded, 2] = dsNew.Tables[0].Rows[i][18].ToString() + "," + dsNew.Tables[0].Rows[i][19].ToString() + "," + dsNew.Tables[0].Rows[i][20].ToString();
                        wsAdded.Cells[cursorAdded++, 3] = dsNew.Tables[0].Rows[i][1].ToString();
                    }
                }
                workBook.SaveAs(resultFileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing);
                MessageBox.Show("导入成功！", "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                app.Quit();
                //this.Close();
                Application.Exit();
                //MessageBox.Show(DateTime.Now.Subtract(dt).ToString());  //循环结束截止时间 



                //progressBar1.Maximum = (int)number;
                //for (int i = 0; i < (int)number; i++)
                //{
                    //progressBar1.Value = i;


                    //Application.DoEvents();
                //}

                
            }
        }
    }
}
