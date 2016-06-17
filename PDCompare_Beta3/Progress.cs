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
                wsLocalization.Tab.Color = System.Drawing.Color.Turquoise;

                workBook.Sheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);     // Option Code
                Excel.Worksheet wsOptionCode = (Excel.Worksheet)workBook.Sheets[1];
                wsOptionCode.Name = "Option Code";
                wsOptionCode.Tab.Color = System.Drawing.Color.GreenYellow;

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

                Excel.Range rangeWsAddedHL = (Excel.Range)wsAdded.get_Range("A1");
                rangeWsAddedHL.Font.Size = 18;
                rangeWsAddedHL.Font.Bold = true;
                rangeWsAddedHL.Font.Underline = true;

                //Excel.Range rangeWsAddedAF = (Excel.Range)wsAdded.get_Range("A1", "C1");
                //rangeWsAddedAF.EntireColumn.AutoFit();

                //Excel.Range rangeWsChangedHL = (Excel.Range)wsChanged.get_Range("A1");
                //rangeWsChangedHL.Font.Size = 18;
                //rangeWsChangedHL.Font.Bold = true;
                //rangeWsChangedHL.Font.Underline = true;

                wsChanged.Rows[1][1].Font.Size = 18;
                wsChanged.Rows[1][1].Font.Bold = true;
                wsChanged.Rows[1][1].Font.Underline = true;

                //wsChanged.Columns[1].ColumnWidth = 200;
                Excel.Range rangeWsChangedA1 = (Excel.Range)wsChanged.get_Range("A1");
                rangeWsChangedA1.ColumnWidth = 30;

                Excel.Range rangeWsOptionCodeHL = (Excel.Range)wsOptionCode.get_Range("A1");
                rangeWsOptionCodeHL.Font.Size = 18;
                rangeWsOptionCodeHL.Font.Bold = true;
                rangeWsOptionCodeHL.Font.Underline = true;

                Excel.Range rangeWsChangedAF = (Excel.Range)wsChanged.get_Range("A2", "C2");
                rangeWsChangedAF.EntireColumn.AutoFit();

                

                // TBD

                #endregion

                bool find = false;

                int cursorAdded = 2;
                int cursorChanged = 2;
                int cursorRemoved = 2;
                int cursorOptionCode = 2;
                int cursorLocalization = 2;
                int cursorSortOrder = 2;

                //progressBar1.Maximum = rowCountNew - 1;
                //this.recordCount = rowCountNew;

                progressBar1.Maximum = rowCountNew;

                for (int i = 1; i < rowCountNew; i++)
                {
                    progressBar1.Value = i+1;

                    find = false;

                    string componentPN_sub_New = dsNew.Tables[0].Rows[i-1][1].ToString().Substring(0, dsNew.Tables[0].Rows[i-1][1].ToString().Length - 1);

                    for (int j = 1; j < rowCountOld; j++)
                    {
                        string componentPN_sub_Old = dsOld.Tables[0].Rows[j-1][1].ToString().Substring(0, dsOld.Tables[0].Rows[j-1][1].ToString().Length - 1);

                        if (componentPN_sub_New == componentPN_sub_Old)
                        {
                            find = true;

                            // Part Number is different
                            // Add to Changed
                            if (dsNew.Tables[0].Rows[i-1][1].ToString() != dsOld.Tables[0].Rows[j-1][1].ToString())
                            {
                                wsChanged.Cells[cursorChanged, 1] = dsNew.Tables[0].Rows[i-1][0].ToString();  // component name
                                wsChanged.Cells[cursorChanged, 2] = dsOld.Tables[0].Rows[j-1][18].ToString().Trim() + "," + dsOld.Tables[0].Rows[j-1][19].ToString().Trim() + "," + dsOld.Tables[0].Rows[j-1][20].ToString().Trim()
                                    + " --> " + dsNew.Tables[0].Rows[i-1][18].ToString().Trim() + "," + dsNew.Tables[0].Rows[i-1][19].ToString().Trim() + "," + dsNew.Tables[0].Rows[i-1][20].ToString().Trim();
                                wsChanged.Cells[cursorChanged++, 3] = dsNew.Tables[0].Rows[i-1][1].ToString() + " --> " + dsOld.Tables[0].Rows[j-1][1].ToString();
                            }
                            // Option Code is different
                            // Add to Option Code
                            if (dsNew.Tables[0].Rows[i-1][3].ToString() != dsOld.Tables[0].Rows[j-1][3].ToString())
                            {
                                wsOptionCode.Cells[cursorOptionCode, 1] = dsNew.Tables[0].Rows[i-1][0].ToString();  // component name
                                wsOptionCode.Cells[cursorOptionCode, 2] = dsNew.Tables[0].Rows[i-1][1].ToString();    // Part Number
                                string optionCodeOld = dsOld.Tables[0].Rows[j - 1][3].ToString();
                                string optionCodeNew = dsNew.Tables[0].Rows[i - 1][3].ToString();
                                if (optionCodeOld == "")
                                {
                                    optionCodeOld = "null";
                                }
                                if (optionCodeNew == "")
                                {
                                    optionCodeNew = "null";
                                }
                                wsOptionCode.Cells[cursorOptionCode++, 3] = optionCodeOld + " --> " + optionCodeNew; // oldOptionCode --> newOptionCode
                            }
                            // Localization
                            // Add to Localization
                            if (dsNew.Tables[0].Rows[i-1][4].ToString() != dsOld.Tables[0].Rows[j-1][4].ToString())
                            {
                                wsLocalization.Cells[cursorLocalization, 1] = dsNew.Tables[0].Rows[i-1][0].ToString();  // component name
                                wsLocalization.Cells[cursorLocalization, 2] = dsNew.Tables[0].Rows[i-1][1].ToString();    // Part Number
                                string localizationOld = dsOld.Tables[0].Rows[j-1][4].ToString();
                                string localizationNew = dsNew.Tables[0].Rows[i-1][4].ToString();
                                if (localizationOld == "")
                                {
                                    localizationOld = "null";
                                }
                                if (localizationNew == "")
                                {
                                    localizationNew = "null";
                                }
                                wsLocalization.Cells[cursorLocalization++, 3] = localizationOld + " --> " + localizationNew; // oldLocalization --> newLocalization
                            }
                            // Sort Order
                            if (dsNew.Tables[0].Rows[i-1][7].ToString() != dsOld.Tables[0].Rows[j-1][7].ToString())
                            {
                                wsSortOrder.Cells[cursorSortOrder, 1] = dsNew.Tables[0].Rows[i-1][0].ToString();  // component name
                                wsSortOrder.Cells[cursorSortOrder, 2] = dsNew.Tables[0].Rows[i-1][1].ToString();    // Part Number
                                wsSortOrder.Cells[cursorSortOrder++, 3] = dsOld.Tables[0].Rows[j-1][7].ToString() + " --> " + dsNew.Tables[0].Rows[i-1][7].ToString(); // oldSortOrder --> newSortOrder
                            }
                            break;
                        }
                    }

                    if (find == false)  // Add to Added
                    {
                        wsAdded.Cells[cursorAdded, 1] = dsNew.Tables[0].Rows[i - 1][0].ToString();  // component name
                        wsAdded.Cells[cursorAdded, 2] = dsNew.Tables[0].Rows[i-1][18].ToString().Trim() + "," + dsNew.Tables[0].Rows[i-1][19].ToString().Trim() + "," + dsNew.Tables[0].Rows[i-1][20].ToString().Trim();    //version
                        wsAdded.Cells[cursorAdded++, 3] = dsNew.Tables[0].Rows[i-1][1].ToString();  // part number
                    }

                    Application.DoEvents();
                }

                // Get removed components
                for (int j = 1; j < rowCountOld; j++)
                {
                    find = false;

                    string componentPN_sub_Old = dsOld.Tables[0].Rows[j - 1][1].ToString().Substring(0, dsOld.Tables[0].Rows[j - 1][1].ToString().Length - 1);

                    for (int i = 1; i < rowCountNew; i++)
                    {
                        string componentPN_sub_New = dsNew.Tables[0].Rows[i - 1][1].ToString().Substring(0, dsNew.Tables[0].Rows[i - 1][1].ToString().Length - 1);

                        if (componentPN_sub_Old == componentPN_sub_New)
                        {
                            find = true;
                            break;
                        }
                    }

                    // Add to Removed
                    if (!find)
                    {
                        wsRemoved.Cells[cursorRemoved, 1] = dsOld.Tables[0].Rows[j - 1][0].ToString();  // component name
                        wsRemoved.Cells[cursorRemoved, 2] = dsOld.Tables[0].Rows[j - 1][18].ToString().Trim() + "," + dsOld.Tables[0].Rows[j - 1][19].ToString().Trim() + "," + dsOld.Tables[0].Rows[j - 1][20].ToString().Trim();  // version
                        wsRemoved.Cells[cursorRemoved++, 3] = dsOld.Tables[0].Rows[j - 1][1].ToString();  // part number
                    }
                }

                workBook.SaveAs(resultFileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing);
                MessageBox.Show("Finished!", "Application Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //app.Quit();
                app.Visible = true;
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
