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
    public partial class Progress : Form
    {
        public Progress(string pdOld, string pdNew, string resultFile)
        {
            InitializeComponent();                  

            string strConOld = " Provider = Microsoft.Jet.OLEDB.4.0 ; Data Source = " + pdOld + ";" + "Extended Properties=Excel 8.0";
            string strConNew = " Provider = Microsoft.Jet.OLEDB.4.0 ; Data Source = " + pdNew + ";" + "Extended Properties=Excel 8.0";

            OleDbConnection myConnOld = new OleDbConnection(strConOld);
            myConnOld.Open();
            OleDbConnection myConnNew = new OleDbConnection(strConNew);
            myConnNew.Open();

            string strCom = "select * from [Sheet1$]";

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

            /*
            //create  
            object Nothing = System.Reflection.Missing.Value;
            var app = new Excel.Application();
            app.Visible = false;
            Excel.Workbook workBook = app.Workbooks.Add(Nothing);
            Excel.Worksheet worksheet = (Excel.Worksheet)workBook.Sheets[1];
            worksheet.Name = "Work";
            //headline  
            worksheet.Cells[1, 1] = "FileName";
            worksheet.Cells[1, 2] = "FindString";
            worksheet.Cells[1, 3] = "ReplaceString";

            worksheet.SaveAs(FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing);
            workBook.Close(false, Type.Missing, Type.Missing);
            //app.Quit();  
            */

            progressBar1.Maximum = rowCountNew - 1;

            for (int i = 1; i < rowCountNew; i++)
            {
                string componentNameNew = dsNew.Tables[0].Rows[i][0].ToString();

                for (int j = 1; j < rowCountOld; j++)
                {
                    string componentNameOld = dsOld.Tables[0].Rows[j][0].ToString();


                }
                //label1.Text = (i + 1).ToString();
                //if (dsOld.Tables[0].Rows[i][0] == null || dsOld.Tables[0].Rows[i][0].ToString() == "")
                //    continue;
                /*
                string query = "select * from old where 编号=" + "'" + dsOld.Tables[0].Rows[i][1].ToString() + "'";
                sc = new MySqlCommand(query, con);
                string sql = null;
                string birthday = dsOld.Tables[0].Rows[i][4].ToString();
                sdr = sc.ExecuteReader();
                if (birthday == "")
                {
                    birthday = Utils.getBirthday(dsOld.Tables[0].Rows[i][2].ToString());
                }
                if (sdr.Read())
                {
                    sdr.Close();
                    sql = "update old set 姓名=" + "'" + dsOld.Tables[0].Rows[i][0].ToString() + "'" + "," + "身份证号码=" + "'" + dsOld.Tables[0].Rows[i][2].ToString() + "'" + "," + "性别=" + "'" + dsOld.Tables[0].Rows[i][3].ToString() + "'" + "," + "出生年月=" + "'" + birthday + "'" + "," + "婚姻状况=" + "'" + dsOld.Tables[0].Rows[i][5].ToString() + "'" + "," + "文化程度=" + "'" + dsOld.Tables[0].Rows[i][6].ToString() + "'" + "," + "健康状况=" + "'" + dsOld.Tables[0].Rows[i][7].ToString() + "'" + "," + "生活自理程度=" + "'" + dsOld.Tables[0].Rows[i][8].ToString() + "'" + "," + "居住地址=" + "'" + dsOld.Tables[0].Rows[i][9].ToString() + "'" + "," + "家庭电话=" + "'" + dsOld.Tables[0].Rows[i][10].ToString() + "'" + "," + "所属街道=" + "'" + dsOld.Tables[0].Rows[i][11].ToString() + "'" + "," + "所属居委=" + "'" + dsOld.Tables[0].Rows[i][12].ToString() + "'" + "," + "户籍地址=" + "'" + dsOld.Tables[0].Rows[i][13].ToString() + "'" + "," + "手机号码=" + "'" + dsOld.Tables[0].Rows[i][14].ToString() + "'" + "," + "子女姓名=" + "'" + dsOld.Tables[0].Rows[i][15].ToString() + "'" + "," + "子女电话=" + "'" + dsOld.Tables[0].Rows[i][16].ToString() + "'" + "," + "探望频率=" + "'" + dsOld.Tables[0].Rows[i][17].ToString() + "'" + "," + "每天问候服务需求=" + "'" + dsOld.Tables[0].Rows[i][18].ToString() + "'" + "," + "每天问候服务落实=" + "'" + dsOld.Tables[0].Rows[i][19].ToString() + "'" + "," + "每天问候服务备注=" + "'" + dsOld.Tables[0].Rows[i][20].ToString() + "'" + "," + "精神慰藉服务需求=" + "'" + dsOld.Tables[0].Rows[i][21].ToString() + "'" + "," + "精神慰藉服务落实=" + "'" + dsOld.Tables[0].Rows[i][22].ToString() + "'" + "," + "精神慰藉服务备注=" + "'" + dsOld.Tables[0].Rows[i][23].ToString() + "'" + "," + "紧急救援服务需求=" + "'" + dsOld.Tables[0].Rows[i][24].ToString() + "'" + "," + "紧急救援服务落实=" + "'" + dsOld.Tables[0].Rows[i][25].ToString() + "'" + "," + "紧急救援服务备注=" + "'" + dsOld.Tables[0].Rows[i][26].ToString() + "'" + "," + "生活照料服务需求=" + "'" + dsOld.Tables[0].Rows[i][27].ToString() + "'" + "," + "生活照料服务落实=" + "'" + dsOld.Tables[0].Rows[i][28].ToString() + "'" + "," + "生活照料服务备注=" + "'" + dsOld.Tables[0].Rows[i][29].ToString() + "'" + "," + "居家养老服务需求=" + "'" + dsOld.Tables[0].Rows[i][30].ToString() + "'" + "," + "居家养老服务落实=" + "'" + dsOld.Tables[0].Rows[i][31].ToString() + "'" + "," + "居家养老服务备注=" + "'" + dsOld.Tables[0].Rows[i][32].ToString() + "'" + "," + "日间照料服务需求=" + "'" + dsOld.Tables[0].Rows[i][33].ToString() + "'" + "," + "日间照料服务落实=" + "'" + dsOld.Tables[0].Rows[i][34].ToString() + "'" + "," + "日间照料服务备注=" + "'" + dsOld.Tables[0].Rows[i][35].ToString() + "'" + "," + "其它服务=" + "'" + dsOld.Tables[0].Rows[i][36].ToString() + "'" + "," + "志愿者身份证=" + "'" + dsOld.Tables[0].Rows[i][37].ToString() + "'" + "," + "志愿者姓名=" + "'" + dsOld.Tables[0].Rows[i][38].ToString() + "'" + " where 编号=" + "'" + dsOld.Tables[0].Rows[i][1].ToString() + "'";
                    cmdUpdate = new MySqlCommand(sql, con);
                    cmdUpdate.ExecuteNonQuery();
                    cmdUpdate.Dispose();
                }
                else
                {
                    sdr.Close();
                    sql = "insert into old(姓名,编号,身份证号码,性别,出生年月,婚姻状况,文化程度,健康状况,生活自理程度,居住地址,家庭电话,所属街道,所属居委,户籍地址,手机号码,子女姓名,子女电话,探望频率,每天问候服务需求,每天问候服务落实,每天问候服务备注,精神慰藉服务需求,精神慰藉服务落实,精神慰藉服务备注,紧急救援服务需求,紧急救援服务落实,紧急救援服务备注,生活照料服务需求,生活照料服务落实,生活照料服务备注,居家养老服务需求,居家养老服务落实,居家养老服务备注,日间照料服务需求,日间照料服务落实,日间照料服务备注,其它服务,志愿者身份证,志愿者姓名) values(" + "'" + dsOld.Tables[0].Rows[i][0].ToString() + "'" + "," + "'" + dsOld.Tables[0].Rows[i][1].ToString() + "'" + "," + "'" + dsOld.Tables[0].Rows[i][2].ToString() + "'" + "," + "'" + dsOld.Tables[0].Rows[i][3].ToString() + "'" + "," + "'" + birthday + "'" + ",'" + dsOld.Tables[0].Rows[i][5].ToString() + "'" + "," + "'" + dsOld.Tables[0].Rows[i][6].ToString() + "'" + "," + "'" + dsOld.Tables[0].Rows[i][7].ToString() + "'" + "," + "'" + dsOld.Tables[0].Rows[i][8].ToString() + "'" + "," + "'" + dsOld.Tables[0].Rows[i][9].ToString() + "'" + "," + "'" + dsOld.Tables[0].Rows[i][10].ToString() + "'" + "," + "'" + dsOld.Tables[0].Rows[i][11].ToString() + "'" + "," + "'" + dsOld.Tables[0].Rows[i][12].ToString() + "'" + "," + "'" + dsOld.Tables[0].Rows[i][13].ToString() + "'" + "," + "'" + dsOld.Tables[0].Rows[i][14].ToString() + "'" + "," + "'" + dsOld.Tables[0].Rows[i][15].ToString() + "'" + "," + "'" + dsOld.Tables[0].Rows[i][16].ToString() + "'" + "," + "'" + dsOld.Tables[0].Rows[i][17].ToString() + "'" + "," + "'" + dsOld.Tables[0].Rows[i][18].ToString() + "'" + "," + "'" + dsOld.Tables[0].Rows[i][19].ToString() + "'" + "," + "'" + dsOld.Tables[0].Rows[i][20].ToString() + "'" + "," + "'" + dsOld.Tables[0].Rows[i][21].ToString() + "'" + "," + "'" + dsOld.Tables[0].Rows[i][22].ToString() + "'" + "," + "'" + dsOld.Tables[0].Rows[i][23].ToString() + "'" + "," + "'" + dsOld.Tables[0].Rows[i][24].ToString() + "'" + "," + "'" + dsOld.Tables[0].Rows[i][25].ToString() + "'" + "," + "'" + dsOld.Tables[0].Rows[i][26].ToString() + "'" + "," + "'" + dsOld.Tables[0].Rows[i][27].ToString() + "'" + "," + "'" + dsOld.Tables[0].Rows[i][28].ToString() + "'" + "," + "'" + dsOld.Tables[0].Rows[i][29].ToString() + "'" + "," + "'" + dsOld.Tables[0].Rows[i][30].ToString() + "'" + "," + "'" + dsOld.Tables[0].Rows[i][31].ToString() + "'" + "," + "'" + dsOld.Tables[0].Rows[i][32].ToString() + "'" + "," + "'" + dsOld.Tables[0].Rows[i][33].ToString() + "'" + "," + "'" + dsOld.Tables[0].Rows[i][34].ToString() + "'" + "," + "'" + dsOld.Tables[0].Rows[i][35].ToString() + "'" + "," + "'" + dsOld.Tables[0].Rows[i][36].ToString() + "'" + "," + "'" + dsOld.Tables[0].Rows[i][37].ToString() + "'" + "," + "'" + dsOld.Tables[0].Rows[i][38].ToString() + "'" + ")";
                    cmdInsert = new MySqlCommand(sql, con);
                    cmdInsert.ExecuteNonQuery();
                    cmdInsert.Dispose();
                }
                progressBar1.Value = i;  // 
                Application.DoEvents();  // 
                Windows7Taskbar.SetProgressState(this.Handle, Windows7Taskbar.ThumbnailProgressState.Normal);
                Windows7Taskbar.SetProgressValue(this.Handle, ulong.Parse(i.ToString()), ulong.Parse((rowCount - 1).ToString()));
            }
            this.Cursor = Cursors.Default;
            sc.Dispose();
            con.Close();
            myConnOld.Close();
            MessageBox.Show("导入成功！", "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
           */
            }
        }
    }
}
