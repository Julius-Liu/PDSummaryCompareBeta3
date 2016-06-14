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
        public Main()
        {
            InitializeComponent();
        }

        private void btnBrowseOld_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.Filter = "Excel files|*.xlsx;*.xls";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                tbOld.Text = openFileDialog1.FileName;
            }
            //this.Close();
        }

        private void btnBrowseNew_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.Filter = "Excel files|*.xlsx;*.xls";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                tbNew.Text = openFileDialog1.FileName;
            }
            //this.Close();
        }

        private void btnGo_Click(object sender, EventArgs e)
        {
            Progress progress1 = new Progress(tbOld.Text, tbNew.Text, tbResult.Text);
            progress1.Show();
            /*
            openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.Filter = "excel files (*.excel)|*.xls";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string strCon = " Provider = Microsoft.Jet.OLEDB.4.0 ; Data Source = " + openFileDialog1.FileName + ";" + "Extended Properties=Excel 8.0";
                OleDbConnection myConn = new OleDbConnection(strCon);
                myConn.Open();
                string strCom = "select * from [Sheet1$] ";
                OleDbDataAdapter myCommand = new OleDbDataAdapter(strCom, myConn);
                DataSet ds = new DataSet();
                myCommand.Fill(ds, "ExcelInfo");
                MySqlConnection con = MySQL.getMySqlCon();
                con.Open();
                MySqlCommand sc = null, cmdUpdate = null, cmdInsert = null;
                MySqlDataReader sdr = null;
                this.Cursor = Cursors.WaitCursor;

                int rowCount = ds.Tables[0].Rows.Count;
                label1.Visible = true;
                label2.Visible = true;
                label2.Text = rowCount.ToString();
                progressBar1.Maximum = rowCount - 1;

                for (int i = 0; i < rowCount; i++)
                {
                    label1.Text = (i + 1).ToString();
                    if (ds.Tables[0].Rows[i][0] == null || ds.Tables[0].Rows[i][0].ToString() == "")
                        continue;
                    string query = "select * from old where 编号=" + "'" + ds.Tables[0].Rows[i][1].ToString() + "'";
                    sc = new MySqlCommand(query, con);
                    string sql = null;
                    string birthday = ds.Tables[0].Rows[i][4].ToString();
                    sdr = sc.ExecuteReader();
                    if (birthday == "")
                    {
                        birthday = Utils.getBirthday(ds.Tables[0].Rows[i][2].ToString());
                    }
                    if (sdr.Read())
                    {
                        sdr.Close();
                        sql = "update old set 姓名=" + "'" + ds.Tables[0].Rows[i][0].ToString() + "'" + "," + "身份证号码=" + "'" + ds.Tables[0].Rows[i][2].ToString() + "'" + "," + "性别=" + "'" + ds.Tables[0].Rows[i][3].ToString() + "'" + "," + "出生年月=" + "'" + birthday + "'" + "," + "婚姻状况=" + "'" + ds.Tables[0].Rows[i][5].ToString() + "'" + "," + "文化程度=" + "'" + ds.Tables[0].Rows[i][6].ToString() + "'" + "," + "健康状况=" + "'" + ds.Tables[0].Rows[i][7].ToString() + "'" + "," + "生活自理程度=" + "'" + ds.Tables[0].Rows[i][8].ToString() + "'" + "," + "居住地址=" + "'" + ds.Tables[0].Rows[i][9].ToString() + "'" + "," + "家庭电话=" + "'" + ds.Tables[0].Rows[i][10].ToString() + "'" + "," + "所属街道=" + "'" + ds.Tables[0].Rows[i][11].ToString() + "'" + "," + "所属居委=" + "'" + ds.Tables[0].Rows[i][12].ToString() + "'" + "," + "户籍地址=" + "'" + ds.Tables[0].Rows[i][13].ToString() + "'" + "," + "手机号码=" + "'" + ds.Tables[0].Rows[i][14].ToString() + "'" + "," + "子女姓名=" + "'" + ds.Tables[0].Rows[i][15].ToString() + "'" + "," + "子女电话=" + "'" + ds.Tables[0].Rows[i][16].ToString() + "'" + "," + "探望频率=" + "'" + ds.Tables[0].Rows[i][17].ToString() + "'" + "," + "每天问候服务需求=" + "'" + ds.Tables[0].Rows[i][18].ToString() + "'" + "," + "每天问候服务落实=" + "'" + ds.Tables[0].Rows[i][19].ToString() + "'" + "," + "每天问候服务备注=" + "'" + ds.Tables[0].Rows[i][20].ToString() + "'" + "," + "精神慰藉服务需求=" + "'" + ds.Tables[0].Rows[i][21].ToString() + "'" + "," + "精神慰藉服务落实=" + "'" + ds.Tables[0].Rows[i][22].ToString() + "'" + "," + "精神慰藉服务备注=" + "'" + ds.Tables[0].Rows[i][23].ToString() + "'" + "," + "紧急救援服务需求=" + "'" + ds.Tables[0].Rows[i][24].ToString() + "'" + "," + "紧急救援服务落实=" + "'" + ds.Tables[0].Rows[i][25].ToString() + "'" + "," + "紧急救援服务备注=" + "'" + ds.Tables[0].Rows[i][26].ToString() + "'" + "," + "生活照料服务需求=" + "'" + ds.Tables[0].Rows[i][27].ToString() + "'" + "," + "生活照料服务落实=" + "'" + ds.Tables[0].Rows[i][28].ToString() + "'" + "," + "生活照料服务备注=" + "'" + ds.Tables[0].Rows[i][29].ToString() + "'" + "," + "居家养老服务需求=" + "'" + ds.Tables[0].Rows[i][30].ToString() + "'" + "," + "居家养老服务落实=" + "'" + ds.Tables[0].Rows[i][31].ToString() + "'" + "," + "居家养老服务备注=" + "'" + ds.Tables[0].Rows[i][32].ToString() + "'" + "," + "日间照料服务需求=" + "'" + ds.Tables[0].Rows[i][33].ToString() + "'" + "," + "日间照料服务落实=" + "'" + ds.Tables[0].Rows[i][34].ToString() + "'" + "," + "日间照料服务备注=" + "'" + ds.Tables[0].Rows[i][35].ToString() + "'" + "," + "其它服务=" + "'" + ds.Tables[0].Rows[i][36].ToString() + "'" + "," + "志愿者身份证=" + "'" + ds.Tables[0].Rows[i][37].ToString() + "'" + "," + "志愿者姓名=" + "'" + ds.Tables[0].Rows[i][38].ToString() + "'" + " where 编号=" + "'" + ds.Tables[0].Rows[i][1].ToString() + "'";
                        cmdUpdate = new MySqlCommand(sql, con);
                        cmdUpdate.ExecuteNonQuery();
                        cmdUpdate.Dispose();
                    }
                    else
                    {
                        sdr.Close();
                        sql = "insert into old(姓名,编号,身份证号码,性别,出生年月,婚姻状况,文化程度,健康状况,生活自理程度,居住地址,家庭电话,所属街道,所属居委,户籍地址,手机号码,子女姓名,子女电话,探望频率,每天问候服务需求,每天问候服务落实,每天问候服务备注,精神慰藉服务需求,精神慰藉服务落实,精神慰藉服务备注,紧急救援服务需求,紧急救援服务落实,紧急救援服务备注,生活照料服务需求,生活照料服务落实,生活照料服务备注,居家养老服务需求,居家养老服务落实,居家养老服务备注,日间照料服务需求,日间照料服务落实,日间照料服务备注,其它服务,志愿者身份证,志愿者姓名) values(" + "'" + ds.Tables[0].Rows[i][0].ToString() + "'" + "," + "'" + ds.Tables[0].Rows[i][1].ToString() + "'" + "," + "'" + ds.Tables[0].Rows[i][2].ToString() + "'" + "," + "'" + ds.Tables[0].Rows[i][3].ToString() + "'" + "," + "'" + birthday + "'" + ",'" + ds.Tables[0].Rows[i][5].ToString() + "'" + "," + "'" + ds.Tables[0].Rows[i][6].ToString() + "'" + "," + "'" + ds.Tables[0].Rows[i][7].ToString() + "'" + "," + "'" + ds.Tables[0].Rows[i][8].ToString() + "'" + "," + "'" + ds.Tables[0].Rows[i][9].ToString() + "'" + "," + "'" + ds.Tables[0].Rows[i][10].ToString() + "'" + "," + "'" + ds.Tables[0].Rows[i][11].ToString() + "'" + "," + "'" + ds.Tables[0].Rows[i][12].ToString() + "'" + "," + "'" + ds.Tables[0].Rows[i][13].ToString() + "'" + "," + "'" + ds.Tables[0].Rows[i][14].ToString() + "'" + "," + "'" + ds.Tables[0].Rows[i][15].ToString() + "'" + "," + "'" + ds.Tables[0].Rows[i][16].ToString() + "'" + "," + "'" + ds.Tables[0].Rows[i][17].ToString() + "'" + "," + "'" + ds.Tables[0].Rows[i][18].ToString() + "'" + "," + "'" + ds.Tables[0].Rows[i][19].ToString() + "'" + "," + "'" + ds.Tables[0].Rows[i][20].ToString() + "'" + "," + "'" + ds.Tables[0].Rows[i][21].ToString() + "'" + "," + "'" + ds.Tables[0].Rows[i][22].ToString() + "'" + "," + "'" + ds.Tables[0].Rows[i][23].ToString() + "'" + "," + "'" + ds.Tables[0].Rows[i][24].ToString() + "'" + "," + "'" + ds.Tables[0].Rows[i][25].ToString() + "'" + "," + "'" + ds.Tables[0].Rows[i][26].ToString() + "'" + "," + "'" + ds.Tables[0].Rows[i][27].ToString() + "'" + "," + "'" + ds.Tables[0].Rows[i][28].ToString() + "'" + "," + "'" + ds.Tables[0].Rows[i][29].ToString() + "'" + "," + "'" + ds.Tables[0].Rows[i][30].ToString() + "'" + "," + "'" + ds.Tables[0].Rows[i][31].ToString() + "'" + "," + "'" + ds.Tables[0].Rows[i][32].ToString() + "'" + "," + "'" + ds.Tables[0].Rows[i][33].ToString() + "'" + "," + "'" + ds.Tables[0].Rows[i][34].ToString() + "'" + "," + "'" + ds.Tables[0].Rows[i][35].ToString() + "'" + "," + "'" + ds.Tables[0].Rows[i][36].ToString() + "'" + "," + "'" + ds.Tables[0].Rows[i][37].ToString() + "'" + "," + "'" + ds.Tables[0].Rows[i][38].ToString() + "'" + ")";
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
                myConn.Close();
                MessageBox.Show("导入成功！", "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            this.Close();
             */
        }

        private void createTestResultFileToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            TestResultExcelForm testResultExcelForm1 = new TestResultExcelForm();
            testResultExcelForm1.Show();
        }
    }
}
