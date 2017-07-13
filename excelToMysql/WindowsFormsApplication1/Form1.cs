using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                comboBox1.Text = openFileDialog1.FileName;
                ExcelToDataSet(comboBox1.Text);
            }
        }
        DataTable dTable;
        #region 打开方法
        public DataTable ExcelToDataSet(string path) {
            string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + @path + ";" + "Extended Properties=Excel 8.0;";
            string strExcel = "";
            DataTable dt = null;
            try
            {
                OleDbConnection conn = new OleDbConnection(strConn);
                conn.Open();
                DataTable excelTabel = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                string sheetName = excelTabel.Rows[0][2].ToString().Trim();
                strExcel = "select * from [" + sheetName + "]";
                OleDbDataAdapter myCommand = new OleDbDataAdapter(strExcel,strConn);
               DataSet ds = new DataSet();
                myCommand.Fill(ds);
                dt = ds.Tables[0];
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (dt.Rows[i][0].ToString()=="")
                    {
                        dt.Rows.RemoveAt(i);
                    }
                }
                if (dt.Columns.Count > 1 && dt.Rows.Count >= 1)
                {
                    dataGridView1.DataSource = dt;
                    dataGridView1.RowHeadersWidth = 18;
                    dataGridView1.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;
                    dataGridView1.AllowUserToResizeRows = false;//不允许调整行大小
                    textBox1.Text = "共有 " + (dataGridView1.RowCount - 1).ToString() + " 条信息";
                }
                else {
                    MessageBox.Show(sheetName + " 表中没有数据，不能导入！", "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
            catch (Exception)
            {

                MessageBox.Show("读取Excel错误！", "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            dTable = dt;
            return dt;
        }
        #endregion

        private void button2_Click(object sender, EventArgs e)
        {
            ImportExcelToDB(dTable);
        }
        public bool ImportExcelToDB(DataTable dt) {
            int num = -1;
            try
            {
                foreach (DataRow row in dTable.Rows)
                {
                    string mysqlconnstring= "Server=127.0.0.1;Port=3306;Database=HR.Business; User=root;Password=;";
                    MySqlConnection conn = new MySqlConnection(mysqlconnstring);
                    if (conn.State== ConnectionState.Closed)
                    {
                        conn.Open();
                    }
                    //"INSERT INTO `tb_System_Account` VALUES ('001e932b-877a-46bf-abb7-b25c8e34bc38', null, '上海博科资讯股份有限公司', '上海博科资讯股份有限公司', '上海博科资讯股份有限公司', '', '', 'company17', 'e10adc3949ba59abbe56e057f20f883e', null, '2', '0', '', null, null, null, '\0', '0', null, null, '1', null, null, null, '2017-03-17 12:52:00', null);"
                    string sql = "insert into tb_System_Account(Id,Name,LoginName,Password,AccountType,CertifiedMobile) values(@Id,@Name,@LoginName,@Password,@AccountType,@CertifiedMobile)";
                    MySqlCommand cmd = new MySqlCommand(sql, conn);
                    cmd.Parameters.AddWithValue("@Id", Guid.NewGuid());
                    cmd.Parameters.AddWithValue("@Name", row["企业名称"].ToString());
                    cmd.Parameters.AddWithValue("@LoginName",  row["企业名称"].ToString());
                    cmd.Parameters.AddWithValue("@Password", "e10adc3949ba59abbe56e057f20f883e");
                    cmd.Parameters.AddWithValue("@AccountType", 2);
                    cmd.Parameters.AddWithValue("@CertifiedMobile", row["手机"].ToString());
                    int i = num = cmd.ExecuteNonQuery();
                    if (i < 0)
                    {
                        MessageBox.Show("xxx");
                    }
                    conn.Close();
                }
                if (num > 0)
                {
                    MessageBox.Show(" 导入成功！", "系统提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}
