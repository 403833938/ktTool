using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace hmd1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string sql = "select * from 黑名单";
            DataTable dt = ExecuteReader1(sql).Tables[0];
            toexcel(dt);
        }
        string strcon1 = "";
        public System.Data.DataSet ExecuteReader1(string sql)
        {
            SqlConnection sqlcon = new SqlConnection(strcon1);
            try // 正常运行
            {
                sqlcon.Open();

                SqlDataAdapter adapter = new SqlDataAdapter(sql, sqlcon);
                DataSet ds = new DataSet();
                adapter.Fill(ds);
                sqlcon.Close();
                return ds;
            }
            catch (OleDbException ole) // 数据库操作异常处理
            {
                if (sqlcon.State == ConnectionState.Open)
                {
                    // 关闭数据库连接
                    sqlcon.Close();
                }
                // 返回失败
                //LogBLL.debug(ole.ToString());
                return null;
            }
            catch (Exception ee)// 异常处理
            {

                if (sqlcon.State == ConnectionState.Open)
                {
                    // 关闭数据库连接

                    sqlcon.Close();
                }
                // 返回失败
                //LogBLL.debug(ee.ToString());
                return null;
            }
        finally // 执行完毕清除在try块中分配的任何资源
            {
                if (sqlcon.State == ConnectionState.Open)
                {
                    // 关闭数据库连接
                    sqlcon.Close();
                }
            }
        }


        private void toexcel(DataTable dt)
        {
            if (dt != null)
            {
                string localFilePath, fileNameExt, FilePath;
                SaveFileDialog sfd = new SaveFileDialog();

                //设置文件类型 
                sfd.Filter = "Excel文件(*.xls)|*.xls|Excel文件(*.xlsx)|*.xlsx";
                //name = Path.GetFileName(nametxt);
                //默认文件名
                sfd.FileName = DateTime.Now.ToString("yyyyMMddHHmmss") + "_反馈" + ".xls";
                //保存对话框是否记忆上次打开的目录 
                //sfd.RestoreDirectory = true;

                //点了保存按钮进入 
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    localFilePath = sfd.FileName.ToString(); //获得文件路径 
                    fileNameExt = localFilePath.Substring(localFilePath.LastIndexOf("\\") + 1); //获取文件名，不带路径

                    //获取文件路径，不带文件名 
                    FilePath = localFilePath.Substring(0, localFilePath.LastIndexOf("\\"));
                    sxst.ExcelHelper ex = new sxst.ExcelHelper();

                    ex.ExportDataTableToExcel(dt, fileNameExt, localFilePath);

                }
            }

        }
    }
}
