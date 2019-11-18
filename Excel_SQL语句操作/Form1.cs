using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
//引用动态链接库
using Excel_ClassLibrary;

namespace Excel_SQL语句操作
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        //创建Excel文件
        private void button5_Click(object sender, EventArgs e)
        {
            //Excel地址
            var filepath = "Excel表格.xls";
            //构建SQL语句
            string sql = "CREATE TABLE 学号信息([学号] INT,[姓名] VarChar,[班级] VarChar,[电话号码] VarChar,[状态] VarChar)";
            //调用数据库的动态链接库
            Excel.Upadte(sql, filepath);
        }
        //查询
        private void button1_Click(object sender, EventArgs e)
        {
            //查询全部信息，还是按学号条件查询
            if (this.textBox1.Text=="")
            {
                //Excel地址
                var filepath = "Excel表格.xls";
                //SQl语句
                string sql = "select 学号,姓名,班级,电话号码 from [学号信息$] where 状态='正常'";
                //执行SQL语句
                this.dataGridView1.DataSource = Excel.GetDataTable(sql, filepath);
            }
            else
            {
                //根据学号查询
                //Excel地址
                var filepath = "Excel表格.xls";
                //SQl语句
                string sql = "select 学号,姓名,班级,电话号码 from [学号信息$] where 学号="+this.textBox1.Text;
                //执行SQL语句
                this.dataGridView1.DataSource = Excel.GetDataTable(sql, filepath);
            }

        }
    }
}
