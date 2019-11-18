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
        //修改的标志位，判断是修改或者插入 数据
        public bool Insert_into_Switch ;


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

        private void button2_Click(object sender, EventArgs e)
        {
            //插入的标志位
            Insert_into_Switch = true;
            this.groupBox1.Enabled = true;
        }

        //提交
        private void button6_Click(object sender, EventArgs e)
        {
            //Excle文件
            var filepath = "Excel表格.xls";
            string sql;

            if (Insert_into_Switch==true)
            {
                //构建SQL语句
                sql = "insert into [学号信息$](学号,姓名,班级,电话号码,状态) values ({0},'{1}','{2}','{3}','{4}')";
                sql = string.Format(sql, this.textBox2.Text, this.textBox3.Text, this.textBox4.Text, this.textBox5.Text,"正常");

            }
            else
            {
                //构建修改语句
                sql = "update [学号信息$] set 姓名='{0}',班级='{1}',电话号码='{2}',状态='正常' where 学号={3} ";
                sql = string.Format(sql, this.textBox3.Text, this.textBox4.Text, this.textBox5.Text, this.textBox2.Text);
            }


            //执行更新数据
            Excel.Upadte(sql, filepath);

            //查询全部信息
            this.textBox1.Text = "";
            //触发button1按钮通过代码的方式
            button1_Click(null, null);
        }

        //删除
        private void button3_Click(object sender, EventArgs e)
        {
            //Excle文件
            var filepath = "Excel表格.xls";
            //判断学号是否有填入东西
            if (this.textBox1.Text=="")
            {
                MessageBox.Show("请输入要删除的学号");
            }
            else
            {
                //构建删除的SQL语句
                string sql = "UPDATE [学号信息$] set 状态='删除' where 学号={0}";
                sql = string.Format(sql, this.textBox1.Text);
                //执行sql语句
                Excel.Upadte(sql, filepath);
                //清空查询编号
                this.textBox1.Text = "";
                //执行查询操作
                button1_Click(null, null);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Insert_into_Switch = false;
            this.groupBox1.Enabled = true;
        }
    }
}
