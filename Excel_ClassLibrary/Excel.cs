﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//引入空间
using System.Data;
using System.Data.OleDb;

namespace Excel_ClassLibrary
{
    public class Excel
    {
        //更新数据
        public static int Upadte(string sql,string path)
        {
        //构建链接语句-
        string sConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" 
                                    + "Data Source=" + path + ";" 
                                    + "Extended Properties='Excel 8.0;HDR=Yes;IMEX=0'";
            //IMEX=0 为汇出模式，这个模式Excle只能用作"写入"用途
            //IMEX=1 为汇入模式，这个模式Excle只能用作"读取"用途
            //IMEX=2 为链接模式, 这个模式Excle同时支持"读写"用途
            //HDR=Yes 创建表头
            return 1;
        }
    }
}