
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;
using System.Resources;
using System.Security.Cryptography;
using System.IO;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;
using System.Xml.Linq;
using System.Diagnostics;
using System.Net.WebSockets;

namespace ZenathRubberfactory
{
    public class Sql
    {
        public static string constring;
        public static SqlConnection con;
        static SqlDataAdapter adapter = new SqlDataAdapter();
        static System.Data.DataTable dt = new System.Data.DataTable();
        static System.Data.DataTable d0 = new System.Data.DataTable();
        static System.Data.DataTable d1 = new System.Data.DataTable();
        static System.Data.DataTable d2 = new System.Data.DataTable();
        static System.Data.DataTable d3 = new System.Data.DataTable();
        static System.Data.DataTable d4 = new System.Data.DataTable();
        static System.Data.DataTable cname = new System.Data.DataTable();
        static System.Data.DataTable sheetname = new System.Data.DataTable();
        public static string CONSTRING
        {
            get {return constring; }
            set { constring = value; }

        }
        public Sql(string Sqlconndetails)
        {CONSTRING= Sqlconndetails;

        }
        /// <summary>
        /// Insert into SQL Table
        /// enter Column Names for all columns Just use * 
        /// use 'value0','value1','value2' format for values string I/p
        /// </summary>
        /// <param name="table_name"></param>
        /// <param name="colmun_names"></param>
        /// <param name="values"></param>
        static public void Insertinsqltable(string table_name, string colmun_names, string values)
        {
            
            con = new SqlConnection(CONSTRING);
            con.Open();

            string tp = "insert into " + table_name + "(" + colmun_names + ")= ('" + values + "')";
            SqlCommand cmd = new SqlCommand(tp,con);
            //new SqlCommand("insert into LoginLog(Username,LoginTime)values('" + usernametxt.Text + "','" + time.Text
            cmd.ExecuteNonQuery();
            con.Close();
        }
        /// <summary>
        /// Updates values in SQL table based on only one comaparison
        /// </summary>
        /// <param name="table_name"></param>
        /// <param name="updatecolmun_names"></param>
        /// <param name="comparisonColumn1Name"></param>
        /// <param name="comparisonColumn1Value"></param>
        /// <param name="newvalues"></param>
        static public void updatesql(string table_name, string updatecolmun_names, string newvalues,string comparisonColumn1Name, string comparisonColumn1Value )
        {

            con = new SqlConnection(CONSTRING);
            con.Open();

            string tp = "update " + table_name + "  set " + updatecolmun_names + "='" + newvalues + "'  where"+ comparisonColumn1Name + " ='" + comparisonColumn1Value + "'";
            SqlCommand cmd = new SqlCommand(tp, con);
          
        }
        /// <summary>
         /// Updates values in SQL table based on only two comaparison
         /// </summary>
         /// <param name="table_name"></param>
         /// <param name="updatecolmun_names"></param>
         /// <param name="comparisonColumn1Name"></param>
         /// <param name="comparisonColumn1Value"></param>
         /// <param name="newvalues"></param>
        static public void updatesql(string table_name, string updatecolmun_names, string newvalues, string comparisonColumn1Name, string comparisonColumn1Value, string comparisonColumn2Name, string comparisonColumn2Value)
        {

            con = new SqlConnection(CONSTRING);
            con.Open();

            string tp = "update " + table_name + "  set " + updatecolmun_names + "='" + newvalues + "'  where" + comparisonColumn1Name + " ='" + comparisonColumn1Value + "' and" +comparisonColumn2Name+"='"+comparisonColumn2Value+"'";
            SqlCommand cmd = new SqlCommand(tp, con);
            con.Close();    
        }
        
       /// <summary>
       /// Delets entire row from table comparing 2 columns values
       /// </summary>
       /// <param name="table_name"></param>
       /// <param name="DeleteColumn1Name"></param>
       /// <param name="DeleteColumn1Value"></param>
       /// <param name="DeleteColumn2Name"></param>
       /// <param name="DeleteColumn2Value"></param>
        static public void deletefrmSql(string table_name,  string DeleteColumn1Name, string DeleteColumn1Value, string DeleteColumn2Name, string DeleteColumn2Value)
        {
            con = new SqlConnection(CONSTRING);
            con.Open();
            string tp = "delete from " + table_name + " where"+ DeleteColumn1Name +"= '" + DeleteColumn1Value + "' and " + DeleteColumn2Name + " ='" + DeleteColumn2Value + "'";
            SqlCommand cmd = new SqlCommand(tp, con);
            
            cmd.ExecuteNonQuery();
            con.Close();
        }
        /// <summary>
        /// Delets entire row from table comparing 1 columns values
        /// </summary>
        /// <param name="table_name"></param>
        /// <param name="DeleteColumn1Name"></param>
        /// <param name="DeleteColumn1Value"></param>
        static public void deletefrmSql(string table_name, string DeleteColumn1Name, string DeleteColumn1Value)
        {
            con = new SqlConnection(CONSTRING);
            con.Open();
            string tp = "delete from " + table_name + " where" + DeleteColumn1Name + "= '" + DeleteColumn1Value + "'";
            SqlCommand cmd = new SqlCommand(tp, con);
            
            cmd.ExecuteNonQuery();
            con.Close();
        }
        /// <summary>
        /// gets Sql table in datagridview in DESCINDING manner
        /// </summary>
        /// <param name="sqltable"></param>
        /// <param name="ColumnName"></param>
        /// <param name="dataGridVeiw"></param>
        /// <param name="orderbycolmn"></param>
        static public void sqlltodatatableDescending(string sqltable, string ColumnName, DataGridView dataGridVeiw, string orderbycolmn)
        {
            con = new SqlConnection(CONSTRING);
            con.Open();

            dt = new System.Data.DataTable();
            string temp = "select "+ColumnName+" from " + sqltable + "  order by " + orderbycolmn + " desc";
            adapter = new SqlDataAdapter(temp, con);
            adapter.Fill(dt);
            dataGridVeiw.DataSource = dt;
            con.Close();

        }
        /// <summary>
        /// gets Sql table in datagridview in Aescending manner
        /// </summary>
        /// <param name="sqltable"></param>
        /// <param name="ColumnName"></param>
        /// <param name="dataGridVeiw"></param>
        /// <param name="orderbycolmn"></param>
        static public void sqlltodatatableAescending(string sqltable,string ColumnName, DataGridView dataGridVeiw, string orderbycolmn)
        {
            con = new SqlConnection(CONSTRING);
            con.Open();

            dt = new System.Data.DataTable();
            string temp = "select "+ ColumnName+" from " + sqltable + "  order by " + orderbycolmn + " asc";
            adapter = new SqlDataAdapter(temp, Sql.con);
            adapter.Fill(dt);
            dataGridVeiw.DataSource = dt;
            con.Close();

        }/// <summary>
        /// Gets rows count of datatable comparing 2columns aminly usefull for logins
        /// </summary>
        /// <param name="table_name"></param>
        /// <param name="comparisonColumn1Name"></param>
        /// <param name="comparisonColumn1Value"></param>
        /// <param name="comparisonColumn2Name"></param>
        /// <param name="comparisonColumn2Value"></param>
        /// <returns></returns>
        static public int sqlgetrowcount(string table_name, string comparisonColumn1Name, string comparisonColumn1Value, string comparisonColumn2Name, string comparisonColumn2Value)
        {
            con = new SqlConnection(CONSTRING);
            con.Open();

            dt = new System.Data.DataTable();
            string temp = "select * from " + table_name + "  where"+ comparisonColumn1Name +"= '" + comparisonColumn1Value + "' and "+ comparisonColumn2Name+"='"+ comparisonColumn2Value+"'";
            adapter = new SqlDataAdapter(temp, con);
            adapter.Fill(dt);
            int count = dt.Rows.Count;
           con.Close();
            return count;

        }
        public static void TruncatefromSQL(string table_name)
        {
            con.Open();
            string temp = $@"Truncate table {table_name}";
            SqlCommand cmd = new SqlCommand(temp, con);
            cmd.ExecuteNonQuery();
            con.Close();
        }
        //static string Fid;
        //static int Fminus;

        //// public static string b;
        //static Microsoft.Office.Interop.Excel.Application Excell;
        //static Microsoft.Office.Interop.Excel.Application Excell2;
        //static Workbook wb;
        //static Worksheet ws;
        //public static void excel(IProgress<int> p)
        //{
        //    try
        //    {

        //        p.Report(4);
        //        Excell = new Microsoft.Office.Interop.Excel.Application();
        //        Excell2 = new Microsoft.Office.Interop.Excel.Application();
        //        // wb = Excell.Workbooks.Open("C:\\Users\\admin\\Desktop\\temp.xlsx");
        //        // wb= new Workbook();
        //        wb = Excell.Workbooks.Add(XlSheetType.xlWorksheet);
        //        constring = Form1.Sqldetails;
        //        con = new SqlConnection(constring);

        //        Sql.con.Open();
        //        // string cmd0 = "select [ReportName] from Reports";
        //        string cmd0 = "select * from ReportsColumns";
        //        adapter = new SqlDataAdapter(cmd0, Sql.con);
        //        adapter.Fill(sheetname);
        //        Sql.con.Close();
        //        int g = -1;
        //        List<string> tm0 = new List<string>();
        //        //string[] tm0 = new string[sheetname.Rows.Count];
        //        List<string> tm1 = new List<string>();
        //        List<string> tm2 = new List<string>();
        //        List<string> tm3 = new List<string>();
        //        List<string> tm4 = new List<string>();
        //        List<string> tm5 = new List<string>();
        //        p.Report(10);
        //        foreach (DataRow row in sheetname.Rows)
        //        {
        //            int rr = -1;
        //            g++;
        //            tm0.Add(row[0].ToString());

        //            foreach (DataColumn col in sheetname.Columns)
        //            {
        //                rr++;
        //                switch (g)
        //                {

        //                    case 0:
        //                        {
        //                            tm1.Add(row[col].ToString());

        //                        }
        //                        break;
        //                    case 1:
        //                        {
        //                            tm2.Add(row[col].ToString());
        //                        }
        //                        break;
        //                    case 2:
        //                        {
        //                            tm3.Add(row[col].ToString());
        //                        }
        //                        break;
        //                    case 3:
        //                        {
        //                            tm4.Add(row[col].ToString());
        //                        }
        //                        break;
        //                    case 4:
        //                        {
        //                            tm5.Add(row[col].ToString());
        //                        }
        //                        break;
        //                }



        //            }
        //        }
        //        p.Report(15);

        //        int number = -1;
        //        for (int f = 0; f < tm0.Count; f++)
        //        {
        //            number++;

        //            Sql.con.Open();
        //            string cmd = $@"Select * from Reports where [ReportName] ='{tm0[f].ToString()}'";
        //            // SqlCommand c= new SqlCommand(cmd);
        //            System.Data.DataTable dtt = new System.Data.DataTable();
        //            adapter = new SqlDataAdapter(cmd, Sql.con);
        //            adapter.Fill(dtt);
        //            Sql.con.Close();
        //            int j = -1;
        //            string[] tm = new string[dtt.Columns.Count];
        //            foreach (DataRow dr in dtt.Rows)
        //            {
        //                try
        //                {
        //                    foreach (DataColumn dataColumn in dtt.Columns)
        //                    {
        //                        j++;
        //                        tm[j] = dr[dataColumn].ToString();
        //                    }
        //                }
        //                catch (Exception ex)
        //                {
        //                    MessageBox.Show(ex.Message);
        //                }

        //            }

        //            string b = "";
        //            int k = 0;
        //            for (int i = 1; i < tm.Length; i++)
        //            {

        //                if (tm[i] != "")
        //                {
        //                    string s = tm[i];
        //                    if (i == 1)
        //                    {
        //                        b = b + "Index" + s;
        //                    }
        //                    else
        //                    {
        //                        b = b + ", Index" + s;
        //                    }

        //                }

        //            }
        //            string tempcmd = "";
        //            if (Form1.Timerange == true)
        //            {
        //                if (Form1.timefrm != null && Form1.timeto != null)
        //                {
        //                    tempcmd = $@"Select [DateTime],{b} from [DataTable] Where  [DateTime] between '{Form1.date} {Form1.timefrm}' and '{Form1.date} {Form1.timeto}'  order by [DateTime] desc ";
        //                }
        //                else
        //                {
        //                    MessageBox.Show("Plz choose Time from and To");
        //                    break;
        //                }

        //            }
        //            else
        //            {
        //                tempcmd = $@"Select [DateTime],{b} from [DataTable] Where CONVERT(varchar(50), [DateTime],126) like '{Form1.date}%' order by [DateTime] desc ";

        //            }

        //            Sql.con.Open();

        //            SqlCommand c2 = new SqlCommand(tempcmd);
        //            adapter = new SqlDataAdapter(tempcmd, Sql.con);
        //            DataTable dt2 = new DataTable();

        //            adapter.Fill(dt2);

        //            Sql.con.Close();
        //            switch (number)
        //            {

        //                case 0:
        //                    {
        //                        p.Report(20);

        //                    }
        //                    break;
        //                case 1:
        //                    {
        //                        p.Report(25);
        //                    }
        //                    break;
        //                case 2:
        //                    {
        //                        p.Report(30);
        //                    }
        //                    break;
        //                case 3:
        //                    {
        //                        p.Report(35);
        //                    }
        //                    break;
        //                case 4:
        //                    {
        //                        p.Report(40);
        //                    }
        //                    break;
        //            }
        //            ws = wb.Sheets.Add();

        //            //ws = wb.Sheets[number + 1];
        //            int r = 1;


        //            foreach (DataRow dataRow in dt2.Rows)
        //            {
        //                r++;
        //                int c = 1;

        //                foreach (DataColumn cc in dt2.Columns)
        //                {
        //                    c++;
        //                    ws.Cells[r, c] = dataRow[cc].ToString();
        //                    ws.Cells[r, 2].NumberFormat = "yyyy-MM-dd HH:mm:ss.000";
        //                }
        //            }
        //            int cuul = 2;

        //            ws.Cells[1, 2] = "DateTime";
        //            ws.Cells[1, 2].Font.Bold = true;
        //            ws.Cells[1, 2].Font.Size = 14;
        //            switch (number)
        //            {

        //                case 0:
        //                    {
        //                        p.Report(47);
        //                        for (int i = 1; i < tm1.Count; i++)
        //                        {
        //                            cuul++;
        //                            ws.Cells[1, cuul] = tm1[i];
        //                            ws.Cells[1, cuul].Font.Bold = true;
        //                            ws.Cells[1, cuul].Font.Size = 14;
        //                            ws.Cells[1, cuul].WrapText = true;
        //                        }
        //                        ws.Name = tm0[f].ToString();
        //                    }
        //                    break;
        //                case 1:
        //                    {
        //                        p.Report(50);
        //                        for (int i = 1; i < tm2.Count; i++)
        //                        {
        //                            cuul++;
        //                            ws.Cells[1, cuul] = tm2[i];
        //                            ws.Cells[1, cuul].Font.Bold = true;
        //                            ws.Cells[1, cuul].Font.Size = 14;
        //                            ws.Cells[1, cuul].WrapText = true;
        //                        }
        //                        ws.Name = tm0[f].ToString();
        //                    }
        //                    break;
        //                case 2:
        //                    {
        //                        p.Report(65);
        //                        for (int i = 1; i < tm3.Count; i++)
        //                        {
        //                            cuul++;
        //                            ws.Cells[1, cuul] = tm3[i];
        //                            ws.Cells[1, cuul].Font.Bold = true;
        //                            ws.Cells[1, cuul].Font.Size = 14;
        //                            ws.Cells[1, cuul].WrapText = true;
        //                        }
        //                        ws.Name = tm0[f].ToString();
        //                    }
        //                    break;
        //                case 3:
        //                    {
        //                        p.Report(75);
        //                        for (int i = 1; i < tm4.Count; i++)
        //                        {
        //                            cuul++;
        //                            ws.Cells[1, cuul] = tm4[i];
        //                            ws.Cells[1, cuul].Font.Bold = true;
        //                            ws.Cells[1, cuul].Font.Size = 14;
        //                            ws.Cells[1, cuul].WrapText = true;
        //                        }
        //                        ws.Name = tm0[f].ToString();
        //                    }
        //                    break;
        //                case 4:
        //                    {
        //                        p.Report(100);
        //                        for (int i = 1; i < tm5.Count; i++)
        //                        {
        //                            cuul++;
        //                            ws.Cells[1, cuul] = tm5[i];
        //                            ws.Cells[1, cuul].Font.Bold = true;
        //                            ws.Cells[1, cuul].Font.Size = 14;
        //                            ws.Cells[1, cuul].WrapText = true;
        //                        }
        //                        ws.Name = tm0[f].ToString();
        //                        tm0.Clear();
        //                        tm1.Clear();
        //                        tm2.Clear();
        //                        tm3.Clear();
        //                        tm4.Clear();
        //                        tm5.Clear();
        //                    }
        //                    break;

        //            }



        //        }


        //        //wb.SaveAs(Environment.CurrentDirectory+"Report"+ DateTime.Now.ToString("dd-MM-yy HH-MM")+".xlsx");
        //        //Excell.Visible = true;

        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message);
        //    }
        //    finally
        //    {

        //        var path = "";

        //        if (!System.IO.Directory.Exists("C:\\Report\\ManualReport\\" + DateTime.Now.ToString("MMMM-yyyy")))
        //        {

        //            System.IO.Directory.CreateDirectory("C:\\Report\\ManualReport\\" + DateTime.Now.ToString("MMMM-yyyy"));

        //            System.IO.Directory.CreateDirectory("C:\\Report\\ManualReport\\" + DateTime.Now.ToString("MMMM-yyyy").ToString() + "\\" + DateTime.Now.ToString("dd-MM-yy"));

        //            path = "C:\\Report\\ManualReport\\" + DateTime.Now.ToString("MMMM-yyyy").ToString() + "\\" + DateTime.Now.ToString("dd-MM-yy");

        //        }

        //        else
        //        {
        //            if (!System.IO.Directory.Exists("C:\\Report\\ManualReport\\" + DateTime.Now.ToString("MMMM-yyyy").ToString() + "\\" + DateTime.Now.ToString("dd-MM-yy")))
        //            {
        //                System.IO.Directory.CreateDirectory("C:\\Report\\ManualReport\\" + DateTime.Now.ToString("MMMM-yyyy").ToString() + "\\" + DateTime.Now.ToString("dd-MM-yy"));

        //                // path = System.IO.Directory.CreateDirectory("C:\\Report\\ManualReport\\" + DateTime.Now.ToString("MMMM-yyyy").ToString() + "\\" + DateTime.Now.ToString("dd-MM-yy") + "\\" + DateTime.Now.ToString("dd-MM-yy")).ToString();
        //            }
        //            //path = Environment.CurrentDirectory + "Report" + DateTime.Now.ToString("dd-MM-yy");
        //            path = "C:\\Report\\ManualReport\\" + DateTime.Now.ToString("MMMM-yyyy").ToString() + "\\" + DateTime.Now.ToString("dd-MM-yy");
        //        }
        //        wb.SaveAs(path + "\\" + DateTime.Now.ToString("dd-MM-yy HH-mm-ss") + ".xlsx");
        //        wb.Close();
        //        Excell.Quit();
        //        MessageBox.Show("Your file is ready in " + path + "\\" + DateTime.Now.ToString("dd-MM-yy HH-mm-ss") + ".xlsx");
        //    }

        //}
        //public static void excel(string h)
        //{
        //    try
        //    {
        //        Excell = new Microsoft.Office.Interop.Excel.Application();
        //        // wb = Excell.Workbooks.Open("C:\\Users\\admin\\Desktop\\temp.xlsx");
        //        wb = Excell.Workbooks.Add(XlSheetType.xlExcel4IntlMacroSheet);
        //        constring = Form1.Sqldetails;
        //        con = new SqlConnection(constring);

        //        Sql.con.Open();
        //        // string cmd0 = "select [ReportName] from Reports";
        //        string cmd0 = "select * from ReportsColumns";
        //        adapter = new SqlDataAdapter(cmd0, Sql.con);
        //        adapter.Fill(sheetname);
        //        Sql.con.Close();
        //        int g = -1;
        //        string[] tm0 = new string[sheetname.Rows.Count];
        //        string[] tm1 = new string[sheetname.Columns.Count];
        //        string[] tm2 = new string[sheetname.Columns.Count];
        //        string[] tm3 = new string[sheetname.Columns.Count];
        //        string[] tm4 = new string[sheetname.Columns.Count];
        //        string[] tm5 = new string[sheetname.Columns.Count];
        //        foreach (DataRow row in sheetname.Rows)
        //        {
        //            int rr = -1;
        //            g++;
        //            tm0[g] = row[0].ToString();

        //            foreach (DataColumn col in sheetname.Columns)
        //            {
        //                rr++;
        //                switch (g)
        //                {

        //                    case 0:
        //                        {
        //                            tm1[rr] = row[col].ToString();

        //                        }
        //                        break;
        //                    case 1:
        //                        {
        //                            tm2[rr] = row[col].ToString();
        //                        }
        //                        break;
        //                    case 2:
        //                        {
        //                            tm3[rr] = row[col].ToString();
        //                        }
        //                        break;
        //                    case 3:
        //                        {
        //                            tm4[rr] = row[col].ToString();
        //                        }
        //                        break;
        //                    case 4:
        //                        {
        //                            tm5[rr] = row[col].ToString();
        //                        }
        //                        break;
        //                }



        //            }
        //        }


        //        int number = -1;
        //        for (int f = 0; f < tm0.Length; f++)
        //        {
        //            number++;

        //            Sql.con.Open();
        //            string cmd = $@"Select * from Reports where [ReportName] ='{tm0[f].ToString()}'";
        //            // SqlCommand c= new SqlCommand(cmd);
        //            System.Data.DataTable dtt = new System.Data.DataTable();
        //            adapter = new SqlDataAdapter(cmd, Sql.con);
        //            adapter.Fill(dtt);
        //            Sql.con.Close();
        //            int j = -1;
        //            string[] tm = new string[dtt.Columns.Count];
        //            foreach (DataRow dr in dtt.Rows)
        //            {
        //                try
        //                {
        //                    foreach (DataColumn dataColumn in dtt.Columns)
        //                    {
        //                        j++;
        //                        tm[j] = dr[dataColumn].ToString();
        //                    }
        //                }
        //                catch (Exception ex)
        //                {
        //                    MessageBox.Show(ex.Message);
        //                }

        //            }

        //            string b = "";
        //            int k = 0;
        //            for (int i = 1; i < tm.Length; i++)
        //            {
        //                if (tm[i].ToString() != "")
        //                {
        //                    string s = tm[i];
        //                    if (i == 1)
        //                    {
        //                        b = b + "Index" + s;
        //                    }
        //                    else
        //                    {
        //                        b = b + ", Index" + s;
        //                    }

        //                }

        //            }
        //            string tempcmd = $@"Select [DateTime],{b} from [DataTable] Where CONVERT(varchar(50), [DateTime],126) like '{h}%' order by [DateTime] desc ";
        //            Sql.con.Open();

        //            SqlCommand c2 = new SqlCommand(tempcmd);
        //            adapter = new SqlDataAdapter(tempcmd, Sql.con);
        //            DataTable dt2 = new DataTable();

        //            adapter.Fill(dt2);
        //            Sql.con.Close();
        //            ws = wb.Sheets.Add();
        //            //ws = wb.Sheets[number + 1];
        //            int r = 1;


        //            foreach (DataRow dataRow in dt2.Rows)
        //            {
        //                r++;
        //                int c = 1;
        //                foreach (DataColumn cc in dt2.Columns)
        //                {
        //                    c++;
        //                    ws.Cells[r, c] = dataRow[cc].ToString();
        //                    ws.Cells[r, 2].NumberFormat = "yyyy-MM-dd HH:mm:ss.000";
        //                }
        //            }
        //            int cuul = 2;
        //            ws.Cells[1, 2] = "DateTime";
        //            switch (number)
        //            {

        //                case 0:
        //                    {
        //                        for (int i = 1; i < tm1.Length; i++)
        //                        {
        //                            cuul++;
        //                            ws.Cells[1, cuul] = tm1[i].ToString();
        //                        }
        //                    }
        //                    break;
        //                case 1:
        //                    {
        //                        for (int i = 1; i < tm2.Length; i++)
        //                        {
        //                            cuul++;
        //                            ws.Cells[1, cuul] = tm2[i].ToString();
        //                        }
        //                    }
        //                    break;
        //                case 2:
        //                    {
        //                        for (int i = 1; i < tm3.Length; i++)
        //                        {
        //                            cuul++;
        //                            ws.Cells[1, cuul] = tm3[i].ToString();
        //                        }
        //                    }
        //                    break;
        //                case 3:
        //                    {
        //                        for (int i = 1; i < tm4.Length; i++)
        //                        {
        //                            cuul++;
        //                            ws.Cells[1, cuul] = tm4[i].ToString();
        //                        }
        //                    }
        //                    break;
        //                case 4:
        //                    {
        //                        for (int i = 1; i < tm5.Length; i++)
        //                        {
        //                            cuul++;
        //                            ws.Cells[1, cuul] = tm5[i].ToString();
        //                        }
        //                    }
        //                    break;

        //            }


        //        }
        //        wb.SaveAs($@"C:\Users\admin\Desktop\{DateTime.Now.ToString("dd:MM:yy")}.xlsx");
        //        Excell.Visible = true;

        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message);
        //    }
        //    finally
        //    {
        //        wb.SaveAs($@"C:\Users\admin\Desktop\{DateTime.Now.ToString("dd:MM:yy")}.xlsx");
        //        Excell.Visible = true;

        //    }

        //}
        //public static void csv(IProgress<int> p)
        //{
        //    string exceptionfile = $@"D:\ReportGenration Software\Exception{DateTime.Now.ToString("dd-MM-yyy HH-mm-ss")}.TEJ";
        //    try
        //    {

        //        p.Report(4);
        //        Excell = new Microsoft.Office.Interop.Excel.Application();
        //        Excell2 = new Microsoft.Office.Interop.Excel.Application();

        //        // wb = Excell.Workbooks.Add(XlSheetType.xlWorksheet);
        //        constring = Form1.Sqldetails;
        //        con = new SqlConnection(constring);
        //        string filename = "";
        //        Sql.con.Open();
        //        // string cmd0 = "select [ReportName] from Reports";
        //        string cmd0 = "select * from ReportsColumns";
        //        adapter = new SqlDataAdapter(cmd0, Sql.con);
        //        adapter.Fill(sheetname);
        //        Sql.con.Close();
        //        int g = -1;
        //        List<string> tm0 = new List<string>();
        //        //string[] tm0 = new string[sheetname.Rows.Count];
        //        List<string> tm1 = new List<string>();
        //        List<string> tm2 = new List<string>();
        //        List<string> tm3 = new List<string>();
        //        List<string> tm4 = new List<string>();
        //        List<string> tm5 = new List<string>();
        //        List<string> tm6 = new List<string>();
        //        p.Report(10);
        //        foreach (DataRow row in sheetname.Rows)
        //        {
        //            int rr = -1;
        //            g++;
        //            tm0.Add(row[0].ToString());

        //            foreach (DataColumn col in sheetname.Columns)
        //            {
        //                rr++;
        //                switch (g)
        //                {

        //                    case 0:
        //                        {
        //                            tm1.Add(row[col].ToString());

        //                        }
        //                        break;
        //                    case 1:
        //                        {
        //                            tm2.Add(row[col].ToString());
        //                        }
        //                        break;
        //                    case 2:
        //                        {
        //                            tm3.Add(row[col].ToString());
        //                        }
        //                        break;
        //                    case 3:
        //                        {
        //                            tm4.Add(row[col].ToString());
        //                        }
        //                        break;
        //                    case 4:
        //                        {
        //                            tm5.Add(row[col].ToString());
        //                        }
        //                        break;
        //                }



        //            }
        //        }
        //        p.Report(15);

        //        int number = -1;
        //        for (int f = 0; f < tm0.Count; f++)
        //        {
        //            number++;

        //            Sql.con.Open();
        //            string cmd = $@"Select * from Reports where [ReportName] ='{tm0[f].ToString()}'";
        //            // SqlCommand c= new SqlCommand(cmd);
        //            System.Data.DataTable dtt = new System.Data.DataTable();
        //            adapter = new SqlDataAdapter(cmd, Sql.con);
        //            adapter.Fill(dtt);
        //            Sql.con.Close();
        //            int j = -1;
        //            string[] tm = new string[dtt.Columns.Count];
        //            foreach (DataRow dr in dtt.Rows)
        //            {
        //                try
        //                {
        //                    foreach (DataColumn dataColumn in dtt.Columns)
        //                    {
        //                        j++;
        //                        tm[j] = dr[dataColumn].ToString();
        //                    }
        //                }
        //                catch (Exception ex)
        //                {
        //                    using (System.IO.StreamWriter file = new System.IO.StreamWriter(exceptionfile, append: true))
        //                    {
        //                        file.WriteLine(ex.Source);
        //                        file.WriteLine(ex);
        //                    }
        //                    MessageBox.Show(ex.Message);
        //                }

        //            }

        //            string b = "";
        //            int k = 0;
        //            for (int i = 1; i < tm.Length; i++)
        //            {

        //                if (tm[i] != "")
        //                {
        //                    string s = tm[i];
        //                    if (i == 1)
        //                    {
        //                        b = b + "Index" + s;
        //                    }
        //                    else
        //                    {
        //                        b = b + ", Index" + s;
        //                    }

        //                }

        //            }
        //            string tempcmd = "";
        //            if (Form1.Timerange == true)
        //            {
        //                if (Form1.timefrm != null && Form1.timeto != null)
        //                {
        //                    tempcmd = $@"Select [DateTime],{b} from [DataTable] Where  [DateTime] between '{Form1.date} {Form1.timefrm}' and '{Form1.date} {Form1.timeto}'  order by [DateTime]  ";
        //                }
        //                else
        //                {
        //                    MessageBox.Show("Plz choose Time from and To");
        //                    break;
        //                }

        //            }
        //            else
        //            {
        //                tempcmd = $@"Select [DateTime],{b} from [DataTable] Where CONVERT(varchar(50), [DateTime],126) like '{Form1.date}%' order by [DateTime]  ";

        //            }

        //            Sql.con.Open();

        //            SqlCommand c2 = new SqlCommand(tempcmd);
        //            adapter = new SqlDataAdapter(tempcmd, Sql.con);
        //            DataTable dt2 = new DataTable();

        //            adapter.Fill(dt2);

        //            Sql.con.Close();
        //            switch (number)
        //            {

        //                case 0:
        //                    {
        //                        p.Report(20);

        //                    }
        //                    break;
        //                case 1:
        //                    {
        //                        p.Report(25);
        //                    }
        //                    break;
        //                case 2:
        //                    {
        //                        p.Report(30);
        //                    }
        //                    break;
        //                case 3:
        //                    {
        //                        p.Report(35);
        //                    }
        //                    break;
        //                case 4:
        //                    {
        //                        p.Report(40);
        //                    }
        //                    break;
        //            }
        //            /// ws = wb.Sheets.Add();

        //            //ws = wb.Sheets[number + 1];
        //            int r = 1;
        //            string path = "C:\\Report";
        //            string rrr = "";

        //            switch (number)
        //            {
        //                case 0:
        //            {
        //                filename = "Extruder";
        //                        tm6.Add(filename);

        //            }break; case 1:

        //            {
        //                filename = "Temperature";
        //                        tm6.Add(filename);
        //                    }
        //                    break; case 2:

        //            {
        //                filename = "Dancer"; tm6.Add(filename);
        //                    } break; case 3:

        //            { filename = "Conveyor"; tm6.Add(filename); }
        //                    break ; case 4:

        //            { filename = "Product";tm6.Add(filename); } break; 
        //        }
        //            using (System.IO.StreamWriter file = new System.IO.StreamWriter(path + "\\" + $@"{filename}.csv", append: true))
        //            {
        //                //    System.Data.OleDb.OleDbConnection cn = new System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\\ARS\\temp.xlsx;Extended Properties=Excel 12.0 Xml;HDR=YES;");
        //                //cn.Open();
        //                //System.Data.OleDb.OleDbCommand cmnd;
        //                int cuul = 0; string[] dd = new string[100];
        //                string ss = "";
        //                switch (number)
        //                {

        //                    case 0:
        //                        {
        //                            p.Report(47);
        //                            for (int i = 1; i < tm1.Count; i++)
        //                            {
        //                                if (tm1[i] != "")
        //                                {



        //                                    cuul++;
        //                                    if (cuul == 1)
        //                                    {
        //                                        ss ="DateAndTime,"+ tm1[i];
        //                                        rrr = rrr + "" + ss;
        //                                    }
        //                                    else
        //                                    {
        //                                        ss = tm1[i];
        //                                        rrr = rrr + "," + ss;

        //                                    }


        //                                }
        //                            }
        //                            file.WriteLine(rrr); cuul = 0;
        //                            rrr = "";
        //                            // ws.Name = tm0[f].ToString();
        //                        }
        //                        break;
        //                    case 1:
        //                        {
        //                            p.Report(50);
        //                            for (int i = 1; i < tm2.Count; i++)
        //                            {

        //                                if (tm2[i] != "")
        //                                {



        //                                    cuul++;
        //                                    if (cuul == 1)
        //                                    {
        //                                        ss ="DateAndTime,"+ tm2[i];
        //                                        rrr = rrr + "" + ss;
        //                                    }
        //                                    else
        //                                    {
        //                                        ss = tm2[i];
        //                                        rrr = rrr + "," + ss;

        //                                    }


        //                                }
        //                            }
        //                            file.WriteLine(rrr);
        //                            rrr = "";
        //                            // ws.Name = tm0[f].ToString();
        //                        }
        //                        break;
        //                    case 2:
        //                        {

        //                            for (int i = 1; i < tm3.Count; i++)
        //                            {

        //                                if (tm3[i] != "")
        //                                {



        //                                    cuul++;
        //                                    if (cuul == 1)
        //                                    {
        //                                        ss = "DateAndTime,"+tm3[i];
        //                                        rrr = rrr + "" + ss;
        //                                    }
        //                                    else
        //                                    {
        //                                        ss = tm3[i];
        //                                        rrr = rrr + "," + ss;

        //                                    }


        //                                }

        //                            }
        //                            file.WriteLine(rrr); cuul = 0;
        //                            rrr = "";
        //                            // ws.Name = tm0[f].ToString();
        //                        }
        //                        break;
        //                    case 3:
        //                        {
        //                            p.Report(75);
        //                            for (int i = 1; i < tm4.Count; i++)
        //                            {

        //                                if (tm4[i] != "")
        //                                {



        //                                    cuul++;
        //                                    if (cuul == 1)
        //                                    {
        //                                        ss = "DateAndTime,"+tm4[i];
        //                                        rrr = rrr + "" + ss;
        //                                    }
        //                                    else
        //                                    {
        //                                        ss = tm4[i];
        //                                        rrr = rrr + "," + ss;

        //                                    }


        //                                }

        //                            }
        //                            file.WriteLine(rrr); cuul = 0;
        //                            rrr = "";
        //                            // ws.Name = tm0[f].ToString();
        //                        }
        //                        break;
        //                    case 4:
        //                        {
        //                            p.Report(100);
        //                            for (int i = 1; i < tm5.Count; i++)
        //                            {

        //                                if (tm5[i] != "")
        //                                {



        //                                    cuul++;
        //                                    if (cuul == 1)
        //                                    {
        //                                        ss = "DateAndTime," + tm5[i];
        //                                        rrr = rrr + "" + ss;
        //                                    }
        //                                    else
        //                                    {
        //                                        ss = tm5[i];
        //                                        rrr = rrr + "," + ss;

        //                                    }


        //                                }

        //                            }
        //                            file.WriteLine(rrr); cuul = 0;
        //                            rrr = "";
        //                            //  ws.Name = tm0[f].ToString();
        //                            tm0.Clear();
        //                            tm1.Clear();
        //                            tm2.Clear();
        //                            tm3.Clear();
        //                            tm4.Clear();
        //                            tm5.Clear();
        //                        }
        //                        break;

        //                }
        //                foreach (DataRow dataRow in dt2.Rows)
        //                {
        //                    r++;
        //                    int c = 1;
        //                    int H = 0;
        //                    foreach (DataColumn cc in dt2.Columns)
        //                    {
        //                        c++;
        //                        //ws.Cells[r, c] = dataRow[cc].ToString();
        //                        //ws.Cells[r, 2].NumberFormat = "yyyy-MM-dd HH:mm:ss.000";

        //                        H++;
        //                        var a =dataRow[cc.ColumnName];

        //                        if (dataRow[cc].ToString() != "")
        //                        {
        //                            string[] d = new string[100];
        //                            string s = dataRow[cc].ToString();
        //                            if (H == 1)
        //                            {
        //                                rrr = rrr + "" + s;
        //                            }
        //                            else
        //                            {
        //                                rrr = rrr + "," + s;

        //                            }




        //                        }
        //                    }
        //                    file.WriteLine(rrr);
        //                    rrr = "";

        //                }/*cn.Close();*/
        //                try
        //                {

        //                    Excell.DisplayAlerts = false;

        //                    wb = Excell.Workbooks.Open(path + "\\" + $@"{filename}.csv");
        //                    ws = wb.Worksheets[1];
        //                    ws.Columns[1].NumberFormat = "yyyy-MM-dd HH:mm:ss.000";
        //                    for (int i = 1; i < 57;i++) 
        //                    {

        //                        ws.Cells[1, i].Font.Bold = true;
        //                        ws.Cells[1, i].Font.Size = 14;
        //                        ws.Cells[1, i].WrapText = true; 
        //                    }

        //                    wb.SaveAs($@"C:\Report\9{number}.xlsx",Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault,Type.Missing,Type.Missing,false,false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange
        //                        ,Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

        //                    wb.Close();

        //                }
        //                catch (Exception ex)
        //                {
        //                    using (System.IO.StreamWriter file2 = new System.IO.StreamWriter(exceptionfile, append: true))
        //                    {
        //                        file2.WriteLine(ex.Source);
        //                        file2.WriteLine(ex);
        //                    }
        //                    MessageBox.Show(ex.Message);

        //                }
        //                finally
        //                {

        //                    //wb.Close();
        //                    //Excell.Quit();
        //                }

        //            }


        //            //ws.Cells[1, 2] = "DateTime";
        //            //ws.Cells[1, 2].Font.Bold = true;
        //            //ws.Cells[1, 2].Font.Size = 14;




        //        }

        //        try
        //        {


        //            Workbook w1 = Excell.Workbooks.Add("C:\\Report\\90.xlsx");
        //            Workbook w2 = Excell.Workbooks.Add("C:\\Report\\91.xlsx");
        //            Workbook w3 = Excell.Workbooks.Add("C:\\Report\\92.xlsx");
        //            Workbook w4 = Excell.Workbooks.Add("C:\\Report\\93.xlsx");
        //            Workbook w5 = Excell.Workbooks.Add("C:\\Report\\94.xlsx");



        //            for (int i = 2; i <= Excell.Workbooks.Count; i++)
        //            {
        //                for (int j = 1; j <= Excell.Workbooks[i].Worksheets.Count; j++)
        //                {
        //                    Worksheet ws = (Worksheet)Excell.Workbooks[i].Worksheets[j];
        //                    ws.Copy(Excell.Workbooks[1].Worksheets[1]);
        //                    Worksheet wss = Excell.Workbooks[1].Worksheets[1];



        //                }
        //            }
        //            var path = "";

        //            if (!System.IO.Directory.Exists("C:\\Report\\ManualReport\\" + DateTime.Now.ToString("MMMM-yyyy")))
        //            {

        //                System.IO.Directory.CreateDirectory("C:\\Report\\ManualReport\\" + DateTime.Now.ToString("MMMM-yyyy"));

        //                System.IO.Directory.CreateDirectory("C:\\Report\\ManualReport\\" + DateTime.Now.ToString("MMMM-yyyy").ToString() + "\\" + DateTime.Now.ToString("dd-MM-yy"));

        //                path = "C:\\Report\\ManualReport\\" + DateTime.Now.ToString("MMMM-yyyy").ToString() + "\\" + DateTime.Now.ToString("dd-MM-yy");

        //            }

        //            else
        //            {
        //                if (!System.IO.Directory.Exists("C:\\Report\\ManualReport\\" + DateTime.Now.ToString("MMMM-yyyy").ToString() + "\\" + DateTime.Now.ToString("dd-MM-yy")))
        //                {
        //                    System.IO.Directory.CreateDirectory("C:\\Report\\ManualReport\\" + DateTime.Now.ToString("MMMM-yyyy").ToString() + "\\" + DateTime.Now.ToString("dd-MM-yy"));

        //                    // path = System.IO.Directory.CreateDirectory("C:\\Report\\ManualReport\\" + DateTime.Now.ToString("MMMM-yyyy").ToString() + "\\" + DateTime.Now.ToString("dd-MM-yy") + "\\" + DateTime.Now.ToString("dd-MM-yy")).ToString();
        //                }
        //                //path = Environment.CurrentDirectory + "Report" + DateTime.Now.ToString("dd-MM-yy");
        //                path = "C:\\Report\\ManualReport\\" + DateTime.Now.ToString("MMMM-yyyy").ToString() + "\\" + DateTime.Now.ToString("dd-MM-yy");
        //            }
        //            Excell.Workbooks[1].SaveCopyAs(path + "\\" + DateTime.Now.ToString("dd-MM-yy HH-mm-ss") + ".xlsx");
        //            //Excell.Workbooks[1].Close();
        //            for (int i = 0; i < 5; i++)
        //            {
        //                File.Delete($@"C:\Report\9{i}.xlsx");
        //                File.Delete("C:\\Report" + "\\" + $@"{tm6[i]}.csv");
        //            }


        //            w1.Close();
        //            w2.Close();
        //            w3.Close();
        //            w4.Close();
        //            w5.Close();
        //            //Excell.Quit();
        //            MessageBox.Show("Your file is ready in " + path + "\\" + DateTime.Now.ToString("dd-MM-yy HH-mm-ss") + ".xlsx");









        //        }
        //        catch (Exception ex)
        //        {
        //            MessageBox.Show(ex.Message);
        //            using (System.IO.StreamWriter file = new System.IO.StreamWriter(exceptionfile, append: true))
        //            {
        //                file.WriteLine(ex.Source);
        //                file.WriteLine(ex);
        //            }
        //        }


        //    }

        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message);
        //        using (System.IO.StreamWriter file = new System.IO.StreamWriter(exceptionfile, append: true))
        //        {
        //            file.WriteLine(ex.Source);
        //            file.WriteLine(ex);
        //        }
        //    }
        //    finally
        //    {
        //        //wb.Close();
        //        Excell.Application.Quit();


        //    }

        //}
    }
}






