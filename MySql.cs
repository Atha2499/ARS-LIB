using MySqlConnector;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp3
{
    public class MySql
    {    public static string  constring;
        public static MySqlConnection con;
        public static MySqlDataAdapter adapter;
        public static System.Data.DataTable dt;
        public static string CONSTRING
        {
            get { return constring; }
            set { constring = value; }

        }
        public MySql(string MySqlconndetails)
        {
            CONSTRING = MySqlconndetails;

        }
        /// <summary>
        /// Insert into SQL Table
        /// enter Column Names for all columns Just use * 
        /// use 'value0','value1','value2' format for values string I/p
        /// </summary>
        /// <param name="table_name"></param>
        /// <param name="colmun_names"></param>
        /// <param name="values"></param>
        static public void InsertinMySqltable(string table_name, string colmun_names, string values)
        {

            con = new MySqlConnection(CONSTRING);
            con.Open();

            string tp = "insert into " + table_name + "(" + colmun_names + ")= ('" + values + "')";
            MySqlCommand cmd = new MySqlCommand(tp, con);
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
        static public void updateMySql(string table_name, string updatecolmun_names, string newvalues, string comparisonColumn1Name, string comparisonColumn1Value)
        {

            con = new MySqlConnection(CONSTRING);
            con.Open();

            string tp = "update " + table_name + "  set " + updatecolmun_names + "='" + newvalues + "'  where" + comparisonColumn1Name + " ='" + comparisonColumn1Value + "'";
            MySqlCommand cmd = new MySqlCommand(tp, con);

        }
        /// <summary>
        /// Updates values in SQL table based on only two comaparison
        /// </summary>
        /// <param name="table_name"></param>
        /// <param name="updatecolmun_names"></param>
        /// <param name="comparisonColumn1Name"></param>
        /// <param name="comparisonColumn1Value"></param>
        /// <param name="newvalues"></param>
        static public void updateMySql(string table_name, string updatecolmun_names, string newvalues, string comparisonColumn1Name, string comparisonColumn1Value, string comparisonColumn2Name, string comparisonColumn2Value)
        {

            con = new MySqlConnection(CONSTRING);
            con.Open();

            string tp = "update " + table_name + "  set " + updatecolmun_names + "='" + newvalues + "'  where" + comparisonColumn1Name + " ='" + comparisonColumn1Value + "' and" + comparisonColumn2Name + "='" + comparisonColumn2Value + "'";
            MySqlCommand cmd = new MySqlCommand(tp, con);
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
        static public void deletefrmMySql(string table_name, string DeleteColumn1Name, string DeleteColumn1Value, string DeleteColumn2Name, string DeleteColumn2Value)
        {
            con = new MySqlConnection(CONSTRING);
            con.Open();
            string tp = "delete from " + table_name + " where" + DeleteColumn1Name + "= '" + DeleteColumn1Value + "' and " + DeleteColumn2Name + " ='" + DeleteColumn2Value + "'";
            MySqlCommand cmd = new MySqlCommand(tp, con);

            cmd.ExecuteNonQuery();
            con.Close();
        }
        /// <summary>
        /// Delets entire row from table comparing 1 columns values
        /// </summary>
        /// <param name="table_name"></param>
        /// <param name="DeleteColumn1Name"></param>
        /// <param name="DeleteColumn1Value"></param>
        static public void deletefrmMySql(string table_name, string DeleteColumn1Name, string DeleteColumn1Value)
        {
            con = new MySqlConnection(CONSTRING);
            con.Open();
            string tp = "delete from " + table_name + " where" + DeleteColumn1Name + "= '" + DeleteColumn1Value + "'";
            MySqlCommand cmd = new MySqlCommand(tp, con);

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
        static public void MySqltoDataGridVeiwDescending(string sqltable, string ColumnName, DataGridView dataGridVeiw, string orderbycolmn)
        {
            con = new MySqlConnection(CONSTRING);
            con.Open();

            dt = new System.Data.DataTable();
            string temp = "select " + ColumnName + " from " + sqltable + "  order by " + orderbycolmn + " desc";
            adapter = new MySqlDataAdapter(temp, con);
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
        static public void MySqltoDataGridVeiwAescending(string sqltable, string ColumnName, DataGridView dataGridVeiw, string orderbycolmn)
        {
            con = new MySqlConnection(CONSTRING);
            con.Open();

            dt = new System.Data.DataTable();
            string temp = "select " + ColumnName + " from " + sqltable + "  order by " + orderbycolmn + " asc";
            adapter = new MySqlDataAdapter(temp, con);
            adapter.Fill(dt);
            dataGridVeiw.DataSource = dt;
            con.Close();

        }
        /// <summary>
         /// Gets rows count of datatable comparing 2columns aminly usefull for logins
         /// </summary>
         /// <param name="table_name"></param>
         /// <param name="comparisonColumn1Name"></param>
         /// <param name="comparisonColumn1Value"></param>
         /// <param name="comparisonColumn2Name"></param>
         /// <param name="comparisonColumn2Value"></param>
         /// <returns></returns>
        static public int MySqlgetrowcount(string table_name, string comparisonColumn1Name, string comparisonColumn1Value, string comparisonColumn2Name, string comparisonColumn2Value)
        {
            con = new MySqlConnection(CONSTRING);
            con.Open();

            dt = new System.Data.DataTable();
            string temp = "select * from " + table_name + "  where" + comparisonColumn1Name + "= '" + comparisonColumn1Value + "' and " + comparisonColumn2Name + "='" + comparisonColumn2Value + "'";
            adapter = new MySqlDataAdapter(temp, con);
            adapter.Fill(dt);
            int count = dt.Rows.Count;
            con.Close();
            return count;

        }
        /// <summary>
        /// Deletes all the enteries from Sql Table
        /// </summary>
        /// <param name="table_name"></param>
        public static void TruncatefromMySQL(string table_name)
        {
            con.Open();
            string temp=$@"Truncate table {table_name}";
            MySqlCommand cmd = new MySqlCommand(temp, con);
            cmd.ExecuteNonQuery();
            con.Close();
        }

    }
}
