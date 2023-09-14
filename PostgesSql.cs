using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Npgsql;

namespace ReportGenrationJindalAcidCopper
{
    public class PostgesSql
    {
        public static NpgsqlConnection con;
        static NpgsqlDataAdapter adapter = new NpgsqlDataAdapter();
        static System.Data.DataTable dt = new System.Data.DataTable();
        static System.Data.DataTable d0 = new System.Data.DataTable();
        static System.Data.DataTable d1 = new System.Data.DataTable();
        static System.Data.DataTable d2 = new System.Data.DataTable();
        static System.Data.DataTable d3 = new System.Data.DataTable();
        static System.Data.DataTable d4 = new System.Data.DataTable();
        static System.Data.DataTable cname = new System.Data.DataTable();
        static System.Data.DataTable sheetname = new System.Data.DataTable();
        public static string constring;
        public static string CONSTRING
        {
            get { return constring; }
            set { constring = value; }

        }
        public PostgesSql(string Sqlconndetails)
        {
            CONSTRING = Sqlconndetails;

        }
        /// <summary>
        /// Selects the data between specific time frame 
        /// </summary>
        /// <param name="DT"></param>
        /// <param name="table_name"></param>
        /// <param name="column_to_be_selected"></param>
        /// <param name="datetimecolumn"></param>
        /// <param name="datetimefrom"></param>
        /// <param name="datetimeto"></param>
        /// <returns></returns>
        public static DataTable select_between_timeframe(DataTable DT, string table_name, string column_to_be_selected, string datetimecolumn, string datetimefrom, string datetimeto)
        {
            try
            {
                // k++;
                //      SELECT TOP(1) [Id]
                //,[timestamp]
                //,[value]
                //,[el_DeviceSetId]
                //      FROM[galvcon_oc65].[dbo].[el_devArchiveSet] WHERE el_DeviceSetId = 4 and timestamp between convert(smalldatetime,'16-12-2022 13:28:39',103) and convert(smalldatetime,'16-12-2022 16:57:59',103)
                con = new NpgsqlConnection(CONSTRING);
                con.Open();
                string temp = $@"Select  {column_to_be_selected} from  {table_name} WHERE TO_TIMESTAMP(timestamp_col, 'DD-MM-YYYY HH24:MI:SS') BETWEEN
                                             TO_TIMESTAMP('{datetimefrom}', 'DD-MM-YYYY HH24:MI:SS') AND
                                             TO_TIMESTAMP('{datetimeto}', 'DD-MM-YYYY HH24:MI:SS') ORDER BY timestamp_col ASC;";
                //StreamWriter writer = new StreamWriter($@"D:\\text{k}.txt");
                //writer.WriteLine(temp);
                //writer.Close();
                adapter = new NpgsqlDataAdapter(temp, con);
                adapter.Fill(DT);
                con.Close();
                //  return DT;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return DT;
        }
        public static DataTable select_data_fromsql(DataTable DT, string table_name, string column_to_be_selected)
        {
            con = new NpgsqlConnection(CONSTRING);
            con.Open();
            string temp = $@"Select {column_to_be_selected} from public.""{table_name}"";";
            adapter = new NpgsqlDataAdapter(temp, con);
            adapter.Fill(DT);
            con.Close();
            return DT;
        }
        public static DataTable select_between_timeframe2(DataTable DT, string table_name, string column_to_be_selected, string datetimecolumn, string datetimefrom, string datetimeto)
        {
            try
            {
                // k++;
                //      SELECT TOP(1) [Id]
                //,[timestamp]
                //,[value]
                //,[el_DeviceSetId]
                //      FROM[galvcon_oc65].[dbo].[el_devArchiveSet] WHERE el_DeviceSetId = 4 and timestamp between convert(smalldatetime,'16-12-2022 13:28:39',103) and convert(smalldatetime,'16-12-2022 16:57:59',103)
                con = new NpgsqlConnection(CONSTRING);
                con.Open();
                string temp = $@"Select  {column_to_be_selected} from  {table_name} where {datetimecolumn} between'{datetimefrom}' and'{datetimeto}' order by{datetimecolumn} asc;";
                //StreamWriter writer = new StreamWriter($@"D:\\text{k}.txt");
                //writer.WriteLine(temp);
                //writer.Close();
                adapter = new NpgsqlDataAdapter(temp, con);
                adapter.Fill(DT);
                con.Close();
                //  return DT;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return DT;
        }
    }
}
