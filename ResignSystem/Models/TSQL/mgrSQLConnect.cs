using System;
using System.Data;
using System.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;

namespace Import_Freight_BOI.Models.TSQL
{
    public class mgrSQLConnect
    {
        private readonly IConfiguration configuration;
        private DataSet ds = new DataSet();

        private string strSQL = "";

        public mgrSQLConnect(IConfiguration configuration)
        {
            this.configuration = configuration;
        }
        
        public DataTable GetDatatables(string Sql)
        {
            var constr = configuration.GetConnectionString("CONN");
            string consString = "Data Source=RTHSRV14;Initial Catalog=ImportExportDB;Persist Security Info=True;User ID=sa;Password=pwpolicy;";
            //var conGFDReport = configuration.GetConnectionString("Con_GFDReport");
            var dt = new DataTable();


            try
            {
                var query = Sql;

                var ds = new DataSet();
                using (var con = new SqlConnection(consString))
                {
                    using (var cmd = new SqlCommand())
                    {
                        cmd.Connection = con;
                        con.Open();
                        var adpterdata = new SqlDataAdapter();
                        adpterdata.SelectCommand = new SqlCommand(query, con);
                        adpterdata.Fill(dt);
                        con.Close();
                        return dt;
                    }
                }
            }
            catch (Exception e)
            {
                var dsa = e;
                return dt;
            }
        }


        public DataTable GetDatatables_SrvBCP(string Sql)
        {
            var constr = configuration.GetConnectionString("CONBCP");
            string consString = "Data Source=RTHBCP;Initial Catalog=MCR;Persist Security Info=True;User ID=sa;Password=pwpolicy;";
            //var conGFDReport = configuration.GetConnectionString("Con_GFDReport");
            var dt = new DataTable();


            try
            {
                var query = Sql;

                var ds = new DataSet();
                using (var con = new SqlConnection(consString))
                {
                    using (var cmd = new SqlCommand())
                    {
                        cmd.Connection = con;
                        con.Open();
                        var adpterdata = new SqlDataAdapter();
                        adpterdata.SelectCommand = new SqlCommand(query, con);
                        adpterdata.Fill(dt);
                        con.Close();
                        return dt;
                    }
                }
            }
            catch (Exception e)
            {
                var dsa = e;
                return dt;
            }
        }
        public string Get_RTHSRVM03(string Sql)
        {
            //string strConString = @ "Data Source=WELCOME-PC\SQLSERVER2008;Initial Catalog=MyDB;Integrated Security=True";
            var Con_RTHSRVM03 = configuration.GetConnectionString("Con_RTHSRVM03");
            var dt = new DataTable();


            try
            {
                var query = Sql;

                var ds = new DataSet();
                using (var con = new SqlConnection(Con_RTHSRVM03))
                {
                    using (var cmd = new SqlCommand())
                    {
                        cmd.Connection = con;
                        con.Open();
                        var adpterdata = new SqlDataAdapter();
                        adpterdata.SelectCommand = new SqlCommand(query, con);
                        adpterdata.Fill(dt);
                        con.Close();
                        if(dt.Rows.Count >= 1)
                        {
                            return "true";
                        }
                        else
                        {
                            return "false";
                        }
                        
                    }
                }
            }
            catch (Exception e)
            {
                var dsa = e.ToString();
                return dsa;
            }
        }


        public string Get_Con_RTHSRVMCR01(string Sql)
        {
            //string strConString = @ "Data Source=WELCOME-PC\SQLSERVER2008;Initial Catalog=MyDB;Integrated Security=True";
            var Con_RTHSRVMCR01 = configuration.GetConnectionString("Con_RTHSRVMCR01");
            var dt = new DataTable();


            try
            {
                var query = Sql;

                var ds = new DataSet();
                using (var con = new SqlConnection(Con_RTHSRVMCR01))
                {
                    using (var cmd = new SqlCommand())
                    {
                        cmd.Connection = con;
                        con.Open();
                        var adpterdata = new SqlDataAdapter();
                        adpterdata.SelectCommand = new SqlCommand(query, con);
                        adpterdata.Fill(dt);
                        con.Close();
                        if (dt.Rows.Count >= 1)
                        {
                            return "true";
                        }
                        else
                        {
                            return "false";
                        }

                    }
                }
            }
            catch (Exception e)
            {
                var dsa = e.ToString();
                return dsa;
            }
        }
        
             public string Get_RTHSRVOPM02(string Sql)
        {
            //string strConString = @ "Data Source=WELCOME-PC\SQLSERVER2008;Initial Catalog=MyDB;Integrated Security=True";
            var Con_RTHSRVOPM02 = configuration.GetConnectionString("Con_RTHSRVOPM02");
            var dt = new DataTable();


            try
            {
                var query = Sql;

                var ds = new DataSet();
                using (var con = new SqlConnection(Con_RTHSRVOPM02))
                {
                    using (var cmd = new SqlCommand())
                    {
                        cmd.Connection = con;
                        con.Open();
                        var adpterdata = new SqlDataAdapter();
                        adpterdata.SelectCommand = new SqlCommand(query, con);
                        adpterdata.Fill(dt);
                        con.Close();
                        if (dt.Rows.Count >= 1)
                        {
                            return "true";
                        }
                        else
                        {
                            return "false";
                        }

                    }
                }
            }
            catch (Exception e)
            {
                var dsa = e.ToString();
                return dsa;
            }
        }
        public string Get_RTHSRVOPM01(string Sql)
        {
            //string strConString = @ "Data Source=WELCOME-PC\SQLSERVER2008;Initial Catalog=MyDB;Integrated Security=True";
            var Con_RTHSRVOPM01 = configuration.GetConnectionString("Con_RTHSRVOPM01");
            var dt = new DataTable();


            try
            {
                var query = Sql;

                var ds = new DataSet();
                using (var con = new SqlConnection(Con_RTHSRVOPM01))
                {
                    using (var cmd = new SqlCommand())
                    {
                        cmd.Connection = con;
                        con.Open();
                        var adpterdata = new SqlDataAdapter();
                        adpterdata.SelectCommand = new SqlCommand(query, con);
                        adpterdata.Fill(dt);
                        con.Close();
                        if (dt.Rows.Count >= 1)
                        {
                            return "true";
                        }
                        else
                        {
                            return "false";
                        }

                    }
                }
            }
            catch (Exception e)
            {
                var dsa = e.ToString();
                return dsa;
            }
        }
          public string Get_RTHSRV07(string Sql)
        {
            //string strConString = @ "Data Source=WELCOME-PC\SQLSERVER2008;Initial Catalog=MyDB;Integrated Security=True";
            var Con_RTHSRV07 = configuration.GetConnectionString("Con_RTHSRV07");
            var dt = new DataTable();


            try
            {
                var query = Sql;

                var ds = new DataSet();
                using (var con = new SqlConnection(Con_RTHSRV07))
                {
                    using (var cmd = new SqlCommand())
                    {
                        cmd.Connection = con;
                        con.Open();
                        var adpterdata = new SqlDataAdapter();
                        adpterdata.SelectCommand = new SqlCommand(query, con);
                        adpterdata.Fill(dt);
                        con.Close();
                        if (dt.Rows.Count >= 1)
                        {
                            return "true";
                        }
                        else
                        {
                            return "false";
                        }

                    }
                }
            }
            catch (Exception e)
            {
                var dsa = e.ToString();
                return dsa;
            }
        }
        public string Get_RTHSRV11(string Sql)
        {
            //string strConString = @ "Data Source=WELCOME-PC\SQLSERVER2008;Initial Catalog=MyDB;Integrated Security=True";
            var Con_RTHSRV11 = configuration.GetConnectionString("Con_RTHSRV11");
            var dt = new DataTable();


            try
            {
                var query = Sql;

                var ds = new DataSet();
                using (var con = new SqlConnection(Con_RTHSRV11))
                {
                    using (var cmd = new SqlCommand())
                    {
                        cmd.Connection = con;
                        con.Open();
                        var adpterdata = new SqlDataAdapter();
                        adpterdata.SelectCommand = new SqlCommand(query, con);
                        adpterdata.Fill(dt);
                        con.Close();
                        if (dt.Rows.Count >= 1)
                        {
                            return "true";
                        }
                        else
                        {
                            return "false";
                        }

                    }
                }
            }
            catch (Exception e)
            {
                var dsa = e.ToString();
                return dsa;
            }
        }

        public string Get_RTHSRV04(string Sql)
        {
            //string strConString = @ "Data Source=WELCOME-PC\SQLSERVER2008;Initial Catalog=MyDB;Integrated Security=True";
            var Con_RTHSRV04 = configuration.GetConnectionString("Con_RTHSRV04");
            var dt = new DataTable();


            try
            {
                var query = Sql;

                var ds = new DataSet();
                using (var con = new SqlConnection(Con_RTHSRV04))
                {
                    using (var cmd = new SqlCommand())
                    {
                        cmd.Connection = con;
                        con.Open();
                        var adpterdata = new SqlDataAdapter();
                        adpterdata.SelectCommand = new SqlCommand(query, con);
                        adpterdata.Fill(dt);
                        con.Close();
                        if (dt.Rows.Count >= 1)
                        {
                            return "true";
                        }
                        else
                        {
                            return "false";
                        }

                    }
                }
            }
            catch (Exception e)
            {
                var dsa = e.ToString();
                return dsa;
            }
        }
        
            public string Get_RTHSRVWSlip(string Sql)
        {
            //string strConString = @ "Data Source=WELCOME-PC\SQLSERVER2008;Initial Catalog=MyDB;Integrated Security=True";
            var Con_RTHSRVWSlip = configuration.GetConnectionString("Con_RTHSRVWSlip");
            var dt = new DataTable();


            try
            {
                var query = Sql;

                var ds = new DataSet();
                using (var con = new SqlConnection(Con_RTHSRVWSlip))
                {
                    using (var cmd = new SqlCommand())
                    {
                        cmd.Connection = con;
                        con.Open();
                        var adpterdata = new SqlDataAdapter();
                        adpterdata.SelectCommand = new SqlCommand(query, con);
                        adpterdata.Fill(dt);
                        con.Close();
                        if (dt.Rows.Count >= 1)
                        {
                            return "true";
                        }
                        else
                        {
                            return "false";
                        }

                    }
                }
            }
            catch (Exception e)
            {
                var dsa = e.ToString();
                return dsa;
            }
        }
        public string Get_RTHSRVTR01(string Sql)
        {
            //string strConString = @ "Data Source=WELCOME-PC\SQLSERVER2008;Initial Catalog=MyDB;Integrated Security=True";
            var Con_RTHSRVTR01 = configuration.GetConnectionString("Con_RTHSRVTR01");
            var dt = new DataTable();


            try
            {
                var query = Sql;

                var ds = new DataSet();
                using (var con = new SqlConnection(Con_RTHSRVTR01))
                {
                    using (var cmd = new SqlCommand())
                    {
                        cmd.Connection = con;
                        con.Open();
                        var adpterdata = new SqlDataAdapter();
                        adpterdata.SelectCommand = new SqlCommand(query, con);
                        adpterdata.Fill(dt);
                        con.Close();
                        if (dt.Rows.Count >= 1)
                        {
                            return "true";
                        }
                        else
                        {
                            return "false";
                        }

                    }
                }
            }
            catch (Exception e)
            {
                var dsa = e.ToString();
                return dsa;
            }
        }
        
        public DataTable GetDataTableCmd(SqlCommand objCmd)
        {
            DataTable objDataTbl;
            SqlDataAdapter objDataAdp;
            //SqlConnection Con;        

            var constr = configuration.GetConnectionString("CONN");
            try
            {
                // make result DataTable instance
                objDataTbl = new DataTable();
                using (var connection = new SqlConnection(constr))
                {

                    objDataAdp = new SqlDataAdapter();
                    objCmd.Connection = connection;
                    objDataAdp.SelectCommand = objCmd;
                    objDataAdp.Fill(objDataTbl);


                    if (connection.State == ConnectionState.Open)
                    {
                        connection.Close();
                    }

                    return objDataTbl;

                }


            }
            catch (SqlException sqlEx)
            {

                throw new Exception(sqlEx.Message);

            }
            catch (Exception ex)
            {
                throw ex;
            }
        
        }


        public DataSet GetDataSets(string SQL)
        {
            try
            {

                DataSet dtSet;
                dtSet = new DataSet();
                var con = configuration.GetConnectionString("CONN");
                using(var Connect = new SqlConnection(con))
                {
                    //using(var cmd = new SqlCommand())
                    //{
                    //    cmd.CommandText = "";
                        Connect.Open();
                        var DataAdapter = new SqlDataAdapter();
                        DataAdapter.SelectCommand = new SqlCommand(SQL, Connect);
                        DataAdapter.Fill(dtSet);
                        Connect.Close();
                    //}

                }


                return dtSet;

            }
            catch (Exception e)
            {
                throw e;
            }
        }


        public int ExcuteProc(string Str)
        {
            int excproc = 0;

            try
            {

                var conn = configuration.GetConnectionString("CONN");
               
                using(var con = new SqlConnection(conn))
                {
                    using(var cmd = new SqlCommand())
                    {
                        cmd.Connection = con;
                        cmd.CommandText = Str;
                        con.Open();
                        excproc = cmd.ExecuteNonQuery();
                        con.Close();
                       
                    }
                }

                return excproc;
            }
            catch (Exception ex)
            {
                throw ex;
            }
           
        }


        public string ExcuteStore(string Str)
        {
            var dt = new DataTable();
            string excproc = "";

            try
            {

                var conn = configuration.GetConnectionString("CONN");

                using (var con = new SqlConnection(conn))
                {
                    using (var cmd = new SqlCommand())
                    {
                        cmd.Connection = con;
                        con.Open();
                        var adpterdata = new SqlDataAdapter();
                        adpterdata.SelectCommand = new SqlCommand(Str, con);
                        adpterdata.Fill(dt);
                        con.Close();
                        excproc = dt.Rows[0][0].ToString();

                    }
                }

                return excproc;
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }




    }

}