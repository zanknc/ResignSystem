using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;

namespace ResignSystem.Models.TSQL
{
    public class TEST
    {
        public DataTable dt(String Test)
        {


            string consString = "Data Source=RTHSRV14;Initial Catalog=ImportExportDB;Persist Security Info=True;User ID=sa;Password=pwpolicy;";
            using(var con = new SqlConnection(consString))
            {
                using(var cmd = new SqlCommand())
                {
                    cmd.Connection = con;
                    con.Open();
                    con.Close();

                }
            }
            var dt = new DataTable();
            return dt;
        }
    }
}
