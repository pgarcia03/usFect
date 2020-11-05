using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Web.Configuration;

namespace SistemaAuditores.DataAccess
{
    public class DataAccess
    {
        public static string Get_ConnectionString()
        {
            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
            builder.DataSource = WebConfigurationManager.AppSettings["DataSource"];
            builder.InitialCatalog = WebConfigurationManager.AppSettings["DataBase"];
            builder.UserID = WebConfigurationManager.AppSettings["User"];
            builder.Password = WebConfigurationManager.AppSettings["Password"];
            return builder.ConnectionString;
        }

        public static DataTable Get_DataTable(string query)
        {
            DataTable dt = new DataTable();
            SqlConnection conn = new SqlConnection(Get_ConnectionString());
            try
            {
                conn.Open();
                SqlDataAdapter da = new SqlDataAdapter(query, conn);
                da.Fill(dt);
                return dt;
            }
            catch
            {
                return null;
            }
            finally
            {
                conn.Close();
            }
        }

        public static string Get_UserbyUserName(string username)
        {
            string query = "select IdUsuario from tblusuario where username='" + username.ToString() + "'";
            DataTable dt = Get_DataTable(query);
            string IdUsuario = Convert.ToString(dt.Rows[0]["IdUsuario"]);
            return IdUsuario;
        }

        public static void Execute_Query(string query)
        {
            SqlConnection conn = new SqlConnection(Get_ConnectionString());
            SqlCommand comm = new SqlCommand(query, conn);
            try
            {
                conn.Open();
                comm.ExecuteNonQuery();
            }
            catch
            {
            }
            finally
            {
                conn.Close();
            }
        }

        public static void Execute_StoredProcedure(string sp, SqlParameterCollection spc)
        {
            SqlConnection conn = new SqlConnection(Get_ConnectionString());
            SqlCommand comm = new SqlCommand(sp, conn);
            comm.CommandType = CommandType.StoredProcedure;
            int index = 0;

            for (int i = 0; i < spc.Count; i++)
            {
                comm.Parameters.Add(spc[i]);
                if (spc[i].Direction == ParameterDirection.InputOutput)
                {
                    index = i;
                    comm.Parameters[comm.Parameters.Count - 1].Direction = ParameterDirection.InputOutput;
                }

            }

        }

    }
}

