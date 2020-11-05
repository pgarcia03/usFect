using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Web;

namespace SistemaAuditores.DataAccess
{
    public class OrderDetailDA : DataAccess
    {
       
        #region medidasNew

        public static string saveObjMaster(int idestacion, int idestilotalla, int idporder, bool estado)
        {

            SqlConnection conn = new SqlConnection(Get_ConnectionString());
            SqlCommand comm = new SqlCommand();

            try
            {

                // string idcpompuesto = string.Concat(idestacion.ToString(), idestilotalla.ToString(), idporder.ToString(), DateTime.Now.ToString().Replace('/',' ').TrimEnd());
                string idcpompuesto = string.Concat(idestacion.ToString(), idestilotalla.ToString(), idporder.ToString(), DateTime.Now.ToShortDateString().Replace('/', '.'), DateTime.Now.ToLongTimeString());

                conn.Open();
                comm.Connection = conn;
                comm.CommandText = "spdGuardarObjMaster";
                comm.CommandType = CommandType.StoredProcedure;
                comm.Parameters.Clear();

                comm.Parameters.Add("@idestacion", SqlDbType.Int);
                comm.Parameters[comm.Parameters.Count - 1].Value = idestacion;

                comm.Parameters.Add("@idestilotalla", SqlDbType.Int);
                comm.Parameters[comm.Parameters.Count - 1].Value = idestilotalla;

                comm.Parameters.Add("@idporder", SqlDbType.Int);
                comm.Parameters[comm.Parameters.Count - 1].Value = idporder;

                comm.Parameters.Add("@estado", SqlDbType.Bit);
                comm.Parameters[comm.Parameters.Count - 1].Value = estado;

                comm.Parameters.Add("@id", SqlDbType.NVarChar);
                comm.Parameters[comm.Parameters.Count - 1].Value = idcpompuesto;


                //comm.Parameters.Add("@idbbrechazo", SqlDbType.Int);
                //comm.Parameters[comm.Parameters.Count - 1].Direction = ParameterDirection.Output;

                comm.ExecuteNonQuery();

                //var id = Convert.ToInt32(comm.Parameters["@idbbrechazo"].Value);

                return idcpompuesto;

            }
            catch (Exception ex)
            {
                string mes = ex.Message;
                return "";
            }
            finally
            {
                conn.Close();
            }
        }

        public static string saveObjDetail(int idespecificacion, float valor, float valorspeck, string estado, string id)
        {

            SqlConnection conn = new SqlConnection(Get_ConnectionString());
            SqlCommand comm = new SqlCommand();

            try
            {
                conn.Open();
                comm.Connection = conn;
                comm.CommandText = "spdGuardarObjDetail";
                comm.CommandType = CommandType.StoredProcedure;
                comm.Parameters.Clear();

                comm.Parameters.Add("@idespecificacion", SqlDbType.Int);
                comm.Parameters[comm.Parameters.Count - 1].Value = idespecificacion;

                comm.Parameters.Add("@valor", SqlDbType.Float);
                comm.Parameters[comm.Parameters.Count - 1].Value = valor;

                comm.Parameters.Add("@valorspeck", SqlDbType.Float);
                comm.Parameters[comm.Parameters.Count - 1].Value = valorspeck;

                comm.Parameters.Add("@estado", SqlDbType.NChar);
                comm.Parameters[comm.Parameters.Count - 1].Value = estado;

                comm.Parameters.Add("@id", SqlDbType.NVarChar);
                comm.Parameters[comm.Parameters.Count - 1].Value = id;


                //comm.Parameters.Add("@idbbrechazo", SqlDbType.Int);
                //comm.Parameters[comm.Parameters.Count - 1].Direction = ParameterDirection.Output;

                comm.ExecuteNonQuery();

                //var id = Convert.ToInt32(comm.Parameters["@idbbrechazo"].Value);

                return "true";

            }
            catch (Exception ex)
            {
                string mes = ex.Message;
                return "";
            }
            finally
            {
                conn.Close();
            }
        }

        public static string saveEspecificacion(int idestilotalla, int idpuntoMedida, double valor, double valorMax, double valorMin, string user)
        {

            SqlConnection conn = new SqlConnection(Get_ConnectionString());
            SqlCommand comm = new SqlCommand();

            try
            {

                conn.Open();
                comm.Connection = conn;
                comm.CommandText = "spdSaveEspecificacion";
                comm.CommandType = CommandType.StoredProcedure;
                comm.Parameters.Clear();

                comm.Parameters.Add("@idestilotalla", SqlDbType.Int);
                comm.Parameters[comm.Parameters.Count - 1].Value = idestilotalla;

                comm.Parameters.Add("@idpuntoMedida", SqlDbType.Int);
                comm.Parameters[comm.Parameters.Count - 1].Value = idpuntoMedida;

                comm.Parameters.Add("@valor", SqlDbType.Float);
                comm.Parameters[comm.Parameters.Count - 1].Value = valor;

                comm.Parameters.Add("@valorMax", SqlDbType.Float);
                comm.Parameters[comm.Parameters.Count - 1].Value = valorMax;

                comm.Parameters.Add("@valormin", SqlDbType.Float);
                comm.Parameters[comm.Parameters.Count - 1].Value = valorMin;

                comm.Parameters.Add("@usuario", SqlDbType.NVarChar);
                comm.Parameters[comm.Parameters.Count - 1].Value = user;

                //comm.Parameters.Add("@idbbrechazo", SqlDbType.Int);
                //comm.Parameters[comm.Parameters.Count - 1].Direction = ParameterDirection.Output;

                comm.ExecuteNonQuery();

                //var id = Convert.ToInt32(comm.Parameters["@idbbrechazo"].Value);

                return "OK";

            }
            catch (Exception ex)
            {
                string mes = ex.Message;
                return "ERROR";
            }
            finally
            {
                conn.Close();
            }
        }

        public static string updateEspecificacion(int idespecificacion, double valor, double valorMax, double valorMin, string user)
        {

            SqlConnection conn = new SqlConnection(Get_ConnectionString());
            SqlCommand comm = new SqlCommand();

            try
            {

                conn.Open();
                comm.Connection = conn;
                comm.CommandText = "spdUpdateEspecificacion";
                comm.CommandType = CommandType.StoredProcedure;
                comm.Parameters.Clear();

                comm.Parameters.Add("@idespecificacion", SqlDbType.Int);
                comm.Parameters[comm.Parameters.Count - 1].Value = idespecificacion;

                comm.Parameters.Add("@valor", SqlDbType.Float);
                comm.Parameters[comm.Parameters.Count - 1].Value = valor;

                comm.Parameters.Add("@valorMax", SqlDbType.Float);
                comm.Parameters[comm.Parameters.Count - 1].Value = valorMax;

                comm.Parameters.Add("@valormin", SqlDbType.Float);
                comm.Parameters[comm.Parameters.Count - 1].Value = valorMin;

                comm.Parameters.Add("@usuario", SqlDbType.NVarChar);
                comm.Parameters[comm.Parameters.Count - 1].Value = user;

                //comm.Parameters.Add("@idbbrechazo", SqlDbType.Int);
                //comm.Parameters[comm.Parameters.Count - 1].Direction = ParameterDirection.Output;

                comm.ExecuteNonQuery();

                //var id = Convert.ToInt32(comm.Parameters["@idbbrechazo"].Value);

                return "OK";

            }
            catch (Exception ex)
            {
                string mes = ex.Message;
                return "ERROR";
            }
            finally
            {
                conn.Close();
            }
        }

        public static string saveAsignacionTallas(int idestilo, int idtalla)
        {

            SqlConnection conn = new SqlConnection(Get_ConnectionString());
            SqlCommand comm = new SqlCommand();

            try
            {

                conn.Open();
                comm.Connection = conn;
                comm.CommandText = "spdSaveEstiloTalla";
                comm.CommandType = CommandType.StoredProcedure;
                comm.Parameters.Clear();

                comm.Parameters.Add("@idestilo", SqlDbType.Int);
                comm.Parameters[comm.Parameters.Count - 1].Value = idestilo;

                comm.Parameters.Add("@idtalla", SqlDbType.Int);
                comm.Parameters[comm.Parameters.Count - 1].Value = idtalla;


                comm.ExecuteNonQuery();

                //var id = Convert.ToInt32(comm.Parameters["@idbbrechazo"].Value);

                return "OK";

            }
            catch (Exception ex)
            {
                string mes = ex.Message;
                return "ERROR";
            }
            finally
            {
                conn.Close();
            }
        }

        public static string saveAsignacionPuntoM(int idestilo, int idtalla, int idpuntom)//Agregar Parte Spec
        {

            SqlConnection conn = new SqlConnection(Get_ConnectionString());
            SqlCommand comm = new SqlCommand();

            try
            {

                conn.Open();
                comm.Connection = conn;
                comm.CommandText = "spdAgregarPuntoMSpec";
                comm.CommandType = CommandType.StoredProcedure;
                comm.Parameters.Clear();

                comm.Parameters.Add("@idestilo", SqlDbType.Int);
                comm.Parameters[comm.Parameters.Count - 1].Value = idestilo;

                comm.Parameters.Add("@idtalla", SqlDbType.Int);
                comm.Parameters[comm.Parameters.Count - 1].Value = idtalla;

                comm.Parameters.Add("@idPuntom", SqlDbType.Int);
                comm.Parameters[comm.Parameters.Count - 1].Value = idpuntom;

                comm.ExecuteNonQuery();

                return "OK";

            }
            catch (Exception ex)
            {
                string mes = ex.Message;
                return "ERROR";
            }
            finally
            {
                conn.Close();
            }
        }

        public static string IngresarSpec(int idespecificacion, float valor, float tmax, float tmin)//Agregar Parte Spec
        {

            SqlConnection conn = new SqlConnection(Get_ConnectionString());
            SqlCommand comm = new SqlCommand();

            try
            {

                conn.Open();
                comm.Connection = conn;
                comm.CommandText = "spdAgregarEspecificacion";
                comm.CommandType = CommandType.StoredProcedure;
                comm.Parameters.Clear();

                comm.Parameters.Add("@idespecificacion", SqlDbType.Int);
                comm.Parameters[comm.Parameters.Count - 1].Value = idespecificacion;

                comm.Parameters.Add("@valor", SqlDbType.Float);
                comm.Parameters[comm.Parameters.Count - 1].Value = valor;

                comm.Parameters.Add("@tmax", SqlDbType.Float);
                comm.Parameters[comm.Parameters.Count - 1].Value = tmax;

                comm.Parameters.Add("@tmin", SqlDbType.Float);
                comm.Parameters[comm.Parameters.Count - 1].Value = tmin;

                comm.ExecuteNonQuery();

                return "OK";

            }
            catch (Exception ex)
            {
                string mes = ex.Message;
                return "ERROR";
            }
            finally
            {
                conn.Close();
            }
        }

        public static int SaveTallaParametroSalida(string talla)
        {
            SqlConnection conn = new SqlConnection(Get_ConnectionString());
            SqlCommand comm = new SqlCommand();
            DataTable a = new DataTable();
            try
            {
                conn.Open();
                comm.Connection = conn;
                comm.CommandText = "spdSaveTallaParametroSalida";
                comm.CommandType = CommandType.StoredProcedure;
                comm.Parameters.Clear();

                comm.Parameters.Add("@talla", SqlDbType.VarChar);
                comm.Parameters[comm.Parameters.Count - 1].Value = talla;

                comm.Parameters.Add("@idtalla", SqlDbType.Int);
                comm.Parameters[comm.Parameters.Count - 1].Direction = ParameterDirection.Output;

                comm.ExecuteNonQuery();
                conn.Close();

                var idtalla =Convert.ToInt32(comm.Parameters["@idtalla"].Value.ToString());

                return idtalla;

            }
            catch (Exception ex)
            {
                string m = ex.Message;
                return 0;
            }
            finally
            {
                conn.Close();
            }

        }

        public static int SavePuntoMedidaParametroSalida(string puntomedida)
        {
            SqlConnection conn = new SqlConnection(Get_ConnectionString());
            SqlCommand comm = new SqlCommand();
            DataTable a = new DataTable();
            try
            {
                conn.Open();
                comm.Connection = conn;
                comm.CommandText = "spdSavePuntoMedidaParametroSalida";
                comm.CommandType = CommandType.StoredProcedure;
                comm.Parameters.Clear();

                comm.Parameters.Add("@puntomedida", SqlDbType.VarChar);
                comm.Parameters[comm.Parameters.Count - 1].Value = puntomedida;

                comm.Parameters.Add("@idpuntoM", SqlDbType.Int);
                comm.Parameters[comm.Parameters.Count - 1].Direction = ParameterDirection.Output;

                comm.ExecuteNonQuery();
                conn.Close();

                var idpuntoM = Convert.ToInt32(comm.Parameters["@idpuntoM"].Value.ToString());

                return idpuntoM;

            }
            catch (Exception ex)
            {
                string m = ex.Message;
                return 0;
            }
            finally
            {
                conn.Close();
            }

        }

        public static string saveEspecificacionNew(int idestilo,int idtalla, int idpuntoMedida, double valor, double valorMax, double valorMin, string user,int rango)
        {

            SqlConnection conn = new SqlConnection(Get_ConnectionString());
            SqlCommand comm = new SqlCommand();

            try
            {
               
                conn.Open();
                comm.Connection = conn;
                comm.CommandText = "spdSaveEspecificacionNew";
                comm.CommandType = CommandType.StoredProcedure;
                comm.Parameters.Clear();

                comm.Parameters.Add("@idestilo", SqlDbType.Int);
                comm.Parameters[comm.Parameters.Count - 1].Value = idestilo;

                comm.Parameters.Add("@idtalla", SqlDbType.Int);
                comm.Parameters[comm.Parameters.Count - 1].Value = idtalla;

                comm.Parameters.Add("@idpuntoMedida", SqlDbType.Int);
                comm.Parameters[comm.Parameters.Count - 1].Value = idpuntoMedida;

                comm.Parameters.Add("@valor", SqlDbType.Float);
                comm.Parameters[comm.Parameters.Count - 1].Value = valor;

                comm.Parameters.Add("@valorMax", SqlDbType.Float);
                comm.Parameters[comm.Parameters.Count - 1].Value = valorMax;

                comm.Parameters.Add("@valormin", SqlDbType.Float);
                comm.Parameters[comm.Parameters.Count - 1].Value = valorMin;

                comm.Parameters.Add("@usuario", SqlDbType.NVarChar);
                comm.Parameters[comm.Parameters.Count - 1].Value = user;

                comm.Parameters.Add("@rango", SqlDbType.Int);
                comm.Parameters[comm.Parameters.Count - 1].Value = rango;

                comm.ExecuteNonQuery();

               
                return "OK";

            }
            catch (Exception ex)
            {
                string mes = ex.Message;
                return "ERROR";
            }
            finally
            {
                conn.Close();
            }
        }

        #endregion


    
    }
}
