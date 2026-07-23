using DevExpress.XtraEditors;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Threading.Tasks;

namespace Daeji_MONITERING
{
    public static class DBConn
    {
        public static SqlConnection dbCon;
        public static SqlTransaction dbTran;
        public static string sqlConn = ComnString.CONNECTION_STRING;
        public static DataTable set;
        public static SqlConnection DbConn()
        {
            string sqlConnection = ComnString.CONNECTION_STRING;

            try
            {
                dbCon = new SqlConnection(sqlConnection);
                dbCon.Open();
                return dbCon;
            }
            catch (Exception ex)
            {
                //XtraMessageBox.Show(ex.ToString());
                return null;
            }
        }

        public static void DbDisConn(SqlConnection dbCon1)
        {
            try
            {
                dbCon1.Close();
                dbCon1.Dispose();
                dbCon1 = null;
            }
            catch (Exception ex)
            {
                //XtraMessageBox.Show(ex.Message);
            }
        }
        public static DataTable GetDataTable(string strSql)
        {
            DataSet ds = new DataSet();

            try
            {
                using (SqlConnection conn = new SqlConnection(sqlConn))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand(strSql, conn))
                    {
                        cmd.CommandType = CommandType.Text;
                        //cmd.CommandText = strSql;

                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        da.Fill(ds);
                    }
                }

                return ds.Tables[0];
            }
            catch (SqlException sqlEx)
            {
                //XtraMessageBox.Show(sqlEx.Message);
                return null;
            }
            catch (Exception ex)
            {
                //XtraMessageBox.Show(ex.Message);
                return null;
            }
            finally
            {
                if ((ds != null)) ds.Dispose();
            }
        }
        public static DataTable GetDataTable(string spName, Dictionary<string, string> paramsDic)
        {
            DataSet ds = new DataSet();

            try
            {
                using (SqlConnection conn = new SqlConnection(sqlConn))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand(spName, conn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        foreach (KeyValuePair<string, string> param in paramsDic)
                        {
                            cmd.Parameters.AddWithValue(string.Format("@{0}", param.Key), param.Value);
                        }
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        da.Fill(ds);
                        set = ds.Tables[0];
                    }
                }

                return ds.Tables[0];
            }
            catch (SqlException sqlEx)
            {
                //XtraMessageBox.Show(sqlEx.Message);
                return null;
            }
            catch (Exception ex)
            {
                //XtraMessageBox.Show(ex.Message);
                return null;
            }
            finally
            {
                if ((ds != null)) ds.Dispose();
            }
        }

        public static DataTable GetDataTable(string spName, Dictionary<string, object> paramsDic)
        {
            DataSet ds = new DataSet();

            try
            {
                using (SqlConnection conn = new SqlConnection(sqlConn))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand(spName, conn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        foreach (KeyValuePair<string, object> param in paramsDic)
                        {
                            cmd.Parameters.AddWithValue(string.Format("@{0}", param.Key), param.Value);
                        }
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        da.Fill(ds);
                        set = ds.Tables[0];
                    }
                }

                return ds.Tables[0];
            }
            catch (SqlException sqlEx)
            {
                //XtraMessageBox.Show(sqlEx.Message);
                return null;
            }
            catch (Exception ex)
            {
                //XtraMessageBox.Show(ex.Message);
                return null;
            }
            finally
            {
                if ((ds != null)) ds.Dispose();
            }
        }

        public static int GetDataByte(SqlConnection dc, string spName, Dictionary<string, string> paramsDic)
        {
            DataSet ds = new DataSet();
            byte[] byteArray;
            int byteLength = 0;
            try
            {
                using (SqlConnection conn = new SqlConnection(sqlConn))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand(spName, conn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        foreach (KeyValuePair<string, string> param in paramsDic)
                        {
                            cmd.Parameters.AddWithValue(string.Format("@{0}", param.Key), param.Value);
                        }
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        da.Fill(ds);
                        ds.AcceptChanges();
                        byteArray = GetByteArray(ds);
                        byteLength = Buffer.ByteLength(byteArray);

                    }
                }

                return byteLength;
            }
            catch (SqlException sqlEx)
            {
                //XtraMessageBox.Show(sqlEx.Message);
                return 0;
            }
            catch (Exception ex)
            {
                //XtraMessageBox.Show(ex.Message);
                return 0;
            }
            finally
            {
                if ((ds != null)) ds.Dispose();
            }
        }

        public static byte[] GetByteArray(DataSet sourceDataSet) {

            MemoryStream memoryStream = new MemoryStream();
            IFormatter formatter = new BinaryFormatter();
            sourceDataSet.RemotingFormat = SerializationFormat.Binary;
            formatter.Serialize(memoryStream, sourceDataSet);
            byte[] targetByteArray = memoryStream.ToArray();
            memoryStream.Close(); memoryStream.Dispose();
            return targetByteArray;
        }



        public static DataTable GetDataTable(SqlConnection dc, string strSql)
        {
            DataSet ds = new DataSet();

            try
            {
                SqlDataAdapter adpt = new SqlDataAdapter(strSql, dc);

                adpt.SelectCommand.Transaction = dbTran;

                adpt.Fill(ds);

                return ds.Tables[0];
            }
            catch (SqlException sqlEx)
            {
                //XtraMessageBox.Show(sqlEx.Message);
                return null;
            }
            catch (Exception ex)
            {
                //XtraMessageBox.Show(ex.Message);
                return null;
            }
            finally
            {
                if ((ds != null)) ds.Dispose();
            }
        }

        public static DataTable GetDataTable(SqlConnection dc, string spName, Dictionary<string, string> paramsDic)
        {
            try
            {
                SqlCommand sqlCmd = new SqlCommand(spName, dc);
                sqlCmd.CommandType = CommandType.StoredProcedure;

                foreach (KeyValuePair<string, string> item in paramsDic)
                {
                    string pname = String.Format("@{0}", item.Key);

                    sqlCmd.Parameters.Add(pname, SqlDbType.NVarChar);
                    sqlCmd.Parameters[pname].Value = item.Value;
                }

                SqlDataAdapter adp = new SqlDataAdapter(sqlCmd);
                DataSet ds = new DataSet();
                adp.Fill(ds);

                return ds.Tables[0];
            }
            catch (SqlException sqlEx)
            {
                //XtraMessageBox.Show(sqlEx.Message);
                return null;
            }
            catch (Exception ex)
            {
                //XtraMessageBox.Show(ex.Message);
                return null;
            }
        }

        public static DataTable GetDataTabl_Oj(SqlConnection dc, string spName, Dictionary<string, object> paramsDic)
        {
            try
            {
                SqlCommand sqlCmd = new SqlCommand(spName, dc);
                sqlCmd.CommandType = CommandType.StoredProcedure;

                foreach (KeyValuePair<string, object> item in paramsDic)
                {
                    string pname = String.Format("@{0}", item.Key);
                    if (item.Key.Equals("USIGN"))
                    {
                        sqlCmd.Parameters.Add(pname, SqlDbType.Image);
                        
                    }
                    else if (item.Key.Equals("PASSWD"))
                    {
                        sqlCmd.Parameters.Add(pname, SqlDbType.VarBinary);

                    }
                    else
                    {
                        sqlCmd.Parameters.Add(pname, SqlDbType.NVarChar);
                    }
                    sqlCmd.Parameters[pname].Value = item.Value;
                }

                SqlDataAdapter adp = new SqlDataAdapter(sqlCmd);
                DataSet ds = new DataSet();
                adp.Fill(ds);

                return ds.Tables[0];
            }
            catch (SqlException sqlEx)
            {
                //XtraMessageBox.Show(sqlEx.Message);
                return null;
            }
            catch (Exception ex)
            {
                //XtraMessageBox.Show(ex.Message);
                return null;
            }
        }

        public static DataTable GetDataTable_1(SqlConnection dc, string spName, Dictionary<string, string> paramsDic)
        {
            try
            {
                SqlCommand sqlCmd = new SqlCommand(spName, dc);
                sqlCmd.CommandType = CommandType.StoredProcedure;

                foreach (KeyValuePair<string, string> item in paramsDic)
                {
                    string pname = String.Format("@{0}", item.Key);

                    sqlCmd.Parameters.Add(pname, SqlDbType.NVarChar);
                    sqlCmd.Parameters[pname].Value = item.Value;
                }

                SqlDataAdapter adp = new SqlDataAdapter(sqlCmd);
                DataSet ds = new DataSet();
                adp.Fill(ds);

                return ds.Tables[0];
            }
            catch (SqlException sqlEx)
            {
                //XtraMessageBox.Show(sqlEx.Message);
                return null;
            }
            catch (Exception ex)
            {
                return null;
            }
        }
        public static DataTable GetDataTable(SqlConnection dc, string spName, Dictionary<string, object> paramsDic)
        {
            try
            {
                SqlCommand sqlCmd = new SqlCommand(spName, dc);
                sqlCmd.CommandType = CommandType.StoredProcedure;
                foreach (KeyValuePair<string, object> item in paramsDic)
                {
                    string pname = String.Format("@{0}", item.Key);
                    sqlCmd.Parameters.AddWithValue(pname, item.Value);
                    //sqlCmd.Parameters.Add(pname, SqlDbType.NVarChar);
                    //sqlCmd.Parameters[pname].Value = item.Value;
                }

                SqlDataAdapter adp = new SqlDataAdapter(sqlCmd);
                DataSet ds = new DataSet();
                adp.Fill(ds);

                return ds.Tables[0];
            }
            catch (SqlException sqlEx)
            {
                //XtraMessageBox.Show(sqlEx.Message);
                return null;
            }
            catch (Exception ex)
            {
                //XtraMessageBox.Show(ex.Message);
                return null;
            }
        }
        public static DataTable GetDataTableByProcedure(SqlConnection dc, string spName, string[] flds, string[] data)
        {
            DataSet ds = new DataSet();

            try
            {
                SqlCommand sqlCmd = new SqlCommand(spName, dc);

                sqlCmd.CommandType = CommandType.StoredProcedure;

                for (int i = 0; i < flds.Length; i++)
                {
                    string pname = String.Format("@{0}", flds[i]);

                    sqlCmd.Parameters.Add(pname, SqlDbType.NVarChar);
                    sqlCmd.Parameters[pname].Value = data[i];
                }

                SqlDataAdapter adp = new SqlDataAdapter(sqlCmd);
                adp.SelectCommand.Transaction = dbTran;
                adp.Fill(ds);

                return ds.Tables[0];
            }
            catch (SqlException sqlEx)
            {
                //XtraMessageBox.Show(sqlEx.Message);
                return null;
            }
            catch (Exception ex)
            {
                //XtraMessageBox.Show(ex.Message);
                return null;
            }
            finally
            {
                if ((ds != null)) ds.Dispose();
            }
        }
        public static DataTable GetDataTableByParam(string strSql, Dictionary<string, string> dicParams)
        {
            DataSet ds = new DataSet();

            try
            {
                using (SqlConnection conn = new SqlConnection(sqlConn))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand())
                    {
                        cmd.Connection = conn;
                        cmd.CommandText = strSql;
                        cmd.Prepare();
                        foreach (KeyValuePair<string, string> param in dicParams)
                        {
                            cmd.Parameters.AddWithValue(string.Format("@{0}", param.Key), param.Value);
                        }
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        da.Fill(ds);
                    }
                }

                return ds.Tables[0];
            }
            catch (SqlException sqlEx)
            {
                //XtraMessageBox.Show(sqlEx.Message);
                return null;
            }
            catch (Exception ex)
            {
                //XtraMessageBox.Show(ex.Message);
                return null;
            }
            finally
            {
                if ((ds != null)) ds.Dispose();
            }
        }
        public static int ExecuteNonQuery(string spName, Dictionary<string, string> paramsDic)
        {
            int cnt = 0;
            SqlConnection tcon = dbCon;

            try
            {

                SqlCommand sqlCmd = new SqlCommand(spName, tcon);
                sqlCmd.CommandType = CommandType.StoredProcedure;

                foreach (KeyValuePair<string, string> item in paramsDic)
                {
                    string pname = String.Format("@{0}", item.Key);

                    sqlCmd.Parameters.Add(pname, SqlDbType.NVarChar);
                    sqlCmd.Parameters[pname].Value = item.Value;
                }

                if (tcon.State == ConnectionState.Closed)
                    tcon.Open();
                cnt = sqlCmd.ExecuteNonQuery();
                tcon.Close();

            }
            catch (Exception ex)
            {
                //XtraMessageBox.Show(ex.Message);
            }
            finally
            {
                if (tcon.State != ConnectionState.Closed)
                    tcon.Close();
            }

            return cnt;

        }

        #region [접속 IP 정보]
        public static string Client_IP
        {
            get
            {
                IPHostEntry host = Dns.GetHostEntry(Dns.GetHostName());
                string ClientIP = string.Empty;
                for (int i = 0; i < host.AddressList.Length; i++)
                {
                    if (host.AddressList[i].AddressFamily == AddressFamily.InterNetwork)
                    {
                        ClientIP = host.AddressList[i].ToString();
                    }
                }
                return ClientIP;
            }
        }
        #endregion
    }
}
