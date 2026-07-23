using DevExpress.XtraEditors;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ComLib
{
    public static class DBConn
    {
        public static SqlConnection dbCon;
        public static SqlTransaction dbTran;
        public static string sqlConnection = "server = 192.168.0.202; uid = sa; pwd = jang*1976; database = daejierp";//실서버(내부-배포용) admin Admin12345
        //public static string sqlConnection = "server = 61.32.101.234,1433; uid = sa; pwd = jang*1976; database = daejierp";//실서버(외부) admin Admin12345
        //public static string sqlConnection = "server = 61.32.101.234,1433; uid = sa; pwd = jang*1976; database = test";//테스트서버 admin Admin12345
        //pubic static string sqlConnection = "server = 192.168.0.202; uid = sa; pwd = jang*1976; database = test";//테스트서버 admin Admin12345
        public static DataTable set;

        public static SqlConnection DbConn()
        {
            try
            {
                dbCon = new SqlConnection(sqlConnection);
                dbCon.Open();
                return dbCon;
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.ToString());
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
                XtraMessageBox.Show(ex.Message);
            }
        }

        public static DataTable GetDataTable(string strSql)
        {
            DataSet ds = new DataSet();

            try
            {
                using (SqlConnection conn = new SqlConnection(sqlConnection))
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
                MessageBox.Show(sqlEx.Message);
                return null;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
            finally
            {
                if ((ds != null)) ds.Dispose();
            }
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
                XtraMessageBox.Show(sqlEx.Message);
                return null;
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
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
                DataTable dataTable = new DataTable();

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


                if(ds.Tables != null && ds.Tables.Count > 0)
                {
                    dataTable = ds.Tables[0];
                }

                return dataTable;
            }
            catch (SqlException sqlEx)
            {
                XtraMessageBox.Show(sqlEx.Message);
                return null;
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
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
                XtraMessageBox.Show(sqlEx.Message);
                return null;
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
                return null;
            }
        }

        //2021-06-15 Log 저장용
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
                XtraMessageBox.Show(sqlEx.Message);
                return null;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public static DataTable GetDataTableByParam(string strSql, Dictionary<string, string> dicParams)
        {
            DataSet ds = new DataSet();

            try
            {
                using(SqlConnection conn = new SqlConnection(sqlConnection))
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
                XtraMessageBox.Show(sqlEx.Message);
                return null;
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
                return null;
            }
            finally
            {
                if ((ds != null)) ds.Dispose();
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
                XtraMessageBox.Show(sqlEx.Message);
                return null;
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
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

                tcon.Open();
                cnt = sqlCmd.ExecuteNonQuery();
                tcon.Close();

            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
            finally
            {
                if (tcon.State != ConnectionState.Closed)
                    tcon.Close();
            }

            return cnt;

        }

        public static DataTable GetDataTable(string spName, Dictionary<string, string> paramsDic)
        {
            DataSet ds = new DataSet();

            try
            {
                using (SqlConnection conn = new SqlConnection(sqlConnection))
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
                XtraMessageBox.Show(sqlEx.Message);
                return null;
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
                return null;
            }
            finally
            {
                if ((ds != null)) ds.Dispose();
            }
        }
    }
}
