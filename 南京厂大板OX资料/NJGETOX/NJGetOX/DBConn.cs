using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Data.OracleClient;
using System.Xml;
using System.Configuration;

#pragma warning disable 0618    // 加此 屏蔽OracleConnection過時的警告提示

namespace NJGetOX
{
    class DBConn
    {
        GetOXClass ec = new GetOXClass();
        public DBConn()
        {
            //
            // TODO: 在此处添加构造函数逻辑
            //
        }
        /*SqlServer2005的連接資訊Begin*/
        public DataSet GetSQLServerDataSet(string strConn, string strSQL)
        {
            SqlConnection conn = new SqlConnection();
            DataSet ds = new DataSet();
            try
            {
                conn = new SqlConnection(strConn);
                SqlDataAdapter sda = new SqlDataAdapter(strSQL, conn);
                sda.Fill(ds);
                sda.Dispose();
                return ds;
            }
            catch (Exception ex)
            {
                conn.Close();
                throw (ex);
            }
            finally
            {
                conn.Close();
            }
        }
        public int ExecuteSQLServerSQL(string strConn, string strSQL)
        {
            SqlConnection sqlConnCon = new SqlConnection();
            sqlConnCon = new SqlConnection(strConn);
            sqlConnCon.Open();
            try
            {
                SqlCommand sqlCmd = new SqlCommand(strSQL, sqlConnCon);
                return sqlCmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                ec.RecordLog(GetOXClass.CstrErrLogFileName, ex.ToString());
                throw ex;
            }
            finally
            {
                sqlConnCon.Close();
            }
        }
        public DataSet GetDataSetBySP(SqlConnection sqlConnCon, string spName, SqlParameter[] parameters)
        {
            try
            {
                if (sqlConnCon.Equals(null))
                {
                    throw new Exception("Connection is null");
                }
                SqlCommand sqlCommand = new SqlCommand();
                sqlCommand.Connection = sqlConnCon;
                sqlCommand.CommandType = CommandType.StoredProcedure;
                sqlCommand.CommandText = spName;
                sqlCommand.Parameters.AddRange(parameters);
                sqlCommand.CommandTimeout = 30;
                SqlDataAdapter sqlDataAdapter = new SqlDataAdapter();
                sqlDataAdapter.SelectCommand = sqlCommand;
                DataSet ds = new DataSet();
                sqlDataAdapter.Fill(ds);
                return ds;
            }
            catch (Exception ex)
            {
                ec.RecordLog(GetOXClass.CstrErrLogFileName, ex.ToString());
                throw ex;
            }
            finally
            {
                sqlConnCon.Close();
            }
        }
        /*SqlServer2005的連接資訊End*/
        /*Oracle的連接資訊Begin*/
        public DataSet GetOracleDataSet(string strConn, string strSQL)
        {
            OracleConnection conn = new OracleConnection();
            conn = new OracleConnection(strConn);
            conn.Open();
            DataSet ds = new DataSet();
            try
            {
                
                OracleDataAdapter sda = new OracleDataAdapter(strSQL, conn);
                sda.Fill(ds);
                sda.Dispose();
                return ds;
            }
            catch (Exception ex)
            {
                conn.Close();
                throw (ex);
            }
            finally
            {
                conn.Close();
            }
        }
        public int ExecuteOracleSQL(string strConn, string strSQL)
        {
            OracleConnection conn = new OracleConnection();
            try
            {
                conn = new OracleConnection(strConn);
                OracleCommand sqlCmd = new OracleCommand(strSQL, conn);
                conn.Open();
                return sqlCmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                conn.Close();
                throw ex;
            }
            finally
            {
                conn.Close();
            }
        }
        /*Oracle的連接資訊End*/

        public bool TestConnectSQLServerDB(string strDBName, string strConnTest)
        {
            try
            {
                SqlConnection conn = new SqlConnection(strConnTest);
                conn.Open();
                conn.Close();
                string[] strChar = strConnTest.Split(';');
                Console.WriteLine(System.DateTime.Now.ToString("yyyyMMddHHmmss") + " " + strDBName + "資料庫" + strChar[0] + ";" + strChar[1] + ";測試連接成功!");
                return true;
            }
            catch(Exception ex)
            {
                ec.RecordLog(GetOXClass.CstrErrLogFileName,ex.ToString());
                return false;
            }
        }
        public bool TestConnectOracleDB(string strDBName, string strConnTest)
        {
            try
            {
                OracleConnection conn = new OracleConnection(strConnTest);
                conn.Open();
                conn.Close();
                string[] strChar = strConnTest.Split(';');
                Console.WriteLine(System.DateTime.Now.ToString("yyyyMMddHHmmss") + " " + strDBName + "資料庫" + strChar[0] + ";" + strChar[1] + ";測試連接成功!");
                return true;
            }
            catch (Exception ex)
            {
                ec.RecordLog(GetOXClass.CstrErrLogFileName,  ex.ToString());
                return false;
            }
        }
    }
}
