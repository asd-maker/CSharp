using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data;
using System.Reflection;

namespace NJCHKOX
{
    class ChkOXClass
    {
        public const string CstrErrLogFileName = "ErrorLog";      /*ErrorLog FileName*/
        public const int CintInterval = 60;
        public static string strConn = "";
        public static string strConnOX = "";
        public static string[] strMailTo;
        public static string strSMTP = "";

        /*API解析INI檔的API*/
        [System.Runtime.InteropServices.DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def, System.Text.StringBuilder retVal, Int32 size, string filePath);
        /// <读取配置档信息，初始化数据库>
        /// 读取配置档信息，初始化数据库
        /// </summary>
        public void ReadConfig()
        {
            Console.WriteLine(System.DateTime.Now.ToString("yyyyMMddHHmmss") + " " + "開始讀取配置檔信息");
            strConn = "";
            strSMTP = "";
            strMailTo = null;
            try
            {
                /*定義INI的名稱，以及文件的路徑*/
                string strFileName = "ChkOX.ini";
                string strPath = Directory.GetCurrentDirectory() + "\\" + strFileName;

                /*判斷檔案是否存在，不存在停止執行程式*/
                FileInfo Fi = new FileInfo(strFileName);
                if (!Fi.Exists)
                {
                    RecordLog(CstrErrLogFileName, "配置檔案不存在，請確認后重新開啟");
                    return;
                }
                Console.WriteLine(System.DateTime.Now.ToString("yyyyMMddHHmmss") + " *********************************************************************************************");
                /*解析INI檔案的內容，并賦值*/
                string strDBServer = "", strDBName = "", strUserID = "", strPassword = "";
                strSMTP = ReadValue("System", "SMTP", strPath).Trim();
                Console.WriteLine(System.DateTime.Now.ToString("yyyyMMddHHmmss") + " " + "運行間隔時間:" + CintInterval.ToString() + "秒; 郵件服務器:" + strSMTP.ToString() + ";");
                Console.Title = "ChkOX " + System.DateTime.Now.ToString("yyyyMMddHHmmss") + " Version:" + AssemblyFileVersion.ToString();

                DBConn conn = new DBConn();
                strDBServer = ReadValue("MES", "DBServer", strPath).Trim();
                strDBName = ReadValue("MES", "DBName", strPath).Trim();
                strUserID = ReadValue("MES", "UserID", strPath).Trim();
                strPassword = ReadValue("MES", "Password", strPath).Trim();
                strConn = "Data Source=" + strDBServer + ";Initial Catalog=" + strDBName + ";Persist Security Info=True;User ID=" + strUserID + ";password=" + strPassword + ";Max Pool Size=10;Min Pool Size=1;Pooling=True;";
                if (conn.TestConnectSQLServerDB("MES", strConn) == false) { return; }

                /*OX MFGINFO DB*/
                strDBServer = ReadValue("OX", "DBServer", strPath).Trim();
                strDBName = ReadValue("OX", "DBName", strPath).Trim();
                strUserID = ReadValue("OX", "UserID", strPath).Trim();
                strPassword = ReadValue("OX", "Password", strPath).Trim();
                strConnOX = "Server=" + strDBServer + ";Data Source=" + strDBName + ";Persist Security Info=True;User ID=" + strUserID + ";password=" + strPassword + ";";
                if (conn.TestConnectOracleDB("OX DB", strConnOX) == false) { return; }
                Console.WriteLine(System.DateTime.Now.ToString("yyyyMMddHHmmss") + " *********************************************************************************************");
                Console.WriteLine(System.DateTime.Now.ToString("yyyyMMddHHmmss") + " " + "結束讀取配置檔信息");
            }
            catch (Exception ex)
            {
                RecordLog(CstrErrLogFileName, ex.ToString());
            }
        }
        /// <检查箱号信息>
        /// 检查箱号信息
        /// </summary>
        /// <param name="strCartonID"></param>
        /// <returns></returns>
        public bool ChkCartonData(string strCartonID)
        {
            try
            {
                DBConn conn = new DBConn();
                string strSQL = "";

                strSQL += "SELECT CARTON_ID FROM dbo.MFGINFO WHERE CARTON_ID='" + strCartonID + "'";

                DataSet ds= conn.GetSQLServerDataSet(strConn, strSQL);
                if (ds.Tables.Count == 0 || ds == null)
                {
                    return false;
                }

                if (ds.Tables[0].Rows.Count > 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                RecordLog(CstrErrLogFileName, ex.ToString());
                return false;
            }
        }
        /// <获取箱号信息>
        /// 获取箱号信息
        /// </summary>
        /// <returns></returns>
        public DataTable GetCartonData()
        {
            DataTable dt = new DataTable();
            try
            {
                DBConn conn = new DBConn();
                string strSQL = " SELECT DATYPE,TARGET,MailTo,[START_DATE],END_DATE FROM EIF.PARAM WHERE DATYPE='CHKOX' ";
                DataTable dtParam = conn.GetSQLServerDataSet(strConn, strSQL).Tables[0];
                if (dtParam.Rows.Count == 1)
                {
                    strMailTo = dtParam.Rows[0]["MailTo"].ToString().Split(';');
                    //strMailTo = "anday.wang@Innolux.com;".Split(';');
                }
                else
                {
                    strMailTo = "lei.l.zhao@Innolux.com;".Split(';');
                }
                dtParam.Clear();
                dtParam.Dispose();

                strSQL = " UPDATE EIF.PARAM SET [START_DATE] = GETDATE(), END_DATE = GETDATE() WHERE DATYPE='CHKOX' ";
                int intExec = conn.ExecuteSQLServerSQL(strConn, strSQL);
                strSQL = "select distinct trackno  from njwms.G_CONTAINER_FOR_MES where to_char(recvdatetime,'YYYY/MM/DD HH24:MI:SS') >to_char(sysdate-1/24,'YYYY/MM/DD HH24:MI:SS')  ";//撈取每小時的入庫資料
                dt = conn.GetOracleDataSet(strConnOX, strSQL).Tables[0];
                return dt;
            }
            catch (Exception ex)
            {
                RecordLog(CstrErrLogFileName, ex.ToString());
                return dt;
            }
        }
        public string AssemblyFileVersion
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyFileVersionAttribute), false);
                if (attributes.Length == 0)
                    return "";
                return ((AssemblyFileVersionAttribute)attributes[0]).Version;
            }
        }
        /// <解析INI檔案的邏輯>
        /// 解析INI檔案的邏輯
        /// </summary>
        /// <param name="strSection"></param>
        /// <param name="strKey"></param>
        /// <param name="strPath"></param>
        /// <returns></returns>
        public string ReadValue(string strSection, string strKey, string strPath)
        {
            /*解析INI檔案的邏輯*/
            System.Text.StringBuilder temp = new System.Text.StringBuilder(4096);
            GetPrivateProfileString(strSection, strKey, "", temp, 4096, strPath);
            return temp.ToString();
        }
        /// <记录检查中的报错信息，并发送邮件提醒>
        /// 记录检查中的报错信息，并发送邮件提醒
        /// </summary>
        /// <param name="strPath"></param>
        /// <param name="strMsg"></param>
        public void RecordLog(string strPath, string strMsg)
        {
            try
            {
                if (!Directory.Exists(strPath))
                {
                    Directory.CreateDirectory(strPath);
                }
                FileStream Fs = new FileStream(strPath + "/" + System.DateTime.Now.ToString("yyyy-MM-dd") + ".txt", FileMode.Append);
                StreamWriter Sw = new StreamWriter(Fs);
                Sw.WriteLine("DateTime : " + System.DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss") + "; Message : " + strMsg);
                Sw.WriteLine("----------------------------------------------------------------------------------------------------------");
                Sw.Close();
                Fs.Close();
                Console.WriteLine(System.DateTime.Now.ToString("yyyyMMddHHmmss") + " " + strMsg);
                if (SendMail(strSMTP, "NJGetOX@Innolux.com", strMailTo, "NJ【◆警告◆：NJCHKOX執行失敗，請值班人員確認處理】", strMsg, "") == false)
                {
                    Console.WriteLine(System.DateTime.Now.ToString("yyyyMMddHHmmss") + " 郵件發送失敗");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(System.DateTime.Now.ToString("yyyyMMddHHmmss") + " " + ex.ToString());
            }
        }
        /// <RecordLog记录报错信息后发送邮件提醒>
        /// RecordLog记录报错信息后发送邮件提醒
        /// </summary>
        /// <param name="strSmtp"></param>
        /// <param name="strFromAdd"></param>
        /// <param name="strToAdd"></param>
        /// <param name="strSubject"></param>
        /// <param name="strBody"></param>
        /// <param name="strFile"></param>
        /// <returns></returns>
        public bool SendMail(string strSmtp, string strFromAdd, string[] strToAdd, string strSubject, string strBody, string strFile)
        {
            try
            {
                int j = 0;
                if (strFromAdd == "")
                {
                    return true;
                }
                if (strToAdd == null)
                {
                    strToAdd = "anday.wang@Innolux.com;".Split(';');
                }
                System.Net.Mail.MailMessage myMail = new System.Net.Mail.MailMessage();
                //myMail.BodyEncoding = System.Text.Encoding.UTF8;
                myMail.IsBodyHtml = true;
                myMail.From = new System.Net.Mail.MailAddress(strFromAdd);
                for (j = 0; j < Convert.ToInt32(strToAdd.Length); j++)
                {
                    if (strToAdd[j] != null && strToAdd[j] != "")
                    {
                        myMail.To.Add(strToAdd[j]);
                    }
                }
                strBody = strBody + " <table> ";
                strBody = strBody + " <tr><td style=background-color:Silver;>此郵件由MES發出(NJChkOX的程式發出,詳細請從VSS中獲取代碼查看)</td></tr> ";
                strBody = strBody + " </table> ";

                myMail.Subject = strSubject;
                myMail.Body = strBody;

                if (strFile != "")
                {
                    myMail.Attachments.Add(new System.Net.Mail.Attachment(strFile));
                }
                System.Net.Mail.SmtpClient SmtpMail = new System.Net.Mail.SmtpClient();
                SmtpMail.Host = strSmtp;
                SmtpMail.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;
                SmtpMail.Send(myMail);

                return true;
            }
            catch (Exception ex)
            {
                RecordLog(CstrErrLogFileName, ex.ToString());
                return false;
            }
        }
    }
}
