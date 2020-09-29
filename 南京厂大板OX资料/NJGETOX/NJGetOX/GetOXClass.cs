using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data;
using System.Reflection;

namespace NJGetOX
{
    class GetOXClass
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
        public void ReadConfig()
        {
            Console.WriteLine(System.DateTime.Now.ToString("yyyyMMddHHmmss") + " " + "開始讀取配置檔信息");
            strConn = "";
            strSMTP = "";
            strMailTo = null;
            try
            {
                /*定義INI的名稱，以及文件的路徑*/
                string strFileName = "GetOX.ini";
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
                Console.Title = "GetOX " + System.DateTime.Now.ToString("yyyyMMddHHmmss") + " Version:" + AssemblyFileVersion.ToString();

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
        public string ReadValue(string strSection, string strKey, string strPath)
        {
            /*解析INI檔案的邏輯*/
            System.Text.StringBuilder temp = new System.Text.StringBuilder(4096);
            GetPrivateProfileString(strSection, strKey, "", temp, 4096, strPath);
            return temp.ToString();
        }
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
                if (SendMail(strSMTP, "NJGetOX@Innolux.com", strMailTo, "NJ【◆警告◆：NJGetOX執行失敗，請值班人員確認處理】", strMsg, "") == false)
                {
                    Console.WriteLine(System.DateTime.Now.ToString("yyyyMMddHHmmss") + " 郵件發送失敗");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(System.DateTime.Now.ToString("yyyyMMddHHmmss") + " " + ex.ToString());
            }
        }
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
                strBody = strBody + " <tr><td style=background-color:Silver;>此郵件由MES發出(NJGetOX的程式發出,詳細請從VSS中獲取代碼查看)</td></tr> ";
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
        public DataTable GetPalletData()
        {
            DataTable dt = new DataTable();
            try
            {
                DBConn conn = new DBConn();
                string strSQL = " SELECT DATYPE,TARGET,MailTo,[START_DATE],END_DATE FROM EIF.PARAM WHERE DATYPE='GETOX' ";
                DataTable dtParam = conn.GetSQLServerDataSet(strConn, strSQL).Tables[0];
                if (dtParam.Rows.Count == 1)
                {
                    strMailTo = dtParam.Rows[0]["MailTo"].ToString().Split(';');
                    //strMailTo = "anday.wang@Innolux.com;".Split(';');
                }
                else
                {
                    strMailTo = "anday.wang@Innolux.com;".Split(';');
                }
                dtParam.Clear();
                dtParam.Dispose();

                strSQL = " UPDATE EIF.PARAM SET [START_DATE] = GETDATE(), END_DATE = GETDATE() WHERE DATYPE='GETOX' ";
                int intExec = conn.ExecuteSQLServerSQL(strConn, strSQL);

                //strSQL = "SELECT DISTINCT PALLET_ID,CARTON_ID FROM MFGINFO";
                //strSQL = strSQL + " WHERE FACTORYID = 'CMO' ";
                ////strSQL = strSQL + " AND SHIPPINGTO = 'NJAL' ";
                //strSQL = strSQL + " AND SHIPPINGTO IN( 'NJAL', 'TOKL') ";
                //strSQL = strSQL + " AND SHIPPINGDATE >= TO_CHAR(SYSDATE - 5, 'YYYYMMDDHH24MISS')";
                //strSQL = strSQL + " AND UPDFLAG = 'N'";
                //strSQL = strSQL + " ORDER BY PALLET_ID,CARTON_ID";
                //2017-03-21 Modify
                strSQL = "SELECT DISTINCT PALLET_ID,CARTON_ID FROM MFGINFO";
                strSQL = strSQL + " WHERE SHIPPINGDATE >= TO_CHAR(SYSDATE - 3, 'YYYYMMDDHH24MISS')";
                strSQL = strSQL + " AND FACTORYID = 'CMO' ";
                strSQL = strSQL + " AND UPDFLAG = 'N'";
                //strSQL = strSQL + " AND SHIPPINGTO IN( 'NJAL', 'TOKL') ";            //2018-10-23 chunhui mark
                strSQL = strSQL + " AND SHIPPINGTO IN( 'NJAL', 'TOKL', 'LWGL','T006','WGTL','SNTL') ";      //2018-10-23 chunhui add  'LWGL'for沃格薄化廠 解決華為DP120 9530M0622092N 圈叉訊息無法收取問題  //20200310 add "WGTL"
                strSQL = strSQL + " AND NVL(UPDUSER,'NJ') <> 'NJMES'";
                strSQL = strSQL + " UNION ALL";
                strSQL = strSQL + " SELECT DISTINCT PALLET_ID,CARTON_ID FROM MFGINFO";
                strSQL = strSQL + " WHERE SHIPPINGDATE >= TO_CHAR(SYSDATE - 3, 'YYYYMMDDHH24MISS')";
                strSQL = strSQL + " AND UPDFLAG = 'N' ";
                strSQL = strSQL + " AND SHIPPINGFROM IN( 'T0', 'T1', 'T2','TP3P')";
                strSQL = strSQL + " AND NVL(SHIPPINGTO,'NJ') IN('NJ','WD01')";
                strSQL = strSQL + " AND SUBSTR(PRODUCTID,6,3) <= '100'";
                strSQL = strSQL + " AND NVL(UPDUSER,'NJ') <> 'NJMES'";
                //2018-07-24 ADD BEGIN 針對長信EOSP收取資料後南京收取不到資料問題修改
                strSQL = strSQL + " UNION ALL";
                strSQL = strSQL + " SELECT DISTINCT PALLET_ID,CARTON_ID FROM MFGINFO";
                strSQL = strSQL + " WHERE SHIPPINGDATE >= TO_CHAR(SYSDATE - 3, 'YYYYMMDDHH24MISS')";
                strSQL = strSQL + " AND UPDFLAG = 'Y' ";
                strSQL = strSQL + " AND SHIPPINGFROM ='T2'";
                strSQL = strSQL + " AND NVL(SHIPPINGTO,'NJ') IN('NJ','WD01')";
                strSQL = strSQL + " AND SUBSTR(PRODUCTID,6,3) <= '100'";
                strSQL = strSQL + " AND NVL(UPDUSER,'NJ') <> 'NJMES'";
                strSQL = strSQL + " AND DELIVERY_NO LIKE '007%'";
                //2018-07-24 ADD END
                dt = conn.GetOracleDataSet(strConnOX, strSQL).Tables[0];
                return dt;
            }
            catch (Exception ex)
            {
                RecordLog(CstrErrLogFileName, ex.ToString());
                return dt;
            }
        }
        public DataTable GetOXData(string strPalletID,string strCartonID)
        {
            DataTable dt = new DataTable();
            try
            {
                DBConn conn = new DBConn();
                string strSQL = "SELECT ";
                #region parameters
                strSQL = strSQL + " DELIVERY_NO";
                strSQL = strSQL + " ,FACTORYID";
                strSQL = strSQL + " ,SHIPPINGFROM";
                strSQL = strSQL + " ,SHIPPINGTO";
                strSQL = strSQL + " ,SHIPPINGDATE";
                strSQL = strSQL + " ,PALLET_ID";
                strSQL = strSQL + " ,CARTON_ID";
                strSQL = strSQL + " ,FORCE_EMPTY";
                strSQL = strSQL + " ,PRODUCTID";
                strSQL = strSQL + " ,PANEL_MODE";
                strSQL = strSQL + " ,PPBOX_GRADE";
                strSQL = strSQL + " ,GLASS_QTY";
                strSQL = strSQL + " ,ERCODE";
                strSQL = strSQL + " ,ERMESSAGE";
                strSQL = strSQL + " ,GSFLAG";
                strSQL = strSQL + " ,CT1FLAG";
                strSQL = strSQL + " ,CHIP_QTY";
                strSQL = strSQL + " ,PPBOX_COUNT";
                strSQL = strSQL + " ,GLASS_TYPE";
                strSQL = strSQL + " ,PE_PRODUCT_ID";
                strSQL = strSQL + " ,REJECT_REASON_CODE";
                strSQL = strSQL + " ,RETURN_COMMENT";
                strSQL = strSQL + " ,PALLET_OWNER";
                strSQL = strSQL + " ,PALLET_TYPE";
                strSQL = strSQL + " ,PALLET_ZONE_CODE";
                strSQL = strSQL + " ,PALLET_SET_CODE";
                strSQL = strSQL + " ,PP_BOX_ZONE_CODE";
                strSQL = strSQL + " ,PP_BOX_SETTING_CODE";
                strSQL = strSQL + " ,CONTAINER_NO";
                strSQL = strSQL + " ,FACTORYID_TO";
                strSQL = strSQL + " ,BOX_PN";
                strSQL = strSQL + " ,CASE WHEN PROCESSED_GLASSID IS Not Null THEN PROCESSED_GLASSID ELSE GLASSID END AS GLASSID";
                strSQL = strSQL + " ,TFT_POSITION"; //2017-03-21 Add
                strSQL = strSQL + " ,MES_SHIP_DATE";
                strSQL = strSQL + " ,SLOT_NO";
                strSQL = strSQL + " ,ORIGINALPRODUCT";
                strSQL = strSQL + " ,ORIGINALPLAN";
                strSQL = strSQL + " ,ORIGINALCELPARTNO";
                strSQL = strSQL + " ,ORIGINALLOTID";
                strSQL = strSQL + " ,ORIGINALLOTTYPE";
                strSQL = strSQL + " ,ORIGINALBATCH";
                strSQL = strSQL + " ,ORIGINALSIZE";
                strSQL = strSQL + " ,OXDATA";
                strSQL = strSQL + " ,ARRAYTESTREASONCODE";
                strSQL = strSQL + " ,ARRAYTESTTXNTIME";
                strSQL = strSQL + " ,CF_OX_INFO";
                strSQL = strSQL + " ,CFTESTREASONCODE";
                strSQL = strSQL + " ,CFTTESTTXNTIME";
                strSQL = strSQL + " ,ASMTESTDATA";
                strSQL = strSQL + " ,ASMTESTREASONCODE";
                strSQL = strSQL + " ,ASMTESTTXNTIME";
                strSQL = strSQL + " ,CT1DATA";
                strSQL = strSQL + " ,CT1REASON";
                strSQL = strSQL + " ,CT1TESTTXNTIME";
                strSQL = strSQL + " ,CT2DATA";
                strSQL = strSQL + " ,CT2REASON";
                strSQL = strSQL + " ,CT2TESTTXNTIME";
                strSQL = strSQL + " ,CELL_REPAIR_FLAG";
                strSQL = strSQL + " ,ENGINEERING";
                strSQL = strSQL + " ,OXDEFECODE";
                strSQL = strSQL + " ,QUADRANT";
                strSQL = strSQL + " ,OXFLAG";
                strSQL = strSQL + " ,OXGRADE";
                strSQL = strSQL + " ,OXLEVEL";
                strSQL = strSQL + " ,OWNER";
                strSQL = strSQL + " ,X";
                strSQL = strSQL + " ,Y";
                strSQL = strSQL + " ,TFT_REPAIR_FLAG";
                strSQL = strSQL + " ,TWO_CUT_DATE_TIME";
                strSQL = strSQL + " ,EXP_NO";
                strSQL = strSQL + " ,LASER_SHORT_RING_CUT";
                strSQL = strSQL + " ,RGB_DROP_HEIGHT";
                strSQL = strSQL + " ,PS_HEIGHT";
                strSQL = strSQL + " ,ODF_PULSE";
                strSQL = strSQL + " ,SHEET_JUDGE";
                strSQL = strSQL + " ,TFT_DEFECT_CODE";
                strSQL = strSQL + " ,TFT_DEFECT_X";
                strSQL = strSQL + " ,TFT_DEFECT_Y";
                strSQL = strSQL + " ,CF_DEFECT_CODE";
                strSQL = strSQL + " ,CF_DEFECT_X";
                strSQL = strSQL + " ,CF_DEFECT_Y";
                strSQL = strSQL + " ,CF_POSITION";
                strSQL = strSQL + " ,LCD_PACK_TIME";
                strSQL = strSQL + " ,CF_GLASS_ID";
                strSQL = strSQL + " ,TFT_POLARIZER";
                strSQL = strSQL + " ,CF_POLARIZER";
                strSQL = strSQL + " ,GOOD_CUT_COUNT";
                strSQL = strSQL + " ,CUT_OX_INFO";
                strSQL = strSQL + " ,UPDFLAG";
                strSQL = strSQL + " ,UPDUSER";
                strSQL = strSQL + " ,UPDTIME";
                #endregion
                strSQL = strSQL + " FROM MFGINFO";
                strSQL = strSQL + " WHERE PALLET_ID = '" + strPalletID + "' ";
                strSQL = strSQL + " AND CARTON_ID = '" + strCartonID + "' ";
                //strSQL = strSQL + " AND UPDFLAG = 'N'";                                //2018-07-24 MARK
                strSQL = strSQL + " AND (UPDFLAG = 'N' OR DELIVERY_NO LIKE '007%' )";    //2018-07-24 ADD
                strSQL = strSQL + " ORDER BY GLASSID,SHIPPINGDATE";//2020-04-21 zhaolei Chang-->Add SHIPPINGDATE
                dt = conn.GetOracleDataSet(strConnOX, strSQL).Tables[0];
                return dt;
            }
            catch (Exception ex)
            {
                RecordLog(CstrErrLogFileName, ex.ToString());
                return dt;
            }
        }
        public bool InsertOXData(string strPalletID, string strCartonID, DataTable dtOX)
        {
            try
            {
                DBConn conn = new DBConn();
                string strSQL = "";
                for (int i = 0; i < dtOX.Rows.Count; i++)
                {
                    string strGlassID = dtOX.Rows[i]["GLASSID"].ToString();

                    strSQL = " INSERT INTO MFGINFOH(DELIVERY_NO,FACTORYID,SHIPPINGFROM,SHIPPINGTO,SHIPPINGDATE,PALLET_ID,CARTON_ID,FORCE_EMPTY,PRODUCTID,PANEL_MODE,PPBOX_GRADE,GLASS_QTY,ERCODE,ERMESSAGE,GSFLAG,CT1FLAG,CHIP_QTY,PPBOX_COUNT,GLASS_TYPE";
                    strSQL = strSQL + " ,PE_PRODUCT_ID,REJECT_REASON_CODE,RETURN_COMMENT,PALLET_OWNER,PALLET_TYPE,PALLET_ZONE_CODE,PALLET_SET_CODE,PP_BOX_ZONE_CODE,PP_BOX_SETTING_CODE,CONTAINER_NO,FACTORYID_TO,BOX_PN,GLASSID,MES_SHIP_DATE,SLOT_NO";
                    strSQL = strSQL + " ,ORIGINALPRODUCT,ORIGINALPLAN,ORIGINALCELPARTNO,ORIGINALLOTID,ORIGINALLOTTYPE,ORIGINALBATCH,ORIGINALSIZE,OXDATA,ARRAYTESTREASONCODE,ARRAYTESTTXNTIME,CF_OX_INFO,CFTESTREASONCODE,CFTTESTTXNTIME,ASMTESTDATA";
                    strSQL = strSQL + " ,ASMTESTREASONCODE,ASMTESTTXNTIME,CT1DATA,CT1REASON,CT1TESTTXNTIME,CT2DATA,CT2REASON,CT2TESTTXNTIME,CELL_REPAIR_FLAG,ENGINEERING,OXDEFECODE,QUADRANT,OXFLAG,OXGRADE,OXLEVEL,OWNER,X,Y,TFT_REPAIR_FLAG";
                    strSQL = strSQL + " ,TWO_CUT_DATE_TIME,EXP_NO,LASER_SHORT_RING_CUT,RGB_DROP_HEIGHT,PS_HEIGHT,ODF_PULSE,SHEET_JUDGE,TFT_DEFECT_CODE,TFT_DEFECT_X,TFT_DEFECT_Y,CF_DEFECT_CODE,CF_DEFECT_X,CF_DEFECT_Y,CF_POSITION,LCD_PACK_TIME";
                    strSQL = strSQL + " ,CF_GLASS_ID,TFT_POLARIZER,CF_POLARIZER,GOOD_CUT_COUNT,CUT_OX_INFO,UPDFLAG,UPDUSER,UPDTIME,TFT_POSITION)";
                    strSQL = strSQL + " SELECT DELIVERY_NO,FACTORYID,SHIPPINGFROM,SHIPPINGTO,SHIPPINGDATE,PALLET_ID,CARTON_ID,FORCE_EMPTY,PRODUCTID,PANEL_MODE,PPBOX_GRADE,GLASS_QTY,ERCODE,ERMESSAGE,GSFLAG,CT1FLAG,CHIP_QTY,PPBOX_COUNT,GLASS_TYPE";
                    strSQL = strSQL + " ,PE_PRODUCT_ID,REJECT_REASON_CODE,RETURN_COMMENT,PALLET_OWNER,PALLET_TYPE,PALLET_ZONE_CODE,PALLET_SET_CODE,PP_BOX_ZONE_CODE,PP_BOX_SETTING_CODE,CONTAINER_NO,FACTORYID_TO,BOX_PN,GLASSID,MES_SHIP_DATE,SLOT_NO";
                    strSQL = strSQL + " ,ORIGINALPRODUCT,ORIGINALPLAN,ORIGINALCELPARTNO,ORIGINALLOTID,ORIGINALLOTTYPE,ORIGINALBATCH,ORIGINALSIZE,OXDATA,ARRAYTESTREASONCODE,ARRAYTESTTXNTIME,CF_OX_INFO,CFTESTREASONCODE,CFTTESTTXNTIME,ASMTESTDATA";
                    strSQL = strSQL + " ,ASMTESTREASONCODE,ASMTESTTXNTIME,CT1DATA,CT1REASON,CT1TESTTXNTIME,CT2DATA,CT2REASON,CT2TESTTXNTIME,CELL_REPAIR_FLAG,ENGINEERING,OXDEFECODE,QUADRANT,OXFLAG,OXGRADE,OXLEVEL,OWNER,X,Y,TFT_REPAIR_FLAG";
                    strSQL = strSQL + " ,TWO_CUT_DATE_TIME,EXP_NO,LASER_SHORT_RING_CUT,RGB_DROP_HEIGHT,PS_HEIGHT,ODF_PULSE,SHEET_JUDGE,TFT_DEFECT_CODE,TFT_DEFECT_X,TFT_DEFECT_Y,CF_DEFECT_CODE,CF_DEFECT_X,CF_DEFECT_Y,CF_POSITION,LCD_PACK_TIME";
                    strSQL = strSQL + " ,CF_GLASS_ID,TFT_POLARIZER,CF_POLARIZER,GOOD_CUT_COUNT,CUT_OX_INFO,UPDFLAG,UPDUSER,UPDTIME,TFT_POSITION";
                    strSQL = strSQL + " FROM dbo.MFGINFO";
                    strSQL = strSQL + " WHERE PALLET_ID = '" + strPalletID + "' AND CARTON_ID = '" + strCartonID + "' AND GLASSID = '" + strGlassID + "';";

                    strSQL = strSQL + " DELETE FROM dbo.MFGINFO";
                    strSQL = strSQL + " WHERE PALLET_ID = '" + strPalletID + "' AND CARTON_ID = '" + strCartonID + "' AND GLASSID = '" + strGlassID + "';";

                    strSQL = strSQL + " INSERT INTO MFGINFO(";
                    #region parameters
                    strSQL = strSQL + " DELIVERY_NO";
                    strSQL = strSQL + " ,FACTORYID";
                    strSQL = strSQL + " ,SHIPPINGFROM";
                    strSQL = strSQL + " ,SHIPPINGTO";
                    strSQL = strSQL + " ,SHIPPINGDATE";
                    strSQL = strSQL + " ,PALLET_ID";
                    strSQL = strSQL + " ,CARTON_ID";
                    strSQL = strSQL + " ,FORCE_EMPTY";
                    strSQL = strSQL + " ,PRODUCTID";
                    strSQL = strSQL + " ,PANEL_MODE";
                    strSQL = strSQL + " ,PPBOX_GRADE";
                    strSQL = strSQL + " ,GLASS_QTY";
                    strSQL = strSQL + " ,GSFLAG";
                    strSQL = strSQL + " ,CT1FLAG";
                    strSQL = strSQL + " ,CHIP_QTY";
                    strSQL = strSQL + " ,PPBOX_COUNT";
                    strSQL = strSQL + " ,GLASS_TYPE";
                    strSQL = strSQL + " ,FACTORYID_TO";
                    strSQL = strSQL + " ,GLASSID";
                    strSQL = strSQL + " ,TFT_POSITION"; //2017-03-21 Add
                    strSQL = strSQL + " ,MES_SHIP_DATE";
                    strSQL = strSQL + " ,OXDATA";
                    strSQL = strSQL + " ,ASMTESTDATA";
                    strSQL = strSQL + " ,ENGINEERING";
                    strSQL = strSQL + " ,OXDEFECODE";
                    strSQL = strSQL + " ,QUADRANT";
                    strSQL = strSQL + " ,OXFLAG";
                    strSQL = strSQL + " ,OXGRADE";
                    strSQL = strSQL + " ,OXLEVEL";
                    strSQL = strSQL + " ,X";
                    strSQL = strSQL + " ,Y";
                    strSQL = strSQL + " ,UPDFLAG";
                    strSQL = strSQL + " ,UPDTIME";
                    #endregion
                    strSQL = strSQL + " )VALUES(";
                    #region parameters
                    strSQL = strSQL + "'" + dtOX.Rows[i]["DELIVERY_NO"].ToString() + "'";
                    strSQL = strSQL + ",'" + dtOX.Rows[i]["FACTORYID"].ToString() + "'";
                    strSQL = strSQL + ",'" + dtOX.Rows[i]["SHIPPINGFROM"].ToString() + "'";
                    strSQL = strSQL + ",'" + dtOX.Rows[i]["SHIPPINGTO"].ToString() + "'";
                    strSQL = strSQL + ",'" + dtOX.Rows[i]["SHIPPINGDATE"].ToString() + "'";
                    strSQL = strSQL + ",'" + dtOX.Rows[i]["PALLET_ID"].ToString() + "'";
                    strSQL = strSQL + ",'" + dtOX.Rows[i]["CARTON_ID"].ToString() + "'";
                    strSQL = strSQL + ",'" + dtOX.Rows[i]["FORCE_EMPTY"].ToString() + "'";
                    strSQL = strSQL + ",'" + dtOX.Rows[i]["PRODUCTID"].ToString() + "'";
                    strSQL = strSQL + ",'" + dtOX.Rows[i]["PANEL_MODE"].ToString() + "'";
                    strSQL = strSQL + ",'" + dtOX.Rows[i]["PPBOX_GRADE"].ToString() + "'";
                    strSQL = strSQL + "," + dtOX.Rows[i]["GLASS_QTY"];
                    strSQL = strSQL + ",'" + dtOX.Rows[i]["GSFLAG"].ToString() + "'";
                    strSQL = strSQL + ",'" + dtOX.Rows[i]["CT1FLAG"].ToString() + "'";
                    strSQL = strSQL + ",'" + dtOX.Rows[i]["CHIP_QTY"].ToString() + "'";
                    strSQL = strSQL + ",'" + dtOX.Rows[i]["PPBOX_COUNT"].ToString() + "'";
                    strSQL = strSQL + ",'" + dtOX.Rows[i]["GLASS_TYPE"].ToString() + "'";
                    strSQL = strSQL + ",'" + dtOX.Rows[i]["FACTORYID_TO"].ToString() + "'";
                    strSQL = strSQL + ",'" + dtOX.Rows[i]["GLASSID"].ToString() + "'";
                    strSQL = strSQL + ",'" + dtOX.Rows[i]["TFT_POSITION"].ToString() + "'"; //2017-03-21 Add
                    strSQL = strSQL + ",'" + dtOX.Rows[i]["MES_SHIP_DATE"].ToString() + "'";
                    strSQL = strSQL + ",'" + dtOX.Rows[i]["OXDATA"].ToString() + "'";
                    strSQL = strSQL + ",'" + dtOX.Rows[i]["ASMTESTDATA"].ToString() + "'";
                    strSQL = strSQL + ",'" + dtOX.Rows[i]["ENGINEERING"].ToString() + "'";
                    strSQL = strSQL + ",'" + dtOX.Rows[i]["OXDEFECODE"].ToString() + "'";
                    strSQL = strSQL + ",'" + dtOX.Rows[i]["QUADRANT"].ToString() + "'";
                    strSQL = strSQL + ",'" + dtOX.Rows[i]["OXFLAG"].ToString() + "'";
                    strSQL = strSQL + ",'" + dtOX.Rows[i]["OXGRADE"].ToString() + "'";
                    strSQL = strSQL + ",'" + dtOX.Rows[i]["OXLEVEL"].ToString() + "'";
                    strSQL = strSQL + ",'" + dtOX.Rows[i]["X"].ToString() + "'";
                    strSQL = strSQL + ",'" + dtOX.Rows[i]["Y"].ToString() + "'";
                    strSQL = strSQL + ",'" + dtOX.Rows[i]["UPDFLAG"].ToString() + "'";
                    strSQL = strSQL + ",CONVERT(VARCHAR(30), GETDATE(), 121) ";
                    #endregion
                    strSQL = strSQL + " )";
                    
                    int intExec = conn.ExecuteSQLServerSQL(strConn, strSQL);
                    Console.WriteLine(System.DateTime.Now.ToString("yyyyMMddHHmmss") + " 寫入MES數據；棧板號：" + strPalletID + " 箱號：" + strCartonID + " 玻璃號：" + strGlassID);
                }
                strSQL = "";
                //strSQL = strSQL + " UPDATE MFGINFO SET UPDFLAG = 'H', UPDTIME = TO_CHAR(SYSDATE,'YYYYMMDDHH24MISS')";
                //strSQL = strSQL + " WHERE PALLET_ID = '" + strPalletID + "' AND CARTON_ID = '" + strCartonID + "'  AND UPDFLAG = 'N' ";
                //strSQL = strSQL + " UPDATE MFGINFO SET UPDFLAG = CASE WHEN SHIPPINGFROM IN('T0','T1','T2') THEN UPDFLAG ELSE 'H' END";
                strSQL = strSQL + " UPDATE MFGINFO SET UPDTIME = TO_CHAR(SYSDATE,'YYYYMMDDHH24MISS')";
                strSQL = strSQL + " , UPDUSER = 'NJMES'";
                //strSQL = strSQL + " WHERE PALLET_ID = '" + strPalletID + "' AND CARTON_ID = '" + strCartonID + "'  AND UPDFLAG = 'N' ";                            //2018-07-24 MARK
                strSQL = strSQL + " WHERE PALLET_ID = '" + strPalletID + "' AND CARTON_ID = '" + strCartonID + "'  AND (UPDFLAG = 'N' OR DELIVERY_NO LIKE '007%')";  //2018-07-24 ADD
                int intExecOracle = conn.ExecuteOracleSQL(strConnOX, strSQL);
                if (intExecOracle == 0)
                {
                    Console.WriteLine(System.DateTime.Now.ToString("yyyyMMddHHmmss") + " 更新收取過的OX狀態；棧板號：" + strPalletID + "；箱號：" + strCartonID + "；無可處理的數據");
                }
                else
                {
                    Console.WriteLine(System.DateTime.Now.ToString("yyyyMMddHHmmss") + " 更新收取過的OX狀態；棧板號：" + strPalletID + "；箱號：" + strCartonID + "；共" + intExecOracle.ToString() + "筆數據處理成功");
                }
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
