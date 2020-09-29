using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data;
using System.Threading.Tasks;
using System.Reflection;
using System.Data.OleDb;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Globalization;
using System.Threading;
using System.Runtime.InteropServices;
using System.Net;
using System.Diagnostics;
using System.Windows.Forms;

namespace W9206B
{

    class Program
    {
        static int count = 0;
        static string strPath = "";
        static string strLogPath = "";
        static string strFabid = "";
        static string strSublineID = "";
        static string strProdid = "";
        static string strIP = "";
        static string strHostName = "";

        static void Main(string[] args)
        {
            //隱藏控制臺，後臺顯示
            Console.Title = "解析9206B文件";
            ConsoleHelper.hideConsole("解析9206B文件");
            //獲取IP和主機名
            strHostName = Dns.GetHostName();
            System.Net.IPAddress[] addressList = Dns.GetHostEntry(Dns.GetHostName()).AddressList;
            for (int i = 0; i < addressList.Length; i++)
            {
                strIP = addressList[i].ToString();
            }
            Console.WriteLine("IP:" + strIP + "\n");
            Console.WriteLine("HostName:" + strHostName);
            //防止多次開啓該程式
            Process proc = RunningInstance();
            if (proc != null) { MessageBox.Show("請不要重複開啓軟件！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }

            //主程序運行
            try
            {
                //讀取配置檔
                string strAppPath = System.IO.Directory.GetCurrentDirectory() + "\\Path.txt";
                bool isPath = FileCheck(strAppPath);
                if (isPath)
                {
                    Console.WriteLine("路徑文件創建成功！");
                }
                else
                {
                    Console.WriteLine("路徑文件創建失敗！");
                    Console.ReadKey();
                }
                string[] strTxt = GetFilePath(strAppPath);
                strPath = strTxt[0];
                strFabid = strTxt[1];
                strSublineID = strTxt[2];
                strProdid = strTxt[3];
                string strMfgDate = DateTime.Now.ToString("yyyyMMdd");
                strPath = strPath + strMfgDate + @"\";

                //循環執行
                while (true)
                {
                    //檢查待解析文件路徑
                    DirectoryInfo dr = new DirectoryInfo(strPath);
                    if (!dr.Exists)
                    {
                        Console.WriteLine("當前不存在此目錄，" + strPath.ToString() + "即將退出本次解析");
                        Thread.Sleep(60 * 1000);
                        continue;
                    }
                    //檢查Log文件路徑
                    strLogPath = strPath + "Log.txt";
                    FileInfo fi = new FileInfo(strLogPath);
                    if (!fi.Exists)
                    {
                        fi.Create();
                    }
                    Console.WriteLine("當前目錄{0}", strPath);
                    Console.WriteLine("資料讀取中...");
                    try
                    {
                        //解析文件
                        ListFiles(new DirectoryInfo(strPath));
                    }
                    catch (IOException e)
                    {
                        File.AppendAllText(strLogPath, "\r\n" + DateTime.Now.ToString("yyyyMMddhhmmss") + e.Message);
                        Console.WriteLine(e.Message);
                    }
                    Console.WriteLine("當前時間處理數據" + count.ToString() + "筆");
                    Thread.Sleep(60 * 1000);//休眠60s
                    count = 0;
                }
            }
            catch (Exception ex)
            {
                File.AppendAllText(strLogPath, "\r\n" + DateTime.Now.ToString("yyyyMMddhhmmss") + ex.Message);
                Console.WriteLine(ex.Message);

            }

        }
        /// <summary>
        /// 遍歷文件夾中的文件
        /// </summary>
        /// <param name="info"></param>
        public static void ListFiles(FileSystemInfo info)
        {
            DataSet dt = new DataSet();
            if (!info.Exists)
                return;
            DirectoryInfo dir = info as DirectoryInfo;
            if (dir == null)
                return;
            FileSystemInfo[] files = dir.GetFileSystemInfos();

            for (int i = 0; i < files.Length; i++)
            {
                FileInfo file = files[i] as FileInfo;
                if (file != null)
                {
                    //排除備份目錄
                    if (!file.Directory.ToString().Contains("_Backup") & !file.ToString().Contains("Log"))
                        //读取TXT资料
                        ReadFileToSql(file.FullName.ToString(), file.Directory.ToString() + "_Backup\\");
                }

                else
                    //对于子目录，进行递归调用 
                    ListFiles(files[i]);
            }
        }
        /// <summary>
        /// 讀取TXT文檔名後寫入數據庫
        /// </summary>
        /// <param name="strFilePath"></param>
        /// <param name="strCpPath"></param>
        private static void ReadFileToSql(string strFilePath, string strCpPath)
        {
            string[] FilePath = strFilePath.Split('\\');
            string FileName = FilePath[FilePath.Length - 1];
            string TestResult = FilePath[FilePath.Length - 2];
            DirectoryInfo dr = new DirectoryInfo(strCpPath);
            //檢查備份目錄是否存在
            if (!dr.Exists & (dr.ToString().Contains("Pass") | dr.ToString().Contains("Fail")))
            {
                dr.Create();
            }

            if (FileName.Contains("txt"))
            {
                //Pass|Fail結果轉換
                if (TestResult == "Pass")
                {
                    TestResult = "OK";
                }
                else
                {
                    TestResult = "NG";
                }

                string strSN = FileName.Split('_')[0];


                DataSet ds = new DataSet();
                SqlConnection sqlconnt = new SqlConnection();
                if (strFabid.ToUpper().Equals("A"))
                {
                     sqlconnt = new SqlConnection(@"Data Source=10.64.3.81\meswip;Initial Catalog=LCM1PWIP;Persist Security Info=True;User ID=meswip;password=h@ppy2008;Max Pool Size=10;Min Pool Size=5;Pooling=True;");
                }
                else 
                {
                     sqlconnt = new SqlConnection(@"Data Source=10.64.3.101\meswip;Initial Catalog=LCM2PWIP;Persist Security Info=True;User ID=meswip2;password=lcm2_2009;Max Pool Size=10;Min Pool Size=5;Pooling=True;");
                }
                StringBuilder str = new StringBuilder();
                str.Append("SELECT TOP 1 FRSN FROM dbo.WSMSX WHERE FABID='" + strFabid + "'AND TOSN='" + strSN + "'AND PFLG='0'ORDER BY TIME_STAMP DESC");
                SqlCommand sqlCommand = new SqlCommand();
                sqlCommand.Connection = sqlconnt;
                sqlCommand.CommandType = CommandType.Text;
                sqlCommand.CommandText = str.ToString();
                sqlCommand.CommandTimeout = 0;
                SqlDataAdapter sqlDataAdapter = new SqlDataAdapter();
                sqlDataAdapter.SelectCommand = sqlCommand;
                sqlconnt.Open();
                sqlDataAdapter.Fill(ds);
                if (ds.Tables[0] == null)
                {
                    return;
                }
                if (ds.Tables[0].Rows.Count > 0)
                {
                    strSN = ds.Tables[0].Rows[0][0].ToString();
                }

                str.Clear();
                ds.Clear();
                //檢查是否資料存在
                str.Append("SELECT RESULT FROM W9206B WITH (NOLOCK) WHERE FABID='A'AND SN='" + strSN + "'");
                sqlCommand.CommandText = str.ToString();
                sqlDataAdapter.Fill(ds);
                if (ds.Tables[0].Rows.Count > 0)
                {
                    //存在即更新
                    str.Append("UPDATE W9206B SET RESULT='" + TestResult + "',Time_Stamp=GETDATE() WHERE FABID='A'AND SN='" + strSN + "'");
                    sqlCommand.CommandText = str.ToString();
                }
                else
                {
                    //不存在即插入
                    str.Append("INSERT INTO W9206B ( Fabid ,SN ,Prodid ,SublineID ,Workctr ,Result ,UserID ,Time_Stamp ,opt1 ,opt2 ,opt3) VALUES  ('" + strFabid + "','" + strSN + "','" + strProdid + "','" + strSublineID + "','P430','" + TestResult + "','AUTO',getdate(),'"+strIP+"','"+strHostName+"','')");
                    sqlCommand.CommandText = str.ToString();
                }
                sqlCommand.ExecuteNonQuery();
                sqlconnt.Close();
                count = count + 1;//計算更新數量
                try
                {
                    //文件備份並刪除
                    File.Copy(strFilePath, strCpPath + FileName);
                    File.Delete(strFilePath);
                }
                // 捕捉异常.
                catch (IOException copyError)
                {
                    File.AppendAllText(strLogPath, "\r\n" + DateTime.Now.ToString("yyyyMMddhhmmss") + copyError.Message);
                    Console.WriteLine(copyError.Message);
                    File.Delete(strFilePath);
                }
            }
        }
        /// <summary>
        /// 讀取Path.txt配置文件
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static string[] GetFilePath(String filePath)
        {
            string[] strData = new string[4];
            try
            {
                StreamReader sr = new StreamReader(filePath, Encoding.Default);
                string line = "";
                int i = 0;
                while ((line = sr.ReadLine()) != null)
                {
                    strData[i] = line.ToString();
                    i = i + 1;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("路徑文件讀取失敗！！！");
                Console.WriteLine(e.Message);
            }
            return strData;
        }
        /// <summary>
        /// 判斷文件是否存在
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static bool FileCheck(string fileName)
        {
            bool flag = false;
            FileInfo fileInfo = new FileInfo(fileName);
            //存在文件
            if (fileInfo.Exists)
            {
                flag = true;
            }
            //不存在，新增文件
            else
            {
                FileStream fileStream = fileInfo.Create();
                fileStream.Close();
                flag = true;
            }
            return flag;
        }
        /// <檢查程序是否已開啓>
        /// 檢查程序是否已開啓
        /// </summary>
        /// <returns></returns>
        private static Process RunningInstance()
        {
            Process[] processes = Process.GetProcesses();  //进程列表
            Process current = Process.GetCurrentProcess(); //当前进程
                                                           //当前进程名称(xxx.vshost.exe或xxx.exe)
            string currentName = Path.GetFileName(Assembly.GetCallingAssembly().Location.Replace("/", "\\\\"));
            currentName = currentName.Split(new char[] { '.' })[0];//只取主文件名(xxx)
            foreach (Process process in processes)
            {
                string procName = process.ProcessName;
                bool vshost = procName.Contains("W9206B");
                procName = procName.Split(new char[] { '.' })[0];
                if (vshost)
                {
                    if (process.Id != current.Id)
                    {
                        return process;
                    }
                }
                //if (!vshost && string.Compare(procName, currentName, true) == 0 && process.Id != current.Id) return process;
            }
            return null;
        }
        public static class ConsoleHelper
        {
            /// <summary>  
            /// 获取窗口句柄  
            /// </summary>  
            /// <param name="lpClassName"></param>  
            /// <param name="lpWindowName"></param>  
            /// <returns></returns>  
            [DllImport("user32.dll", SetLastError = true)]
            private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

            /// <summary>  
            /// 设置窗体的显示与隐藏  
            /// </summary>  
            /// <param name="hWnd"></param>  
            /// <param name="nCmdShow"></param>  
            /// <returns></returns>  
            [DllImport("user32.dll", SetLastError = true)]
            private static extern bool ShowWindow(IntPtr hWnd, uint nCmdShow);

            /// <summary>  
            /// 隐藏控制台  
            /// </summary>  
            /// <param name="ConsoleTitle">控制台标题(可为空,为空则取默认值)</param>  
            public static void hideConsole(string ConsoleTitle = "")
            {
                ConsoleTitle = String.IsNullOrEmpty(ConsoleTitle) ? Console.Title : ConsoleTitle;
                IntPtr hWnd = FindWindow("ConsoleWindowClass", ConsoleTitle);
                if (hWnd != IntPtr.Zero)
                {
                    ShowWindow(hWnd, 0);
                }
            }

            /// <summary>  
            /// 显示控制台  
            /// </summary>  
            /// <param name="ConsoleTitle">控制台标题(可为空,为空则去默认值)</param>  
            public static void showConsole(string ConsoleTitle = "")
            {
                ConsoleTitle = String.IsNullOrEmpty(ConsoleTitle) ? Console.Title : ConsoleTitle;
                IntPtr hWnd = FindWindow("ConsoleWindowClass", ConsoleTitle);
                if (hWnd != IntPtr.Zero)
                {
                    ShowWindow(hWnd, 1);
                }
            }
            /// <summary>  
            /// 最小化窗口 
            /// </summary>  
            /// <param name="ConsoleTitle">控制台标题(可为空,为空则去默认值)</param>  
            public static void HideConsole2(string ConsoleTitle = "")
            {
                ConsoleTitle = String.IsNullOrEmpty(ConsoleTitle) ? Console.Title : ConsoleTitle;
                IntPtr hWnd = FindWindow("ConsoleWindowClass", ConsoleTitle);
                if (hWnd != IntPtr.Zero)
                {
                    ShowWindow(hWnd, 2);
                }
            }

        }
    }
}
