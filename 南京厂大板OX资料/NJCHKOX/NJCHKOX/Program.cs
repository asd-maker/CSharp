using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Data;

namespace NJCHKOX
{
    class Program
    {
        static void Main(string[] args)
        {
            ChkOXClass ec = new ChkOXClass();
            /*--------------------------------------------------------------------------------------------------------------------------------------------------------------------*/
            /*判斷程式是否執行中，執行中終止再次執行程式*/
            string strAppName = System.Diagnostics.Process.GetCurrentProcess().ProcessName;
            System.Diagnostics.Process[] Process = System.Diagnostics.Process.GetProcessesByName(strAppName);
            if (Process.Length > 1)
            {
                Console.WriteLine("程式已經開啟，請勿重複運行");
                Console.ReadLine();
                return;
            }
            Console.WindowHeight = 40;
            Console.WindowWidth = 120;
            /*--------------------------------------------------------------------------------------------------------------------------------------------------------------------*/
            Console.WriteLine(System.DateTime.Now.ToString("yyyyMMddHHmmss") + " " + "程式開始運行!");
            /*--------------------------------------------------------------------------------------------------------------------------------------------------------------------*/
            ec.ReadConfig();
            /*--------------------------------------------------------------------------------------------------------------------------------------------------------------------*/
            ChkOXData();
            /*--------------------------------------------------------------------------------------------------------------------------------------------------------------------*/

        }

        private static void ChkOXData()
        {
            Thread.Sleep(ChkOXClass.CintInterval * 1000);
            do
            {
                ChkOXClass ec = new ChkOXClass();
                DateTime dtNow = DateTime.Now;
                if (Convert.ToInt32(dtNow.ToString("mm")) >= 55 && Convert.ToInt32(dtNow.ToString("mm")) <= 57)
                {
                    bool Flg = false;
                    Console.WriteLine(System.DateTime.Now.ToString("yyyyMMddHHmmss") + " 開始處理數據");
                    DataTable dtCartonID = new DataTable();
                    dtCartonID = ec.GetCartonData();
                    if (dtCartonID.Rows.Count > 0)
                    {
                        string strNCartonID = string.Empty;
                        for (int i = 0; i < dtCartonID.Rows.Count; i++)
                        {
                            string strCartonID = dtCartonID.Rows[i]["trackno"].ToString();

                            if (!ec.ChkCartonData(strCartonID))
                            {
                                Flg = true;
                                strNCartonID += strCartonID + "     ";
                                Console.WriteLine(System.DateTime.Now.ToString("yyyyMMddHHmmss") + "此箱號：" + strCartonID + "無OX信息");
                            }
                            else
                            {
                                Console.WriteLine(System.DateTime.Now.ToString("yyyyMMddHHmmss") + "此箱號：" + strCartonID + "已接收OX資料");
                            }
                        }
                        if (Flg)
                        {
                            if (ec.SendMail(ChkOXClass.strSMTP, "NJGetOX@Innolux.com", ChkOXClass.strMailTo, "NJ【◆警告◆：NJCHKOX檢查出無OX信息箱號，請值班人員確認處理】", "箱號：" + strNCartonID,"") == false)
                            {
                                Console.WriteLine(System.DateTime.Now.ToString("yyyyMMddHHmmss") + " 郵件發送失敗");
                            }
                        }
                        Console.WriteLine(System.DateTime.Now.ToString("yyyyMMddHHmmss") + " 結束處理數據");
                    }
                    else
                    {
                        Console.WriteLine(System.DateTime.Now.ToString("yyyyMMddHHmmss") + " 無數據可處理，結束處理數據");
                    }
                }
                else
                {
                    Console.WriteLine(System.DateTime.Now.ToString("yyyyMMddHHmmss") + " 每小時的55分~57分時間段內才執行，其他時間不做處理");
                }
                Thread.Sleep(ChkOXClass.CintInterval * 1000);
            } while (0 < 1);
        }
    }
}
