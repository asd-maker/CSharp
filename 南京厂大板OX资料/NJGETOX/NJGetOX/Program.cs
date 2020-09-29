using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Data;

namespace NJGetOX
{
    class Program
    {
        static Timer tm;
        static void Main(string[] args)
        {
            GetOXClass ec = new GetOXClass();
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
            ExecOXData();
            /*--------------------------------------------------------------------------------------------------------------------------------------------------------------------*/
        }
        static void ExecOXData()
        {
            Thread.Sleep(GetOXClass.CintInterval * 1000);
            do
            {
                GetOXClass ec = new GetOXClass();
                DateTime dtNow = DateTime.Now;
                if (Convert.ToInt32(dtNow.ToString("mm")) >= 1 && Convert.ToInt32(dtNow.ToString("mm")) <= 3)
                {
                    Console.WriteLine(System.DateTime.Now.ToString("yyyyMMddHHmmss") + " 開始處理數據");
                    DataTable dtPalletID = new DataTable();
                    dtPalletID = ec.GetPalletData();
                    if (dtPalletID.Rows.Count > 0)
                    {
                        for (int i = 0; i < dtPalletID.Rows.Count; i++)
                        {
                            string strPalletID = dtPalletID.Rows[i]["PALLET_ID"].ToString();
                            string strCartonID = dtPalletID.Rows[i]["CARTON_ID"].ToString();

                            DataTable dtGetOX = ec.GetOXData(strPalletID, strCartonID);
                            if (dtGetOX.Rows.Count > 0)
                            {
                                bool blOX = ec.InsertOXData(strPalletID, strCartonID, dtGetOX);
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
                    Console.WriteLine(System.DateTime.Now.ToString("yyyyMMddHHmmss") + " 每小時的1分~3分時間段內才執行，其他時間不做處理");
                }
                Thread.Sleep(GetOXClass.CintInterval * 1000);
            } while (0 < 1);
        }
    }
}
