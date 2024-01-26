using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Management;
using System.Text;
using System.Web;
using System.Web.Mvc;

namespace PoolWorkWeb.Controllers
{
    public class HomeController : Controller
    {


        public ActionResult Index()
        {
            return View();
        }


        [HttpPost]
        public ActionResult Index(HttpPostedFileBase upload, string pool)
        {
            if (upload != null)
            {
                // получаем имя файла
                string fileName = System.IO.Path.GetFileName(upload.FileName);
                string path = "c:\\pubdocs\\" + fileName;
                // сохраняем файл в папку Files в проекте
                upload.SaveAs(path);

                // путь файлов проекта
                String extractPath = "";
                if (pool == "OldGeo")
                    extractPath = @"C:/inetpub/wwwroot3/oldgeo";
                else if (pool == "nuska")
                    extractPath = @"C:/inetpub/nuska";
                else if (pool == "DefaultAppPool")
                    extractPath = @"C:/inetpub/wwwroot2/analiz";
                else if (pool == "ananizapi")
                    extractPath = @"C:/inetpub/wwwroot/analiz_site";
                else if (pool == "LogPool")
                    extractPath = @"C:/inetpub/logapp";
                else if (pool == "HSCApp")
                    extractPath = @"C:/inetpub/wwwroot/HSCApp";

                // backup 
                ZipFile.CreateFromDirectory(extractPath, "C:\\inetpub\\!backups\\" + pool + "-" + DateTime.Now.ToString(format: "HH-dd") + ".zip");

                // stop pool
                performRequestedAction("localhost", pool, "stop");

                // Удалить все старые файлы
                System.IO.DirectoryInfo di = new DirectoryInfo(extractPath);
                foreach (FileInfo file in di.EnumerateFiles())
                    file.Delete();
                foreach (DirectoryInfo dir in di.GetDirectories())
                    if(dir.Name != "wwwroot")
                        dir.Delete(true);

                // Извлечим архив файл
                ZipFile.ExtractToDirectory(path, extractPath);
                // start pool
                performRequestedAction("localhost", pool, "start");
            }
            return RedirectToAction("About");
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        private static bool IsApplicationPoolRunning(string servername, string strAppPool)
        {
            string sb = ""; // String to store return value

            // Connection options for WMI object
            ConnectionOptions options = new ConnectionOptions();

            // Packet Privacy means authentication with encrypted connection.
            options.Authentication = AuthenticationLevel.PacketPrivacy;

            // EnablePrivileges : Value indicating whether user privileges 
            // need to be enabled for the connection operation. 
            // This property should only be used when the operation performed 
            // requires a certain user privilege to be enabled.
            options.EnablePrivileges = true;

            // Connect to IIS WMI namespace \\root\\MicrosoftIISv2
            ManagementScope scope = new ManagementScope(@"\\" + servername + "\\root\\MicrosoftIISv2", options);

            // Query IIS WMI property IISApplicationPoolSetting
            ObjectQuery oQueryIISApplicationPoolSetting = new ObjectQuery("SELECT * FROM IISApplicationPoolSetting");

            // Search and collect details thru WMI methods
            ManagementObjectSearcher moSearcherIISApplicationPoolSetting =
                new ManagementObjectSearcher(scope, oQueryIISApplicationPoolSetting);
            ManagementObjectCollection collectionIISApplicationPoolSetting =
                            moSearcherIISApplicationPoolSetting.Get();

            // Loop thru every object
            foreach (ManagementObject resIISApplicationPoolSetting
                in collectionIISApplicationPoolSetting)
            {
                // IISApplicationPoolSetting has a property called Name which will 
                // return Application Pool full name /W3SVC/AppPools/DefaultAppPool
                // Extract Application Pool Name alone using Split()
                if (resIISApplicationPoolSetting
                    ["Name"].ToString().Split('/')[2] == strAppPool)
                {
                    // IISApplicationPoolSetting has a property 
                    // called AppPoolState which has following values
                    // 2 = started 4 = stopped 1 = starting 3 = stopping
                    if (resIISApplicationPoolSetting["AppPoolState"].ToString() != "2")
                    {
                        return false;
                    }
                }
            }
            return true;
        }

        public static void performRequestedAction(String servername, String AppPoolName, String action)
        {
            StringBuilder sb = new StringBuilder();
            ConnectionOptions options = new ConnectionOptions();
            options.Authentication = AuthenticationLevel.PacketPrivacy;
            options.EnablePrivileges = true;
            ManagementScope scope = new ManagementScope(@"\\" + servername + "\\root\\MicrosoftIISv2", options);

            // IIS WMI object IISApplicationPool to perform actions on IIS Application Pool
            ObjectQuery oQueryIISApplicationPool = new ObjectQuery("SELECT * FROM IISApplicationPool");

            ManagementObjectSearcher moSearcherIISApplicationPool =
                new ManagementObjectSearcher(scope, oQueryIISApplicationPool);
            ManagementObjectCollection collectionIISApplicationPool =
                moSearcherIISApplicationPool.Get();
            foreach (ManagementObject resIISApplicationPool in collectionIISApplicationPool)
            {
                if (resIISApplicationPool["Name"].ToString().Split('/')[2] == AppPoolName)
                {
                    // InvokeMethod - start, stop, recycle can be passed as parameters as needed.
                    resIISApplicationPool.InvokeMethod(action, null);
                }
            }
        }
    }
}