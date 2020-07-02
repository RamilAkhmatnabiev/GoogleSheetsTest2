using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Npgsql;
using System.Configuration;
using System.Collections.Specialized;




namespace GoogleSheetsTest2
{
    class Program
    {



        static void Main(string[] args)
        {
            List<string[]> listOfServerInfo;
            NameValueCollection serversFromConfMngr = ConfigurationManager.AppSettings;
            ServerManipulator sm;
            DataRecorder dataRecorder;
            string spreadSheetsId = "18bjPMlVNxm7yQ0Rg1weso9_db6Rg6NrfHgpFj2S-u7s"; // вставить сюда айди таблицы
            int countOfServers = serversFromConfMngr.AllKeys.Length;
            int refreshesCount = 0;
            int timeOfRefresh = 1000;




            Console.WriteLine("Set the time between the refreshes in ms:");
            timeOfRefresh = int.Parse(Console.ReadLine());

            if (countOfServers > 0)
            {
                dataRecorder = new DataRecorder(countOfServers, spreadSheetsId);

                while (true)
                {
                    for (int i = 0; i < countOfServers; i++)
                    {
                        sm = new ServerManipulator(ConfigurationManager.AppSettings.Get(i));

                        sm.DoConnection();
                        listOfServerInfo = sm.GetDatabaseInfoForGoogleSheets();

                        dataRecorder.FillSpreadSheets(listOfServerInfo, dataRecorder.GetThisIndexSheetId(i));

                        sm.DoDesconnection();
                    }

                    System.Threading.Thread.Sleep(timeOfRefresh);
                    Console.WriteLine("{0}#Refresh the table", refreshesCount);
                    refreshesCount++;
                }


            }
        }

    }
}
