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
            //NameValueCollection all = ConfigurationManager.AppSettings;
            //foreach (string s in all.AllKeys) Console.WriteLine("Key: " + s + " Value: " + all.Get(s)); Console.ReadLine();
            
            List<string[]> listOfServerInfo;
            NameValueCollection serversFromConfMngr = ConfigurationManager.AppSettings;
            ServerManipulator sm;
            DataRecorder dataRecorder;
            string spreadSheetsId = "18bjPMlVNxm7yQ0Rg1weso9_db6Rg6NrfHgpFj2S-u7s"; // вставить сюда айди таблицы







            for (int i = 0; i < serversFromConfMngr.AllKeys.Length; i++)
            {
                sm = new ServerManipulator(ConfigurationManager.AppSettings.Get(i));

                sm.DoConnection();
                listOfServerInfo = sm.GetDatabaseInfoForGoogleSheets();

                dataRecorder = new DataRecorder(i, spreadSheetsId);

                dataRecorder.FillSpreadSheets(listOfServerInfo);


                //string apset = ConfigurationManager.AppSettings.Get(0);

                sm.DoDesconnection();
            }


            //foreach (string s in serversFromConfMngr)
            //{
            //    sm = new ServerManipulator(s);
            //    sm.DoConnection();
            //    listOfServerInfo = sm.GetDatabaseInfoForGoogleSheets();

            //    dataRecorder = new DataRecorder(currentServerNumber);

            //    dataRecorder.FillSpreadSheets(listOfServerInfo);


            //    //string apset = ConfigurationManager.AppSettings.Get(0);

            //    sm.DoDesconnection();
            //    currentServerNumber++;
            //}





            //ServerManipulator sm = new ServerManipulator();
            //sm.DoConnection();
            //listOfServerInfo =  sm.GetDatabaseInfoForGoogleSheets();

            //DataRecorder dataRecorder = new DataRecorder();
            //dataRecorder.FillSpreadSheets(listOfServerInfo);


            //string apset = ConfigurationManager.AppSettings.Get(0);
            //
            //sm.DoDesconnection();




        }

        

        

        //private static string GetFirstCell(SheetsService service, string range, string spreadSheetsId)
        //{
        //    SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(spreadSheetsId, range);
        //    ValueRange response = request.Execute();

        //    string result = null;
        //    foreach (var value in response.Values)
        //    {
        //        result += " " + value[0];
        //    }
        //    return result;
        //}


    }
}
