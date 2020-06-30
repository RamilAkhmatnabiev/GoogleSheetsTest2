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

namespace GoogleSheetsTest2
{
    class Program
    {
        //private static string ClientSecret = "client_secret.json";
        //private static readonly string[] ScopeSheets = { SheetsService.Scope.Spreadsheets }; // права на использование
        //private static readonly string AppName = "ProgramForPostgressTest"; // имя приложения
        //private static readonly string SpreadSheetsId = "18bjPMlVNxm7yQ0Rg1weso9_db6Rg6NrfHgpFj2S-u7s"; // айди таблицы
        //private const string Range = "'Sheet1' A1:F"; // диапазон получаемых ячеек строки
        



        static void Main(string[] args)
        {
            List<string[]> listOfServerInfo;
            ServerManipulator sm = new ServerManipulator();
            sm.DoConnection();
            listOfServerInfo =  sm.GetDatabaseInfoForGoogleSheets();

            DataRecorder dataRecorder = new DataRecorder();
            dataRecorder.FillSpreadSheets(listOfServerInfo);


            sm.DoDesconnection();


            
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
