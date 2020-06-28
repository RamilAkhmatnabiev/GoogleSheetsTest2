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

namespace GoogleSheetsTest2
{
    class Program
    {
        private static string ClientSecret = "client_secret.json";
        private static readonly string[] ScopeSheets = { SheetsService.Scope.Spreadsheets }; // права на использование
        private static readonly string AppName = "ProgramForPostgressTest"; // имя приложения
        private static readonly string SpreadSheetsId = "18bjPMlVNxm7yQ0Rg1weso9_db6Rg6NrfHgpFj2S-u7s"; // айди таблицы
        private const string Range = "'Sheet1' A1:F"; // диапазон получаемых ячеек строки
        private static string[,] Data = new string[,]
        {
            {"11","12","14"},
            {"11","12","14"},
            {"11","12","14"}
        };



        static void Main(string[] args)
        {
            Console.WriteLine("Get credentials");
            var credential = GetSheetCredential(ClientSecret); // Два метода из QuickstartSheet с получением credntial и service
                                                               // Параметром передаем файл client_secret.json

            Console.WriteLine("Get servise");
            var service = GetService(credential); // connecting to google



            Console.WriteLine("Fill Data");
            FillSpreadSheets(service, SpreadSheetsId, Data); // filling sheet with Data

            Console.WriteLine("Get result");
            //string result = GetFirstCell(service, Range, SpreadSheetsId);
            //Console.WriteLine("result : {0}", result);

            Console.WriteLine("Done");
            Console.ReadLine();

        }

        

        private static UserCredential GetSheetCredential(string ClientSecret)
        {

            using (var stream =
                new FileStream(ClientSecret, FileMode.Open, FileAccess.Read))
            {
                string credPath = "token.json";
                return GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                                ScopeSheets,
                                "user",
                                CancellationToken.None,
                                new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }
        }
        private static SheetsService GetService(UserCredential credential)
        {
            return new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = AppName,
            });
        }

        private static void FillSpreadSheets(SheetsService service, string spreadSheetsId, string[,] data)
        {
            List<Request> requests = new List<Request>(); // создаем массив запросов

            for (int i = 0; i < data.GetLength(0); i++)
            {
                List<CellData> values = new List<CellData>(); //создаем массив значний

                for (int j = 0; j < data.GetLength(1); j++)
                {
                    values.Add(new CellData
                    {
                        UserEnteredValue = new ExtendedValue
                        {
                            StringValue = data[i, j]
                        }
                    }
                    );
                }

                requests.Add(
                    new Request
                    {
                        UpdateCells = new UpdateCellsRequest
                        {
                            Start = new GridCoordinate
                            {
                                SheetId = 0,
                                RowIndex = i,
                                ColumnIndex = 0
                            },
                            Rows = new List<RowData> { new RowData { Values = values } },
                            Fields = "userEnteredValue"
                        }
                    });

            }

            BatchUpdateSpreadsheetRequest busr = new BatchUpdateSpreadsheetRequest
            {
                Requests = requests
            };
            service.Spreadsheets.BatchUpdate(busr, spreadSheetsId).Execute();
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
