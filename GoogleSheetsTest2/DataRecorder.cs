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
    public class DataRecorder
    {
        private string ClientSecret;
        //private string ClientSecret = "client_secret.json";
        private readonly string[] ScopeSheets = { SheetsService.Scope.Spreadsheets }; // права на использование
        private readonly string AppName = "ProgramForPostgressTest"; // имя приложения

        //private string AppName; // имя приложения
        //private static readonly string SpreadSheetsId = "18bjPMlVNxm7yQ0Rg1weso9_db6Rg6NrfHgpFj2S-u7s"; // айди таблицы
        private string spreadSheetsId; // айди таблицы
        private string Range; // диапазон получаемых ячеек строки

        private UserCredential credential;// нужен для хранения credential
        private SheetsService service;  // нужен для хранения service





        public DataRecorder(string spreadSheetsId = "18bjPMlVNxm7yQ0Rg1weso9_db6Rg6NrfHgpFj2S-u7s", string range = "'Sheet1' A1:F")
        {
            this.spreadSheetsId = spreadSheetsId;
            this.Range = range;
            ClientSecret = "client_secret.json";

            credential = GetSheetCredential(ClientSecret); // формируем credential на базе файла client secret

            service = GetService(credential); // подключаемся к googlt  с credential  и получаем service


        }

        public void FillSpreadSheets(List<string[]> info)  
        {
            List<Request> requests = new List<Request>(); // создаем массив запросов

            for (int i = 0; i < info.Count; i++)
            {
                List<CellData> values = new List<CellData>(); //создаем массив значний

                for (int j = 0; j < 2; j++)
                {
                    values.Add(new CellData
                    {
                        UserEnteredValue = new ExtendedValue
                        {
                            StringValue = info[i][j]
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
            //service.Spreadsheets.BatchUpdate(busr, spreadSheetsId).Execute();
            service.Spreadsheets.BatchUpdate(busr, spreadSheetsId).Execute();

        }

        private UserCredential GetSheetCredential(string ClientSecret)
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


        private SheetsService GetService(UserCredential credential)
        {
            return new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = AppName,
            });
        }



    }
}
