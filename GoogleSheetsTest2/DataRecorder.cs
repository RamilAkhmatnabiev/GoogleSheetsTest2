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
        //реализовать удаление листов с совпадающим именем сервева

        private string ClientSecret = "client_secret.json";
        private readonly string[] ScopeSheets = { SheetsService.Scope.Spreadsheets }; // права на использование
        private readonly string AppName = "ProgramForPostgressTest"; // имя приложения
        private string spreadSheetsId; // айди таблицы
        private int count = 0; // счетчик для sheet title name
       
        //public bool isFirstLoop { get; set; } = true;  // нужен для того, чтобы не создавать листы при следующих прогонах программы
        private List<Tuple<int, string>> allSheets = new List<Tuple<int, string>>(); // массив с парами знчений id и name всех листов таблицы

        private UserCredential credential;// нужен для хранения credential
        private SheetsService service;  // нужен для хранения service

        



        //18bjPMlVNxm7yQ0Rg1weso9_db6Rg6NrfHgpFj2S-u7s
        //10dFGbsi3Av-sqViDzCkxRaqE9TPFTozbwaZN_sQdhEo

        public DataRecorder(int serversCount, string spreadSheetsId = "18bjPMlVNxm7yQ0Rg1weso9_db6Rg6NrfHgpFj2S-u7s")
        {
            this.spreadSheetsId = spreadSheetsId;

            credential = GetSheetCredential(ClientSecret); // формируем credential на базе файла client secret
            service = GetService(credential); // подключаемся к google  с credential  и получаем service

            for (int i = 0; i < serversCount; i++)
            {
               
                    CreateNewSheet();
            }

        }

        public void FillSpreadSheets(List<string[]> info, int sheetId)  //заполняет таблицу данными
        {
            List<Request> requests = new List<Request>(); // создаем массив запросов


            for (int i = 0; i < info.Count; i++)
            {
                List<CellData> values = new List<CellData>(); //создаем массив значний

                for (int j = 0; j < info[i].Length; j++)
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
                values.Add(new CellData
                {
                    UserEnteredValue = new ExtendedValue
                    {
                        StringValue = DateTime.Now.ToString()
                    }
                }
                    );
                requests.Add(
                    new Request
                    {
                        UpdateCells = new UpdateCellsRequest
                        {
                            Start = new GridCoordinate
                            {
                                SheetId = sheetId,
                                RowIndex = i + 1,
                                ColumnIndex = 1
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

        public int GetThisIndexSheetId(int index)
        {
            return allSheets[index].Item1;
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

        private void CreateNewSheet()
        {
            string name = "server" + (count + 1).ToString();
            int id;

            var addSheetRequest = new Request
            {
                AddSheet = new AddSheetRequest
                {
                    Properties = new SheetProperties
                    {
                        //SheetId = currentServerNumber,
                        Title = name
                    }
                }
            };

            List<Request> requests = new List<Request> { addSheetRequest };
            BatchUpdateSpreadsheetRequest batchUpdateRequest = new BatchUpdateSpreadsheetRequest();
            batchUpdateRequest.Requests = requests;
            service.Spreadsheets.BatchUpdate(batchUpdateRequest, spreadSheetsId).Execute();

            id = GetLastListId();
            allSheets.Add(new Tuple<int, string>(id, name));
            count++;
        }

        private void DeleteSheet(int deletedSheetId)
        {


            var deleteSheetRequest = new Request { DeleteSheet = new DeleteSheetRequest { SheetId = deletedSheetId } };

            List<Request> requests = new List<Request> { deleteSheetRequest };
            BatchUpdateSpreadsheetRequest batchUpdateRequest = new BatchUpdateSpreadsheetRequest();
            batchUpdateRequest.Requests = requests;
            service.Spreadsheets.BatchUpdate(batchUpdateRequest, spreadSheetsId).Execute();
        }

        private Tuple<string, int> GetListIdAndTitle(int index)
        {
            Spreadsheet spreadsheet;
            int loopCount = -1;
            int returnedId;
            string returnedTitle;
            Tuple<string, int> returnedPair;



            spreadsheet = service.Spreadsheets.Get(spreadSheetsId).Execute();

            try
            {
                foreach (Sheet sheet in spreadsheet.Sheets)
                {
                    loopCount++;
                    if (loopCount != index)
                    {
                        continue;
                    }
                    returnedId = (int)sheet.Properties.SheetId;
                    returnedTitle = sheet.Properties.Title;
                    returnedPair = new Tuple<string, int>(returnedTitle, returnedId);
                    return returnedPair;
                }
            }
            catch (Exception)
            {

                Console.WriteLine("Error in GetListIdAndTitle");
            }
            return null;

        }
        private int GetLastListId()
        {
            var spreadsheet = service.Spreadsheets.Get(spreadSheetsId).Execute();
            return (int)spreadsheet.Sheets[spreadsheet.Sheets.Count - 1].Properties.SheetId;
        }
    }
}
