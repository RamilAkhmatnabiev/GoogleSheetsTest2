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


        private string ClientSecret = "client_secret.json"; // файл из гугла с апи
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

            DeleteAllSheets(); //очищаем таблицу
            RenameFirstSheet(); // 

            for (int i = 0; i < serversCount; i++)
            {

                CreateNewSheet();
            }
            DeleteFirstSheet();

        }

        public void FillSpreadSheets(List<string[]> info, int sheetId, int index, double discSize)  //заполняет таблицу данными
        {
            List<Request> requests = new List<Request>(); // создаем массив запросов
            string Date;


            for (int i = 0; i < info.Count; i++)
            {
                List<CellData> values = new List<CellData>(); //создаем массив значний

                for (int j = 0; j < info[i].Length; j++) // добавляем каждое значение массива из листа в value, которое в далее передается в request
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
                        //StringValue = DateTime.Now.ToString()
                        StringValue = (DateTime.Today.ToString()).Replace("0:00:00", "")
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

            CreateBottomStringInSheet(sheetId, index, discSize);

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

            CreateTitleStringInSheet(id);


        }

        private void DeleteSheet(int sheetId) // долго работает в массиве
        {


            var deleteSheetRequest = new Request { DeleteSheet = new DeleteSheetRequest { SheetId = sheetId } };

            List<Request> requests = new List<Request> { deleteSheetRequest };
            BatchUpdateSpreadsheetRequest batchUpdateRequest = new BatchUpdateSpreadsheetRequest();
            batchUpdateRequest.Requests = requests;
            service.Spreadsheets.BatchUpdate(batchUpdateRequest, spreadSheetsId).Execute();
        }

        private void DeleteAllSheets()
        {
            Spreadsheet spreadsheet;

            List<int> ids = new List<int>();
            List<Request> deleteRequests = new List<Request>();
            int count = 0;



            spreadsheet = service.Spreadsheets.Get(spreadSheetsId).Execute();



            foreach (Sheet sheet in spreadsheet.Sheets)
            {
                if (count == 0)
                {
                    count++;
                    continue;
                }
                deleteRequests.Add(new Request { DeleteSheet = new DeleteSheetRequest { SheetId = sheet.Properties.SheetId } });
                count++;

            }
            if (count > 1)
            {


                BatchUpdateSpreadsheetRequest batchUpdateRequest = new BatchUpdateSpreadsheetRequest();
                batchUpdateRequest.Requests = deleteRequests;
                service.Spreadsheets.BatchUpdate(batchUpdateRequest, spreadSheetsId).Execute();
            }
        }

        private void DeleteFirstSheet()
        {
            Spreadsheet spreadsheet;
            List<Request> deleteRequests = new List<Request>();
            int deletedSheetId;
            spreadsheet = service.Spreadsheets.Get(spreadSheetsId).Execute();
            foreach (Sheet sheet in spreadsheet.Sheets)
            {

                deletedSheetId = (int)sheet.Properties.SheetId;
                DeleteSheet(deletedSheetId);
                break;
            }

        }
        private void RenameFirstSheet()
        {
            int id = new int();
            id = 333;
            //var request = new Request { UpdateSpreadsheetProperties = new UpdateSpreadsheetPropertiesRequest {Properties = new SpreadsheetProperties { Title = "Del" }} }; //{ Properties = new SheetProperties { Title = "Delet" } } };
            var request = new Request()
            {
                UpdateSheetProperties = new UpdateSheetPropertiesRequest
                {
                    Properties = new SheetProperties()
                    {
                        Title = "Deleted",
                        SheetId = GetThisIndexSheetId(0)


                    },
                    Fields = "Title"
                }


            };
            List<Request> requests = new List<Request>();
            requests.Add(request);

            BatchUpdateSpreadsheetRequest busr = new BatchUpdateSpreadsheetRequest
            {
                Requests = requests
            };
            service.Spreadsheets.BatchUpdate(busr, spreadSheetsId).Execute();
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
        private int GetLastListId()// создает первую заголовочную строку 
        {
            var spreadsheet = service.Spreadsheets.Get(spreadSheetsId).Execute();
            return (int)spreadsheet.Sheets[spreadsheet.Sheets.Count - 1].Properties.SheetId;
        }

        public int GetThisIndexSheetId(int index)
        {

            int loop = -1;
            int id = -1;
            var spreadsheet = service.Spreadsheets.Get(spreadSheetsId).Execute();
            foreach (Sheet sheet in spreadsheet.Sheets)
            {
                loop++;
                if (loop != index)
                {
                    continue;
                }
                id = (int)sheet.Properties.SheetId;

            }

            return id;
        }

        private void CreateTitleStringInSheet(int sheetId)
        {
            string[] data = { "Сервер", "База данных", "Размер", "Время обновления" };
            List<Request> requests = new List<Request>(); // создаем массив запросов

            List<CellData> values = new List<CellData>(); //создаем массив значний

            for (int i = 0; i < data.Length; i++) // добавляем каждое значение массива из листа в value, которое в далее передается в request
            {
                values.Add(new CellData
                {
                    UserEnteredValue = new ExtendedValue
                    {
                        StringValue = data[i]
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
                            SheetId = sheetId,
                            RowIndex = 0,
                            ColumnIndex = 0
                        },
                        Rows = new List<RowData> { new RowData { Values = values } },
                        Fields = "userEnteredValue"
                    }
                });



            BatchUpdateSpreadsheetRequest busr = new BatchUpdateSpreadsheetRequest
            {
                Requests = requests
            };
            service.Spreadsheets.BatchUpdate(busr, spreadSheetsId).Execute();

        }

        private void CreateBottomStringInSheet(int sheetId, int index, double discSize)
        {


            string discSizeField = "=";
            discSizeField += discSize.ToString();
            for (int i = 0; i < allSheets.Count; i++)
            {
                discSizeField += "-C" + (i + 2).ToString();
            }

            string timeOfRefresh = DateTime.Today.ToString();
            timeOfRefresh = timeOfRefresh.Replace("0:00:00", " ");

            string[] data = { allSheets[index].Item2, "Свободно", discSizeField, timeOfRefresh };


            List<Request> requests = new List<Request>(); // создаем массив запросов

            List<CellData> values = new List<CellData>(); //создаем массив значний

            for (int i = 0; i < data.Length; i++) // добавляем каждое значение массива из листа в value, которое в далее передается в request
            {
                values.Add(new CellData
                {
                    UserEnteredValue = new ExtendedValue
                    {
                        StringValue = data[i]
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
                            SheetId = sheetId,
                            RowIndex = allSheets.Count + 2,
                            ColumnIndex = 0
                        },
                        Rows = new List<RowData> { new RowData { Values = values } },
                        Fields = "userEnteredValue"
                    }
                });



            BatchUpdateSpreadsheetRequest busr = new BatchUpdateSpreadsheetRequest
            {
                Requests = requests
            };
            service.Spreadsheets.BatchUpdate(busr, spreadSheetsId).Execute();

        }
    }
}
