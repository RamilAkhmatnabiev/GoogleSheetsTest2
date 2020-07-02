﻿using Google.Apis.Auth.OAuth2;
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
        private string ClientSecret = "client_secret.json";
        //private string ClientSecret = "client_secret.json";
        private readonly string[] ScopeSheets = { SheetsService.Scope.Spreadsheets }; // права на использование
        private readonly string AppName = "ProgramForPostgressTest"; // имя приложения

        //private string AppName; // имя приложения
        //private static readonly string SpreadSheetsId = "18bjPMlVNxm7yQ0Rg1weso9_db6Rg6NrfHgpFj2S-u7s"; // айди таблицы
        private string spreadSheetsId; // айди таблицы
        private string Range = "'Sheet1' B2:F"; // диапазон получаемых ячеек строки

        private int serversNumber;
        private int currentServerNumber; //нужен для создания нового листа, если серверов несколько
        private static bool isFirstCreate = true;// нужен для проверки условия при удалении первого листа, если подключаемся к таблице впервые

        private UserCredential credential;// нужен для хранения credential
        private SheetsService service;  // нужен для хранения service




        //18bjPMlVNxm7yQ0Rg1weso9_db6Rg6NrfHgpFj2S-u7s
        //10dFGbsi3Av-sqViDzCkxRaqE9TPFTozbwaZN_sQdhEo

        public DataRecorder(int currentServerNumber,int serversCount, string spreadSheetsId = "18bjPMlVNxm7yQ0Rg1weso9_db6Rg6NrfHgpFj2S-u7s")
        {
            this.spreadSheetsId = spreadSheetsId;

            this.currentServerNumber = currentServerNumber;

            credential = GetSheetCredential(ClientSecret); // формируем credential на базе файла client secret

            service = GetService(credential); // подключаемся к google  с credential  и получаем service


            if (isFirstCreate) // в первом подключении удаляем листы по необходимым ID, и создаем их заново (в дальнейшем нужно оптимизировать код убрав эту часть)
            {
                for (int i = 0; i < serversCount; i++)
                {
                   
                    DeleteSheet(i);
                    CreateNewSheet(i);
                }
                                
                isFirstCreate = false;
            }

            

        }

        public void FillSpreadSheets(List<string[]> info)  //заполняет таблицу данными
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
                                SheetId = currentServerNumber,
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

        private void CreateNewSheet(int currentServerNumber)
        {
            string name = "server" + currentServerNumber.ToString();

            var addSheetRequest = new Request
            {
                AddSheet = new AddSheetRequest
                {
                    Properties = new SheetProperties
                    {
                        SheetId = currentServerNumber,
                        Title = name
                    }
                }
            };
            
            List<Request> requests = new List<Request> { addSheetRequest };
            BatchUpdateSpreadsheetRequest batchUpdateRequest = new BatchUpdateSpreadsheetRequest();
            batchUpdateRequest.Requests = requests;
            service.Spreadsheets.BatchUpdate(batchUpdateRequest, spreadSheetsId).Execute();
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
        private int GetListId(int index)
        {
            var spreadsheet = service.Spreadsheets.Get(spreadSheetsId).Execute();
            return (int)spreadsheet.Sheets[1].Properties.SheetId;
        }
    }
}
