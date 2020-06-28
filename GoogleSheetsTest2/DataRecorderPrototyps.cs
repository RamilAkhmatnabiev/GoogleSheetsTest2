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
    partial class DataRecorder
    {
        private string ClientSecret;
        //private string ClientSecret = "client_secret.json";
        private static readonly string[] ScopeSheets = { SheetsService.Scope.Spreadsheets }; // права на использование
        //private static readonly string AppName = "ProgramForPostgressTest"; // имя приложения
        private readonly string AppName; // имя приложения
        //private static readonly string SpreadSheetsId = "18bjPMlVNxm7yQ0Rg1weso9_db6Rg6NrfHgpFj2S-u7s"; // айди таблицы
        private readonly string SpreadSheetsId; // айди таблицы
        private string Range; // диапазон получаемых ячеек строки

        public DataRecorder(string appName = "ProgramForPostgressTest", string spreadSheetsId = "18bjPMlVNxm7yQ0Rg1weso9_db6Rg6NrfHgpFj2S-u7s", string range = "'Sheet1' A1:F")
        {
            this.AppName = appName;
            this.SpreadSheetsId = spreadSheetsId;
            this.Range = range;
            ClientSecret = "client_secret.json";
            var credential = GetSheetCredential(ClientSecret);


        }
    }
}
