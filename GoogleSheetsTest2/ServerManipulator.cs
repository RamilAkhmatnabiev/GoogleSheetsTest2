using Npgsql;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Text;
using System.Threading.Tasks;

namespace GoogleSheetsTest2
{
    class ServerManipulator
    {

        //Переделать :все параметры conn и command получать их в конструктор

        private NpgsqlConnection conn;
        private NpgsqlCommand command = new NpgsqlCommand();
        public ServerManipulator(string connSettings = "Server=127.0.0.1;User Id=postgres;Password=root;")
        {
            conn = new NpgsqlConnection(connSettings);
        }

        public void DoConnection()
        {
            conn.Open();
            command = new NpgsqlCommand("SELECT t1.datname AS db_name,pg_size_pretty(pg_database_size(t1.datname)) AS db_size FROM pg_database t1; ", conn);
        }

        public List<string[]> GetDatabaseInfoForGoogleSheets() // возвращает лист из массива строк с информацией о каждой бд
        {                                                      // возможно есть более выгодный контейнер. Тут именно массив т.к.
            List<string[]> info = new List<string[]>();        // дальнейший код сработает при любом количестве элементов в массиве
            NpgsqlDataReader dr = command.ExecuteReader();

            string[] currentDatabaseInfo;

            while (dr.Read())
            {
                currentDatabaseInfo = new string[]{ dr[0].ToString(), dr[1].ToString()};

                info.Add(currentDatabaseInfo);

                
            }
            return info;
        }
        public void DoDesconnection()
        {
            conn.Close();
        }



       
    //conn.Close();
    }
}
