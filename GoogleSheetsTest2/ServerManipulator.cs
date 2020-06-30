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

        private NpgsqlConnection conn = new NpgsqlConnection("Server=127.0.0.1;User Id=postgres; " + "Password=root;");
        private NpgsqlCommand command = new NpgsqlCommand();
        public ServerManipulator()
        {

        }

        public void DoConnection()
        {
            conn.Open();
            command = new NpgsqlCommand("SELECT t1.datname AS db_name,pg_size_pretty(pg_database_size(t1.datname)) AS db_size FROM pg_database t1; ", conn);
        }

        public List<Tuple<string, string>> GetDatabaseInfoForGoogleSheets()
        {
            List<Tuple<string, string>> info = new List<Tuple<string, string>>();
            NpgsqlDataReader dr = command.ExecuteReader();

            Tuple<string, string> currentDatabaseInfo;

            while (dr.Read())
            {
                currentDatabaseInfo = new Tuple<string, string>(dr[0].ToString(), dr[1].ToString());

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
