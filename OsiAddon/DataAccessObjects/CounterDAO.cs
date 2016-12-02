using System;
using System.Collections.Generic;
using MySql.Data.MySqlClient;


namespace DataAccessObjects
{
    public class CounterDAO: DataAccessBase
    {
        public CounterDAO(MySqlConnection mySqlConnection)
        {
            this.mySqlConnection = mySqlConnection;
        }

        public Dictionary<int, String> GetAllCounter()
        {
            Dictionary<int, String> counters = new Dictionary<int, String>();

            String query = "SELECT * FROM `addoncontratos`.`contador`";
            MySqlCommand command = new MySqlCommand(query, this.mySqlConnection);
            MySqlDataReader dataReader = command.ExecuteReader();
            while (dataReader.Read())
            {
                counters.Add((int)dataReader["id"], (String)dataReader["nome"]);
            }
            dataReader.Close();

            return counters;
        }
    }

}
