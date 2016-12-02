using System;
using System.IO;
using System.Xml;
using System.Reflection;
using System.Data;
using System.Data.SqlClient;
using MySql.Data.MySqlClient;


namespace OsiAddon
{
    public class DataConnector
    {
        private MySqlConnection mySqlConnection;

        private SqlConnection sqlServerConnection;

        public MySqlConnection MySqlConnection
        {
            get { return mySqlConnection; }
        }

        public SqlConnection SqlServerConnection
        {
            get { return sqlServerConnection; }
        }

        public DataConnector()
        {
            String baseDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location.ToString());
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(baseDir + @"\Xml\DataAccess.xml");
            XmlNodeList nodeList = xmlDoc.ChildNodes[1].ChildNodes;
            XmlNode mylSqlNode = null;
            XmlNode sqlServerNode = null;
            foreach (XmlNode node in nodeList)
            {
                if (node.Attributes["name"].Value == "MySQL 5.5")
                    mylSqlNode = node;
                if (node.Attributes["name"].Value == "SQL Server 2008")
                    sqlServerNode = node;
            }
            if ((mylSqlNode == null) || (sqlServerNode == null)) return;

            String server = mylSqlNode.SelectSingleNode("server").InnerText;
            String database = mylSqlNode.SelectSingleNode("database").InnerText;
            String username = mylSqlNode.SelectSingleNode("username").InnerText;
            String password = mylSqlNode.SelectSingleNode("password").InnerText;
            String connectionString = "server=" + server + ";user id=" + username + ";password=" + password + ";";
            if (!String.IsNullOrEmpty(database)) connectionString += "database=" + database + ";";
            mySqlConnection = new MySqlConnection(connectionString);

            server = sqlServerNode.SelectSingleNode("server").InnerText;
            database = sqlServerNode.SelectSingleNode("database").InnerText;
            username = sqlServerNode.SelectSingleNode("username").InnerText;
            password = sqlServerNode.SelectSingleNode("password").InnerText;
            connectionString = @"Data Source=" + server + ";Initial Catalog=" + database + ";User=" + username + "; password=" + password;
            sqlServerConnection = new SqlConnection(connectionString);
        }

        public void OpenConnection(String databaseType)
        {
            if ((databaseType == "mySql") || (databaseType == "both"))
                mySqlConnection.Open();

            if ((databaseType == "sqlServer") || (databaseType == "both"))
                sqlServerConnection.Open();
        }

        public void CloseConnection()
        {
            if (mySqlConnection.State == ConnectionState.Open)
                mySqlConnection.Close();

            if (sqlServerConnection.State == ConnectionState.Open)
                sqlServerConnection.Close();
        }
    }

}
