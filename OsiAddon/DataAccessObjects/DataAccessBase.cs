using System;
using System.Data.Common;
using System.Data.SqlClient;
using MySql.Data.MySqlClient;


namespace DataAccessObjects
{
    public abstract class DataAccessBase
    {
        protected SqlConnection sqlServerConnection;

        protected MySqlConnection mySqlConnection;


        protected String GetStringValue(DbDataReader dataReader, String fieldName)
        {
            if (dataReader[fieldName] is DBNull) return null;
            return (String)dataReader[fieldName];
        }

        protected int GetIntegerValue(DbDataReader dataReader, String fieldName)
        {
            if (dataReader[fieldName] is DBNull) return 0;
            return (int)dataReader[fieldName];
        }

        protected DateTime GetDateTimeValue(DbDataReader dataReader, String fieldName)
        {
            if (dataReader[fieldName] is DBNull) return new DateTime();
            return (DateTime)dataReader[fieldName];
        }

        protected Decimal GetFloatValue(DbDataReader dataReader, String fieldName)
        {
            if (dataReader[fieldName] is DBNull) return 0;
            return (Decimal)dataReader[fieldName];
        }
    }

}
