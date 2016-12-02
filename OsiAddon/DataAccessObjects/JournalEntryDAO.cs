using System;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Collections.Specialized;
using DataTransferObjects;


namespace DataAccessObjects
{
    public class JournalEntryDAO: DataAccessBase
    {
        public JournalEntryDAO(SqlConnection sqlConnection)
        {
            this.sqlServerConnection = sqlConnection;
        }

        public List<JournalEntryDTO> GetBoeCredits(int boeNumber)
        {
            List<JournalEntryDTO> creditList = new List<JournalEntryDTO>();

            String query = "SELECT RefDate, Memo, SysTotal FROM OJDT WHERE Ref1 LIKE '%I-" + boeNumber + "' AND TransType = 30";
            SqlCommand command = new SqlCommand(query, sqlServerConnection);
            SqlDataReader dataReader = command.ExecuteReader();
            while (dataReader.Read())
            {
                JournalEntryDTO credit = new JournalEntryDTO();
                credit.refDate = GetDateTimeValue(dataReader, "RefDate");
                credit.memo = GetStringValue(dataReader, "Memo");
                credit.SysTotal = GetFloatValue(dataReader, "SysTotal");

                creditList.Add(credit);
            }
            dataReader.Close();

            return creditList;
        }

        public List<JournalEntryDTO> GetEntriesByPeriod(DateTime startDate, DateTime endDate)
        {
            List<JournalEntryDTO> entryList = new List<JournalEntryDTO>();

            String query = "SELECT TransId, RefDate, Memo, SysTotal, Number FROM OJDT WHERE RefDate >= @startDate AND RefDate <= @endDate ORDER BY RefDate";
            SqlParameter param1 = new SqlParameter("@startDate", System.Data.SqlDbType.DateTime);
            param1.Value = startDate;
            SqlParameter param2 = new SqlParameter("@endDate", System.Data.SqlDbType.DateTime);
            param2.Value = endDate;

            SqlCommand command = new SqlCommand(query, sqlServerConnection);
            command.Parameters.Add(param1);
            command.Parameters.Add(param2);
            SqlDataReader dataReader = command.ExecuteReader();
            while (dataReader.Read())
            {
                JournalEntryDTO entry = new JournalEntryDTO();
                entry.transId = GetIntegerValue(dataReader, "TransId");
                entry.refDate = GetDateTimeValue(dataReader, "RefDate");
                entry.memo = GetStringValue(dataReader, "Memo");
                entry.SysTotal = GetFloatValue(dataReader, "SysTotal");
                entry.number = GetIntegerValue(dataReader, "Number");

                entryList.Add(entry);
            }
            dataReader.Close();

            return entryList;
        }

        public List<JournalEntryItemDTO> GetItems(int transId)
        {
            List<JournalEntryItemDTO> itemList = new List<JournalEntryItemDTO>();

            String query = "SELECT Account, Debit, Credit, LineMemo FROM JDT1 WHERE TransId = " + transId + ";";
            SqlCommand command = new SqlCommand(query, sqlServerConnection);
            SqlDataReader dataReader = command.ExecuteReader();
            while (dataReader.Read())
            {
                JournalEntryItemDTO item = new JournalEntryItemDTO();
                item.account = GetStringValue(dataReader, "Account");
                item.debit = GetFloatValue(dataReader, "Debit");
                item.credit = GetFloatValue(dataReader, "Credit");
                item.lineMemo = GetStringValue(dataReader, "LineMemo");

                itemList.Add(item);
            }
            dataReader.Close();

            return itemList;
        }
    }

}
