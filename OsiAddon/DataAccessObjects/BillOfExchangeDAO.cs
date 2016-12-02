using System;
using System.Data.SqlClient;
using System.Collections.Generic;
using DataTransferObjects;


namespace DataAccessObjects
{
    public class BillOfExchangeDAO: DataAccessBase
    {
        public BillOfExchangeDAO(SqlConnection sqlConnection)
        {
            this.sqlServerConnection = sqlConnection;
        }

        public BillOfExchangeDTO GetBillOfExchange(int boeNum)
        {
            String query = "SELECT BoeNum, DueDate, BoeSum, OurNum, OurNumChk, RefNum, CardCode, CardName FROM OBOE WHERE BoeNum = " + boeNum;
            SqlCommand command = new SqlCommand(query, sqlServerConnection);
            SqlDataReader dataReader = command.ExecuteReader();
            if (!dataReader.Read()) return null;
            BillOfExchangeDTO boe = new BillOfExchangeDTO();
            boe.BoeNum = GetIntegerValue(dataReader, "BoeNum");
            boe.DueDate = GetDateTimeValue(dataReader, "DueDate");
            boe.BoeSum = GetFloatValue(dataReader, "BoeSum");
            boe.OurNum = GetIntegerValue(dataReader, "OurNum");
            boe.OurNumChk = GetStringValue(dataReader, "OurNumChk");
            boe.RefNum = GetStringValue(dataReader, "RefNum");
            boe.CardCode = GetStringValue(dataReader, "CardCode");
            boe.CardName = GetStringValue(dataReader, "CardName");
            dataReader.Close();

            return boe;
        }

        public List<BillOfExchangeDTO> GetBillsOfExchange(int[] boeNumbers)
        {
            List<BillOfExchangeDTO> boeList = new List<BillOfExchangeDTO>();

            String boeFilter = null;
            foreach (int boeNum in boeNumbers)
            {
                if (boeFilter != null) boeFilter += ", ";
                boeFilter += boeNum;
            }
            String query = "SELECT BoeNum, DueDate, BoeSum, OurNum, OurNumChk, RefNum, CardCode, CardName FROM OBOE " + String.Format("WHERE BoeNum IN ({0})", boeFilter);
            SqlCommand command = new SqlCommand(query, sqlServerConnection);
            SqlDataReader dataReader = command.ExecuteReader();
            while (dataReader.Read())
            {
                BillOfExchangeDTO boe = new BillOfExchangeDTO();
                boe.BoeNum = GetIntegerValue(dataReader, "BoeNum");
                boe.DueDate = GetDateTimeValue(dataReader, "DueDate");
                boe.BoeSum = GetFloatValue(dataReader, "BoeSum");
                boe.OurNum = GetIntegerValue(dataReader, "OurNum");
                boe.OurNumChk = GetStringValue(dataReader, "OurNumChk");
                boe.RefNum = GetStringValue(dataReader, "RefNum");
                boe.CardCode = GetStringValue(dataReader, "CardCode");
                boe.CardName = GetStringValue(dataReader, "CardName");
                boeList.Add(boe);
            }
            dataReader.Close();

            return boeList;
        }
    }

}
