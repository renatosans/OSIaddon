using System;
using System.Data.SqlClient;
using System.Collections.Generic;
using DataTransferObjects;


namespace DataAccessObjects
{
    public class BoeTransactionDAO: DataAccessBase
    {
        public BoeTransactionDAO(SqlConnection sqlConnection)
        {
            this.sqlServerConnection = sqlConnection;
        }

        public List<BoeTransactionDTO> GetBoePayments()
        {
            List<BoeTransactionDTO> paymentList = new List<BoeTransactionDTO>();

            String query = "DECLARE @BOT1 AS TABLE(AbsEntry INT, BOENumber INT)" + Environment.NewLine +
                           "INSERT INTO @BOT1 SELECT MIN(AbsEntry), BOENumber FROM BOT1 GROUP BY BOENumber" + Environment.NewLine +
                           "DECLARE @BOT AS TABLE(AbsEntry INT, PostDate DATETIME)" + Environment.NewLine +
                           "INSERT INTO @BOT SELECT AbsEntry, PostDate FROM OBOT" + Environment.NewLine +
                           "SELECT BOL.[BoeNum] AS 'BoeNumber', BOL.[BoeSum] AS 'BoeSum', BOL.[PmntNum] AS PaymentNumber, TOBOT.[PostDate] AS 'PaymentDate' FROM OBOE BOL" + Environment.NewLine +
                           "LEFT JOIN @BOT1 TBOT1 ON BOL.BoeNum = TBOT1.BOENumber" + Environment.NewLine +
                           "LEFT JOIN @BOT TOBOT ON TBOT1.AbsEntry = TOBOT.AbsEntry" + Environment.NewLine +
                           "WHERE BOL.[BoeStatus] = 'P' AND BOL.[U_PaymentDate] IS NULL";
            SqlCommand command = new SqlCommand(query, sqlServerConnection);
            SqlDataReader dataReader = command.ExecuteReader();
            while(dataReader.Read())
            {
                BoeTransactionDTO boeTransaction = new BoeTransactionDTO();
                boeTransaction.boeNumber = GetIntegerValue(dataReader, "BoeNumber");
                boeTransaction.boeSum = GetFloatValue(dataReader, "BoeSum");
                boeTransaction.paymentNumber = GetIntegerValue(dataReader, "PaymentNumber");
                boeTransaction.paymentDate = GetDateTimeValue(dataReader, "PaymentDate");

                paymentList.Add(boeTransaction);
            }
            dataReader.Close();

            return paymentList;
        }
    }

}
