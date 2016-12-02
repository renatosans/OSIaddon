using System;
using System.Data.SqlClient;
using System.Collections.Generic;
using DataTransferObjects;


namespace DataAccessObjects
{
    public class InvoicePaymentDAO: DataAccessBase
    {
        public InvoicePaymentDAO(SqlConnection sqlConnection)
        {
            this.sqlServerConnection = sqlConnection;
        }

        public InvoicePaymentDTO GetPayment(int invoiceNum)
        {
            InvoicePaymentDTO payment = null;

            String query = "SELECT T0.DocNum, T0.CardCode, T0.CardName, T0.DocTotal FROM ORCT T0" + Environment.NewLine +
                           "LEFT JOIN RCT2 T1 ON T1.DocNum = T0.DocNum" + Environment.NewLine +
                           "LEFT JOIN OINV T2 On T2.DocNum = T1.DocEntry" + Environment.NewLine +
                           "WHERE T2.DocNum = " + invoiceNum;
            SqlCommand command = new SqlCommand(query, sqlServerConnection);
            SqlDataReader dataReader = command.ExecuteReader();
            if (dataReader.Read())
            {
                payment = new InvoicePaymentDTO();
                payment.docNum = GetIntegerValue(dataReader, "DocNum");
                payment.cardCode = GetStringValue(dataReader, "CardCode");
                payment.cardName = GetStringValue(dataReader, "CardName");
                payment.docTotal = GetFloatValue(dataReader, "DocTotal");
            }
            dataReader.Close();

            return payment;
        }
    }

}
