using System;
using System.Data.SqlClient;
using System.Collections.Generic;
using DataTransferObjects;


namespace DataAccessObjects
{
    public class InvoiceDAO : DataAccessBase
    {
        public InvoiceDAO(SqlConnection sqlServerConnection)
        {
            this.sqlServerConnection = sqlServerConnection;
        }

        public List<InvoiceDTO> GetAllInvoices(String filter)
        {
            List<InvoiceDTO> invoiceList = new List<InvoiceDTO>();
            if (String.IsNullOrEmpty(filter)) return invoiceList; // retorna a lista vazia

            String query = "SELECT OINV.DocNum, OINV.DocDate, OINV.Comments, OINV.DocTotal, OINV.U_demFaturamento FROM OINV ";
            query = query + "JOIN INV1 ON OINV.DocEntry = INV1.DocEntry WHERE " + filter;
            SqlCommand command = new SqlCommand(query, sqlServerConnection);
            SqlDataReader dataReader = command.ExecuteReader();
            while (dataReader.Read())
            {
                InvoiceDTO invoice = new InvoiceDTO();
                invoice.docNum = (int)dataReader["DocNum"];
                invoice.docDate = (DateTime)dataReader["DocDate"];
                invoice.comments = (String)dataReader["Comments"];
                invoice.docTotal = (decimal)dataReader["DocTotal"];
                invoice.demFaturamento = GetIntegerValue(dataReader, "U_demFaturamento");

                invoiceList.Add(invoice);
            }
            dataReader.Close();

            return invoiceList;
        }

        public List<InvoiceDTO> GetReturnedInvoices(String filter)
        {
            List<InvoiceDTO> invoiceList = new List<InvoiceDTO>();
            if (String.IsNullOrEmpty(filter)) return invoiceList; // retorna a lista vazia

            String query = "SELECT DISTINCT T3.DocNum, T3.DocDate, T3.Comments, T3.DocTotal, T3.U_demFaturamento FROM ";
            query = query + "ORIN T0 INNER JOIN RIN1 T1 ON T0.DocEntry = T1.DocEntry LEFT JOIN ";
            query = query + "INV1 T2 ON T1.BaseEntry = T2.DocEntry AND T1.BaseLine = T2.LineNum AND T1.BaseType = 13 INNER JOIN ";
            query = query + "OINV T3 ON T2.DocEntry = T3.DocEntry WHERE " + filter;
            SqlCommand command = new SqlCommand(query, sqlServerConnection);
            SqlDataReader dataReader = command.ExecuteReader();
            while (dataReader.Read())
            {
                InvoiceDTO invoice = new InvoiceDTO();
                invoice.docNum = (int)dataReader["DocNum"];
                invoice.docDate = (DateTime)dataReader["DocDate"];
                invoice.comments = (String)dataReader["Comments"];
                invoice.docTotal = (decimal)dataReader["DocTotal"];
                invoice.demFaturamento = GetIntegerValue(dataReader, "U_demFaturamento");

                invoiceList.Add(invoice);
            }
            dataReader.Close();

            return invoiceList;
        }
    }

}
