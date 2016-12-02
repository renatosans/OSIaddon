using System;
using System.Data.SqlClient;
using DataTransferObjects;


namespace DataAccessObjects
{
    public class PartnerContactDAO: DataAccessBase
    {
        public PartnerContactDAO(SqlConnection sqlConnection)
        {
            this.sqlServerConnection = sqlConnection;
        }

        public PartnerContactDTO GetContact(String cardCode, String contactName)
        {
            String query = "SELECT CntctCode, Name, E_MailL FROM OCPR WHERE CardCode = '" + cardCode + "' AND Name = '" + contactName + "'";
            SqlCommand command = new SqlCommand(query, sqlServerConnection);
            SqlDataReader dataReader = command.ExecuteReader();
            if (!dataReader.Read()) return null;
            PartnerContactDTO partnerContact = new PartnerContactDTO();
            partnerContact.CntctCode = GetIntegerValue(dataReader, "CntctCode");
            partnerContact.Name = GetStringValue(dataReader, "Name");
            partnerContact.Email = GetStringValue(dataReader, "E_MailL");
            dataReader.Close();

            return partnerContact;
        }
    }

}
