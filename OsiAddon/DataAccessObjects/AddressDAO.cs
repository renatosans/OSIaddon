using System;
using System.Data.SqlClient;
using System.Collections.Generic;
using DataTransferObjects;


namespace DataAccessObjects
{
    public class AddressDAO: DataAccessBase
    {
        public AddressDAO(SqlConnection sqlConnection)
        {
            this.sqlServerConnection = sqlConnection;
        }

        public List<AddressDTO> GetPartnerAddresses(String cardCode)
        {
            List<AddressDTO> addressList = new List<AddressDTO>();

            String query = "SELECT Address, CardCode, AddrType, Street, StreetNo, Block, Building, ZipCode, City, State, County, Country FROM CRD1 WHERE CardCode = '" + cardCode + "'";
            SqlCommand command = new SqlCommand(query, sqlServerConnection);
            SqlDataReader dataReader = command.ExecuteReader();
            while (dataReader.Read())
            {
                AddressDTO address = new AddressDTO();
                address.Address = GetStringValue(dataReader, "Address");
                address.CardCode = GetStringValue(dataReader, "CardCode");
                address.AddrType = GetStringValue(dataReader, "AddrType");
                address.Street = GetStringValue(dataReader, "Street");
                address.StreetNo = GetStringValue(dataReader, "StreetNo");
                address.Block = GetStringValue(dataReader, "Block");
                address.Building = GetStringValue(dataReader, "Building");
                address.ZipCode = GetStringValue(dataReader, "ZipCode");
                address.City = GetStringValue(dataReader, "City");
                address.State = GetStringValue(dataReader, "State");
                address.County = GetStringValue(dataReader, "County");
                address.Country = GetStringValue(dataReader, "Country");

                addressList.Add(address);
            }
            dataReader.Close();

            return addressList;
        }
    }

}
