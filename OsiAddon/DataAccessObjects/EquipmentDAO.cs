using System;
using System.Data.SqlClient;
using System.Collections.Generic;
using DataTransferObjects;


namespace DataAccessObjects
{
    public class EquipmentDAO: DataAccessBase
    {
        public EquipmentDAO(SqlConnection sqlServerConnection)
        {
            this.sqlServerConnection = sqlServerConnection;
        }

        public List<EquipmentDTO> GetCustomerEquipments(String customerCardCode)
        {
            List<EquipmentDTO> equipmentList = new List<EquipmentDTO>();

            String query = "SELECT InsID, ManufSN, InternalSN, ItemCode, ItemName, AddrType, Street, StreetNo, Building, Zip, Block, City, State, County, Country, InstLction, Status FROM OINS WHERE Customer = '" + customerCardCode + "' ORDER BY Status ASC, InsID DESC";
            SqlCommand command = new SqlCommand(query, sqlServerConnection);
            SqlDataReader dataReader = command.ExecuteReader();
            while (dataReader.Read())
            {
                EquipmentDTO equipment = new EquipmentDTO();
                equipment.InsID = GetIntegerValue(dataReader, "InsID");
                equipment.ManufSN = GetStringValue(dataReader, "ManufSN");
                equipment.InternalSN = GetStringValue(dataReader, "InternalSN");
                equipment.ItemCode = GetStringValue(dataReader, "ItemCode");
                equipment.ItemName = GetStringValue(dataReader, "ItemName");
                equipment.AddrType = GetStringValue(dataReader, "AddrType");
                equipment.Street = GetStringValue(dataReader, "Street");
                equipment.StreetNo = GetStringValue(dataReader, "StreetNo");
                equipment.Block = GetStringValue(dataReader, "Block");
                equipment.Building = GetStringValue(dataReader, "Building");
                equipment.Zip = GetStringValue(dataReader, "Zip");
                equipment.City = GetStringValue(dataReader, "City");
                equipment.State = GetStringValue(dataReader, "State");
                equipment.County = GetStringValue(dataReader, "County");
                equipment.Country = GetStringValue(dataReader, "Country");
                equipment.InstLocation = GetStringValue(dataReader, "InstLction");
                equipment.Status = GetStringValue(dataReader, "Status");

                equipmentList.Add(equipment);
            }
            dataReader.Close();

            return equipmentList;
        }

        // Recebe como parâmetro os ids separados por vírgula
        public List<EquipmentDTO> GetEquipments(String equipmentIds)
        {
            List<EquipmentDTO> equipmentList = new List<EquipmentDTO>();

            if (String.IsNullOrEmpty(equipmentIds)) equipmentIds = "0";
            String query = "SELECT InsID, ManufSN, InternalSN, ItemCode, ItemName, Customer, InstLction, Status FROM OINS WHERE InsID IN (" + equipmentIds + ") AND status = 'A' ORDER BY ManufSN";
            SqlCommand command = new SqlCommand(query, sqlServerConnection);
            SqlDataReader dataReader = command.ExecuteReader();
            while (dataReader.Read())
            {
                EquipmentDTO equipment = new EquipmentDTO();
                equipment.InsID = GetIntegerValue(dataReader, "InsID");
                equipment.ManufSN = GetStringValue(dataReader, "ManufSN");
                equipment.InternalSN = GetStringValue(dataReader, "InternalSN");
                equipment.ItemCode = GetStringValue(dataReader, "ItemCode");
                equipment.ItemName = GetStringValue(dataReader, "ItemName");
                equipment.Customer = GetStringValue(dataReader, "Customer");
                equipment.InstLocation = GetStringValue(dataReader, "InstLction");
                equipment.Status = GetStringValue(dataReader, "Status");

                equipmentList.Add(equipment);
            }
            dataReader.Close();

            return equipmentList;
        }

        public EquipmentDTO GetEquipment(int insID)
        {
            EquipmentDTO equipment = null;

            String query = "SELECT InsID, ManufSN, InternalSN, ItemCode, ItemName, Customer, InstLction, Status FROM OINS WHERE InsID=" + insID;
            SqlCommand command = new SqlCommand(query, sqlServerConnection);
            SqlDataReader dataReader = command.ExecuteReader();
            if (dataReader.Read())
            {
                equipment = new EquipmentDTO();
                equipment.InsID = GetIntegerValue(dataReader, "InsID");
                equipment.ManufSN = GetStringValue(dataReader, "ManufSN");
                equipment.InternalSN = GetStringValue(dataReader, "InternalSN");
                equipment.ItemCode = GetStringValue(dataReader, "ItemCode");
                equipment.ItemName = GetStringValue(dataReader, "ItemName");
                equipment.Customer = GetStringValue(dataReader, "Customer");
                equipment.InstLocation = GetStringValue(dataReader, "InstLction");
                equipment.Status = GetStringValue(dataReader, "Status");
            }
            dataReader.Close();

            return equipment;
        }

        // Devolve a descrição do status a partir da sigla
        public static String GetStatusDescription(String status)
        {
            switch (status)
            {
                case "R": return "Devolvido";
                case "T": return "Encerrado";
                case "L": return "Emprestado";
                case "I": return "Em reparo";
                default: return "Ativo";
            }
        }

        public void SetSLA(int equipmentCode, int sla)
        {
            String commandText = "UPDATE OINS SET U_SLA=" + sla + " WHERE InsID = " + equipmentCode;
            SqlCommand command = new SqlCommand(commandText, this.sqlServerConnection);
            command.ExecuteNonQuery();
        }
    }

}
