using System;
using System.Data.SqlClient;
using System.Collections.Generic;
using DataTransferObjects;


namespace DataAccessObjects
{
    public class EmployeeDAO: DataAccessBase
    {
        public EmployeeDAO(SqlConnection sqlConnection)
        {
            this.sqlServerConnection = sqlConnection;
        }

        public List<EmployeeDTO> GetAllTechnicians()
        {
            List<EmployeeDTO> technicianList = new List<EmployeeDTO>();

            String subQuery = "SELECT posId FROM OHPS WHERE name LIKE '%Técnico%' OR name LIKE '%Tecnico%'";
            String query = "SELECT empId, firstName, lastName, jobTitle FROM OHEM WHERE position IN (" + subQuery + ")";
            SqlCommand command = new SqlCommand(query, sqlServerConnection);
            SqlDataReader dataReader = command.ExecuteReader();
            while (dataReader.Read())
            {
                EmployeeDTO technician = new EmployeeDTO();
                technician.empID = GetIntegerValue(dataReader, "empID");
                technician.firstName = GetStringValue(dataReader, "firstName");
                technician.lastName = GetStringValue(dataReader, "lastName");
                technician.jobTitle = GetStringValue(dataReader, "jobTitle");

                technicianList.Add(technician);
            }
            dataReader.Close();

            return technicianList;
        }
    }

}
