using System;
using System.Data.SqlClient;
using System.Collections.Generic;
using DataTransferObjects;


namespace DataAccessObjects
{
    public class AccountDAO: DataAccessBase
    {
        public AccountDAO(SqlConnection sqlConnection)
        {
            this.sqlServerConnection = sqlConnection;
        }

        // Recupera somente as contas que não estiverem na raiz da arvore de contas
        public List<AccountDTO> GetLeafAccounts()
        {
            List<AccountDTO> accountList = new List<AccountDTO>();

            String query = "SELECT AcctCode, AcctName, Levels FROM OACT WHERE Levels <> 1 ORDER BY AcctCode";
            SqlCommand command = new SqlCommand(query, sqlServerConnection);
            SqlDataReader dataReader = command.ExecuteReader();
            while (dataReader.Read())
            {
                AccountDTO account = new AccountDTO();
                account.acctCode = GetStringValue(dataReader, "AcctCode");
                account.acctName = GetStringValue(dataReader, "AcctName");
                account.level = (short)dataReader["Levels"];

                accountList.Add(account);
            }
            dataReader.Close();

            return accountList;
        }
    }

}
