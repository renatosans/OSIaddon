using System;
using System.Collections.Generic;
using DataTransferObjects;
using MySql.Data.MySqlClient;


namespace DataAccessObjects
{
    public class BillingDAO: DataAccessBase
    {
        public BillingDAO(MySqlConnection mySqlConnection)
        {
            this.mySqlConnection = mySqlConnection;
        }

        public BillingDTO GetBilling(int id)
        {
            BillingDTO billing = null;

            String query = "SELECT * FROM `addoncontratos`.`faturamento` WHERE id=" + id;
            MySqlCommand command = new MySqlCommand(query, this.mySqlConnection);
            MySqlDataReader dataReader = command.ExecuteReader();
            if (dataReader.Read())
            {
                billing = new BillingDTO();
                billing.id = (int)dataReader["id"];
                billing.businessPartnerCode = (String)dataReader["businessPartnerCode"];
                billing.businessPartnerName = (String)dataReader["businessPartnerName"];
                billing.mailing_id = (int)dataReader["mailing_id"];
                billing.dataInicial = (DateTime)dataReader["dataInicial"];
                billing.dataFinal = (DateTime)dataReader["dataFinal"];
                billing.mesReferencia = (int)dataReader["mesReferencia"];
                billing.anoReferencia = (int)dataReader["anoReferencia"];
                billing.acrescimoDesconto = (float)dataReader["acrescimoDesconto"];
                billing.total = (float)dataReader["total"];
                billing.obs = (String)dataReader["obs"];
                billing.incluirRelatorio = (Boolean)dataReader["incluirRelatorio"];
            }
            dataReader.Close();

            return billing;
        }

        public List<BillingDTO> GetAllBillings(String filter)
        {
            List<BillingDTO> billingList = new List<BillingDTO>();

            String query = "SELECT * FROM `addoncontratos`.`faturamento`";
            if (!String.IsNullOrEmpty(filter)) query = "SELECT * FROM `addoncontratos`.`faturamento` WHERE " + filter;
            MySqlCommand command = new MySqlCommand(query, this.mySqlConnection);
            MySqlDataReader dataReader = command.ExecuteReader();
            while (dataReader.Read())
            {
                BillingDTO billing = new BillingDTO();
                billing.id = (int)dataReader["id"];
                billing.businessPartnerCode = (String)dataReader["businessPartnerCode"];
                billing.businessPartnerName = (String)dataReader["businessPartnerName"];
                billing.mailing_id = (int)dataReader["mailing_id"];
                billing.dataInicial = (DateTime)dataReader["dataInicial"];
                billing.dataFinal = (DateTime)dataReader["dataFinal"];
                billing.mesReferencia = (int)dataReader["mesReferencia"];
                billing.anoReferencia = (int)dataReader["anoReferencia"];
                billing.acrescimoDesconto = (float)dataReader["acrescimoDesconto"];
                billing.total = (float)dataReader["total"];
                billing.obs = (String)dataReader["obs"];
                billing.incluirRelatorio = (Boolean)dataReader["incluirRelatorio"];

                billingList.Add(billing);
            }
            dataReader.Close();

            return billingList;
        }

        public void SetBilling(BillingDTO billing)
        {
            IFormatProvider invariantCulture = System.Globalization.CultureInfo.InvariantCulture;

            String commandText = "UPDATE `addoncontratos`.`faturamento` SET mesReferencia = " + billing.mesReferencia + ", anoReferencia = " + billing.anoReferencia + ", acrescimoDesconto = " + String.Format(invariantCulture, "{0:0.00}", billing.acrescimoDesconto) + ", obs='" + billing.obs + "' WHERE id=" + billing.id;
            MySqlCommand command = new MySqlCommand(commandText, this.mySqlConnection);
            command.ExecuteNonQuery();
        }
    }

}
