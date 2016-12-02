using System;
using System.Collections.Generic;
using DataTransferObjects;
using MySql.Data.MySqlClient;


namespace DataAccessObjects
{
    public class BillingItemDAO : DataAccessBase
    {
        public BillingItemDAO(MySqlConnection mySqlConnection)
        {
            this.mySqlConnection = mySqlConnection;
        }

        public List<BillingItemDTO> GetBillingItems(int billingId)
        {
            List<BillingItemDTO> billingItems = new List<BillingItemDTO>();

            String fieldList = "ITM.id, ITM.codigoFaturamento, ITM.contrato_id, ITM.subContrato_id, ITM.codigoCartaoEquipamento, ITM.tipoLocacao, ITM.counterId, CNTR.nome as counterName, ";
            fieldList = fieldList + "IF(ITM.dataLeitura LIKE '%0000-00-00%', null, ITM.dataLeitura) as dataLeitura, ITM.medicaoFinal, ITM.medicaoInicial, ITM.consumo, ITM.ajuste, ";
            fieldList = fieldList + "ITM.franquia, ITM.excedente, ITM.tarifaSobreExcedente, ITM.fixo, ITM.variavel, ITM.total, ITM.acrescimoDesconto";

            String query = "SELECT " + fieldList + " FROM `addoncontratos`.`itemFaturamento` ITM JOIN `addoncontratos`.`contador` CNTR ON CNTR.id = ITM.counterId WHERE codigoFaturamento=" + billingId;
            MySqlCommand command = new MySqlCommand(query, this.mySqlConnection);
            MySqlDataReader dataReader = command.ExecuteReader();
            while (dataReader.Read())
            {
                BillingItemDTO billing = new BillingItemDTO();
                billing.id = (int)dataReader["id"];
                billing.codigoFaturamento = (int)dataReader["codigoFaturamento"];
                billing.contrato_id = (int)dataReader["contrato_id"];
                billing.subContrato_id = (int)dataReader["subContrato_id"];
                billing.codigoCartaoEquipamento = (int)dataReader["codigoCartaoEquipamento"];
                billing.tipoLocacao = (String)dataReader["tipoLocacao"];
                billing.counterId = (int)dataReader["counterId"];
                billing.counterName = (String)dataReader["counterName"];
                billing.dataLeitura = GetDateTimeValue(dataReader, "dataLeitura");
                billing.medicaoFinal = (decimal)dataReader["medicaoFinal"];
                billing.medicaoInicial = (decimal)dataReader["medicaoInicial"];
                billing.consumo = (decimal)dataReader["consumo"];
                billing.ajuste = (decimal)dataReader["ajuste"];
                billing.franquia = (decimal)dataReader["franquia"];
                billing.excedente = (decimal)dataReader["excedente"];
                billing.tarifaSobreExcedente = (float)dataReader["tarifaSobreExcedente"];
                billing.fixo = (float)dataReader["fixo"];
                billing.variavel = (float)dataReader["variavel"];
                billing.total = (float)dataReader["total"];
                billing.acrescimoDesconto = (float)dataReader["acrescimoDesconto"];

                billingItems.Add(billing);
            }
            dataReader.Close();

            return billingItems;
        }
    }

}
