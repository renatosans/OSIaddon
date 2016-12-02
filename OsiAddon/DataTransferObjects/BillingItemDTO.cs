using System;


namespace DataTransferObjects
{
    public class BillingItemDTO
    {
        public int id;
        public int codigoFaturamento;
        public int contrato_id;
        public int subContrato_id;
        public int codigoCartaoEquipamento;
        public String tipoLocacao;
        public int counterId;
        public String counterName;
        public DateTime dataLeitura;
        public decimal medicaoFinal;
        public decimal medicaoInicial;
        public decimal consumo;
        public decimal ajuste;
        public decimal franquia;
        public decimal excedente;
        public float tarifaSobreExcedente;
        public float fixo;
        public float variavel;
        public float total;
        public float acrescimoDesconto;


        public BillingItemDTO()
        {
        }
    }

}
