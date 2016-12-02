using System;


namespace DataTransferObjects
{
    public class BillingDTO
    {
        public int id;
        public String businessPartnerCode;
        public String businessPartnerName;
        public int mailing_id;
        public DateTime dataInicial;
        public DateTime dataFinal;
        public int mesReferencia;
        public int anoReferencia;
        public float acrescimoDesconto;
        public float total;
        public String obs;
        public Boolean incluirRelatorio;


        public BillingDTO()
        {
        }
    }

}
