using System;


namespace ClassLibrary
{
    public class BillingSummary
    {
        public int counterId;
        public String counterName;
        public decimal consumo;
        public decimal franquia;
        public decimal excedente;
        public float fixo;
        public float variavel;
        public float total;

        public BillingSummary(int counterId, String counterName)
        {
            this.counterId = counterId;
            this.counterName = counterName;
        }
    }

}
