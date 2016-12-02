using System;


namespace DataTransferObjects
{
    public class InvoicePaymentDTO
    {
        public int docNum;
        public String cardCode;
        public String cardName;
        public Decimal docTotal;


        public InvoicePaymentDTO()
        {
        }
    }

}
