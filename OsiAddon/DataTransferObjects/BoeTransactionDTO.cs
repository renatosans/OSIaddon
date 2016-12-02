using System;


namespace DataTransferObjects
{
    public class BoeTransactionDTO
    {
        public int boeNumber;
        public Decimal boeSum;
        public int paymentNumber;
        public DateTime paymentDate;


        public BoeTransactionDTO()
        {
        }
    }

}
