using System;


namespace DataTransferObjects
{
    public class BillOfExchangeDTO
    {
        public int BoeNum;
        public DateTime DueDate;
        public Decimal BoeSum;
        public int OurNum;
        public String OurNumChk;
        public String RefNum;
        public String CardCode;
        public String CardName;


        public BillOfExchangeDTO()
        {
        }
    }

}
