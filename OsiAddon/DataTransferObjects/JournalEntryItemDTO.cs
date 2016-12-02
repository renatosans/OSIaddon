using System;


namespace DataTransferObjects
{
    public class JournalEntryItemDTO
    {
        public String account;
        public Decimal debit;
        public Decimal credit;
        public String lineMemo;


        public JournalEntryItemDTO()
        {
        }
    }

}
