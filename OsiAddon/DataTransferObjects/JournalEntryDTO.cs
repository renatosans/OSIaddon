using System;


namespace DataTransferObjects
{
    public class JournalEntryDTO
    {
        public int transId;
        public DateTime refDate;
        public String memo;
        public Decimal SysTotal;
        public int number;


        public JournalEntryDTO()
        {
        }
    }

}
