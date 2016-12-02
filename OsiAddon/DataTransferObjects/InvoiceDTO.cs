using System;


namespace DataTransferObjects
{
    public class InvoiceDTO: IEquatable<InvoiceDTO>
    {
        public int docNum;
        public DateTime docDate;
        public String comments;
        public decimal docTotal;
        public int? demFaturamento;

        public InvoiceDTO()
        {
        }

        public bool Equals(InvoiceDTO other)
        {
            if (this.docNum == other.docNum)
                return true;
            else
                return false;
        }
    }

}
