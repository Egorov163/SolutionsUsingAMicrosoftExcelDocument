using System;

namespace SolutionsUsingAMicrosoftExcelDocument.Models
{
    public class Order
    {
        public int Code { get; set; }
        public int ProductCode { get; set; }
        public int ClientCode { get; set; }
        public int OrderNumber { get; set; }
        public int Quantity { get; set; }
        public DateTime Date { get; set; }
    }
}
