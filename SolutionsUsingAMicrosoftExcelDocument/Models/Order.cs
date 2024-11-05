using System;

namespace SolutionsUsingAMicrosoftExcelDocument.Models
{
    /// <summary>
    /// Модель заявки
    /// </summary>
    public class Order
    {
        //Код заявки
        public int Code { get; set; }
        //Код товара
        public int ProductCode { get; set; }
        //Код клиента
        public int ClientCode { get; set; }
        //Номер заявки
        public int OrderNumber { get; set; }
        //Требуемое количество
        public int Quantity { get; set; }
        //Дата размещения
        public DateTime Date { get; set; }
    }
}
