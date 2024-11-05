namespace SolutionsUsingAMicrosoftExcelDocument.Models
{
    /// <summary>
    /// Модель клиента
    /// </summary>
    public class Client
    {
        //Код клиента
        public int Code { get; set; }
        //Наименование организации
        public string Name { get; set; }
        //Адрес
        public string Address { get; set; }
        //Контактное лицо (ФИО)
        public string ContactPerson { get; set; }
    }
}
