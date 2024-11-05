namespace SolutionsUsingAMicrosoftExcelDocument.Models
{
    /// <summary>
    /// Товары
    /// </summary>
    public class Product
    {
        //Код товара
        public int Code { get; set; }
        //Наименование
        public string Name { get; set; }
        //Ед. измерения
        public string Unit { get; set; }
        //Цена товара за единицу
        public decimal Price { get; set; }
    }
}
