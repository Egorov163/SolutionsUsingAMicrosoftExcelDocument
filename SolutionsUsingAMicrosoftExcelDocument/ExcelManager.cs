using ClosedXML.Excel;
using SolutionsUsingAMicrosoftExcelDocument.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SolutionsUsingAMicrosoftExcelDocument
{
    public class ExcelManager
    {
        // Поле для хранения пути к файлу Excel
        private string _filePath;
        // Поле для хранения экземпляра XLWorkbook
        private XLWorkbook _workbook;

        /// <summary>
        /// Создание ExcelManager
        /// </summary>
        /// <param name="path"> Путь к файлу </param>
        public ExcelManager(string path)
        {
            _filePath = path;
            _workbook = new XLWorkbook(_filePath);
        }

        /// <summary>
        /// Получение списка продуктов из таблицы "Товары"
        /// </summary>
        public List<Product> GetProducts()
        {
            // Получение рабочего листа "Товары"
            var worksheet = _workbook.Worksheet("Товары");
            // Создание списка продуктов
            var products = new List<Product>();

            // Итерация по всем строкам рабочего листа, начиная со второй (первая строка - заголовок)
            foreach (var row in worksheet.RowsUsed().Skip(1))
            {
                // Создание нового объекта Product и заполнение его данными из соответствующих ячеек строки
                products.Add(new Product
                {
                    Code = row.Cell(1).GetValue<int>(),
                    Name = row.Cell(2).GetValue<string>(),
                    Unit = row.Cell(3).GetValue<string>(),
                    Price = row.Cell(4).GetValue<decimal>()
                });
            }
            // Возврат списка продуктов
            return products;
        }

        /// <summary>
        /// Получение списка клиентов из таблицы "Клиенты"
        /// </summary>
        public List<Client> GetClients()
        {
            // Получение рабочего листа "Клиенты"
            var worksheet = _workbook.Worksheet("Клиенты");
            // Создание списка клиентов
            var clients = new List<Client>();

            // Итерация по всем строкам рабочего листа, начиная со второй (первая строка - заголовок)
            foreach (var row in worksheet.RowsUsed().Skip(1))
            {
                // Создание нового объекта Client и заполнение его данными из соответствующих ячеек строки
                clients.Add(new Client
                {
                    Code = row.Cell(1).GetValue<int>(),
                    Name = row.Cell(2).GetValue<string>(),
                    Address = row.Cell(3).GetValue<string>(),
                    ContactPerson = row.Cell(4).GetValue<string>()
                });
            }
            // Возврат списка клиентов
            return clients;
        }

        /// <summary>
        /// Получение списка заказов из таблицы "Заявки"
        /// </summary>
        public List<Order> GetOrders()
        {
            // Получение рабочего листа "Заявки"
            var worksheet = _workbook.Worksheet("Заявки");
            // Создание списка заказов
            var orders = new List<Order>();

            // Итерация по всем строкам рабочего листа, начиная со второй (первая строка - заголовок)
            var rows = worksheet.RowsUsed().Skip(1);
            foreach (var row in rows)
            {
                // Создание нового объекта Order и заполнение его данными из соответствующих ячеек строки
                orders.Add(new Order
                {
                    Code = row.Cell(1).GetValue<int>(),
                    ProductCode = row.Cell(2).GetValue<int>(),
                    ClientCode = row.Cell(3).GetValue<int>(),
                    OrderNumber = row.Cell(4).GetValue<int>(),
                    Quantity = row.Cell(5).GetValue<int>(),
                    Date = row.Cell(6).GetValue<DateTime>()
                });
            }
            // Возврат списка заказов
            return orders;
        }

        /// <summary>
        /// Обновление контактного лица клиента в таблице "Клиенты"
        /// </summary>
        /// <param name="clientName">Имя клиента</param>
        /// <param name="newContactPerson">Новое контактное лицо</param>
        public void UpdateClientContactPerson(string clientName, string newContactPerson)
        {
            // Получение рабочего листа "Клиенты"
            var worksheet = _workbook.Worksheet("Клиенты");
            // Поиск строки клиента по его наименованию
            var clientRow = worksheet.RowsUsed().FirstOrDefault(r => r.Cell(2).GetValue<string>() == clientName);

            // Если строка клиента найдена
            if (clientRow != null)
            {
                // Обновление значения ячейки с контактным лицом в найденной строке
                clientRow.Cell(4).Value = newContactPerson;
                // Сохранение изменений в книге Excel
                _workbook.Save();
                // Вывод сообщения об успешном обновлении
                Console.WriteLine($"Контактное лицо для {clientName} успешно обновлено.");
            }
            else
            {
                // Вывод сообщения о том, что клиент не найден
                Console.WriteLine($"Клиент {clientName} не найден.");
            }
        }

    }
}
