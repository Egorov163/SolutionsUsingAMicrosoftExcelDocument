using ClosedXML.Excel;
using SolutionsUsingAMicrosoftExcelDocument.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SolutionsUsingAMicrosoftExcelDocument
{
    public class ExcelManager
    {
        private string _filePath;
        private XLWorkbook _workbook;

        public ExcelManager(string path)
        {
            _filePath = path;
            _workbook = new XLWorkbook(_filePath);
        }

        public List<Product> GetProducts()
        {
            var worksheet = _workbook.Worksheet("Товары");
            var products = new List<Product>();

            foreach (var row in worksheet.RowsUsed().Skip(1))
            {
                products.Add(new Product
                {
                    Code = row.Cell(1).GetValue<int>(),
                    Name = row.Cell(2).GetValue<string>(),
                    Unit = row.Cell(3).GetValue<string>(),
                    Price = row.Cell(4).GetValue<decimal>()
                });
            }
            return products;
        }

        public List<Client> GetClients()
        {
            var worksheet = _workbook.Worksheet("Клиенты");
            var clients = new List<Client>();

            foreach (var row in worksheet.RowsUsed().Skip(1))
            {
                clients.Add(new Client
                {
                    Code = row.Cell(1).GetValue<int>(),
                    Name = row.Cell(2).GetValue<string>(),
                    Address = row.Cell(3).GetValue<string>(),
                    ContactPerson = row.Cell(4).GetValue<string>()
                });
            }
            return clients;
        }

        public List<Order> GetOrders()
        {
            var worksheet = _workbook.Worksheet("Заявки");
            var orders = new List<Order>();

            var rows = worksheet.RowsUsed().Skip(1);
            foreach (var row in rows)
            {
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

            return orders;
        }

        public void UpdateClientContactPerson(string clientName, string newContactPerson)
        {
            var worksheet = _workbook.Worksheet("Клиенты");
            var clientRow = worksheet.RowsUsed().FirstOrDefault(r => r.Cell(2).GetValue<string>() == clientName);

            if (clientRow != null)
            {
                clientRow.Cell(4).Value = newContactPerson;
                _workbook.Save();
                Console.WriteLine($"Контактное лицо для {clientName} успешно обновлено.");
            }
            else
            {
                Console.WriteLine($"Клиент {clientName} не найден.");
            }
        }

    }
}
