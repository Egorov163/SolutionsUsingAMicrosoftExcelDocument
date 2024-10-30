using System;
using System.Linq;

namespace SolutionsUsingAMicrosoftExcelDocument
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Введите путь к Excel файлу: ");
            string filePath = Console.ReadLine();

            var excelManager = new ExcelManager(filePath);

            while (true)
            {
                Console.WriteLine("\nВыберите действие:");
                Console.WriteLine("1. Информация о клиентах по наименованию товара");
                Console.WriteLine("2. Изменить контактное лицо клиента");
                Console.WriteLine("3. Определить золотого клиента");
                Console.WriteLine("4. Выход");

                string choice = Console.ReadLine();

                switch (choice)
                {
                    case "1":
                        GetClientInfoByProduct(excelManager);
                        break;
                    case "2":
                        UpdateClientContactPerson(excelManager);
                        break;
                    case "3":
                        GetGoldenClient(excelManager);
                        break;
                    case "4":
                        return;
                    default:
                        Console.WriteLine("Неверный выбор. Попробуйте снова.");
                        break;
                }
            }
        }
        static void GetClientInfoByProduct(ExcelManager excelManager)
        {
            Console.Write("Введите наименование товара: ");
            string productName = Console.ReadLine();

            var products = excelManager.GetProducts();
            var clients = excelManager.GetClients();
            var orders = excelManager.GetOrders();

            var product = products.FirstOrDefault(p => p.Name.Equals(productName, StringComparison.OrdinalIgnoreCase));

            if (product == null)
            {
                Console.WriteLine("Товар не найден.");
                return;
            }

            var clientOrders = orders.Where(o => o.ProductCode == product.Code)
                .Join(clients,
                    o => o.ClientCode,
                    c => c.Code,
                    (o, c) => new { Order = o, Client = c });

            foreach (var clientOrder in clientOrders)
            {
                Console.WriteLine($"Клиент: {clientOrder.Client.Name}");
                Console.WriteLine($"Количество: {clientOrder.Order.Quantity}");
                Console.WriteLine($"Цена: {product.Price * clientOrder.Order.Quantity} рублей");
                Console.WriteLine($"Дата заказа: {clientOrder.Order.Date:d}");
                Console.WriteLine();
            }
        }
        static void UpdateClientContactPerson(ExcelManager excelManager)
        {
            Console.Write("Введите название организации: ");
            string clientName = Console.ReadLine();

            Console.Write("Введите ФИО нового контактного лица: ");
            string newContactPerson = Console.ReadLine();

            excelManager.UpdateClientContactPerson(clientName, newContactPerson);
        }
        static void GetGoldenClient(ExcelManager excelManager)
        {
            Console.Write("Введите год: ");
            int year = int.Parse(Console.ReadLine());

            Console.Write("Введите месяц (1-12): ");
            int month = int.Parse(Console.ReadLine());

            var orders = excelManager.GetOrders();
            var clients = excelManager.GetClients();

            var goldenClient = orders.Where(o => o.Date.Year == year && o.Date.Month == month)
                .GroupBy(o => o.ClientCode)
                .OrderByDescending(g => g.Count())
                .FirstOrDefault();

            if (goldenClient == null)
            {
                Console.WriteLine("Нет заказов за указанный период.");
                return;
            }

            var client = clients.First(c => c.Code == goldenClient.Key);
            Console.WriteLine($"Золотой клиент: {client.Name}");
            Console.WriteLine($"Количество заказов: {goldenClient.Count()}");
        }
    }
}
