using System;
using System.Linq;

namespace SolutionsUsingAMicrosoftExcelDocument
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Запрос у пользователя пути к файлу Excel
            Console.WriteLine("Введите путь к Excel файлу: ");
            string filePath = Console.ReadLine();

            // Создание экземпляра класса ExcelManager
            var excelManager = new ExcelManager(filePath);

            // Основной цикл взаимодействия с пользователем
            while (true)
            {
                // Вывод доступных действий пользователю
                Console.WriteLine("\nВыберите действие:");
                Console.WriteLine("1. Информация о клиентах по наименованию товара");
                Console.WriteLine("2. Изменить контактное лицо клиента");
                Console.WriteLine("3. Определить золотого клиента");
                Console.WriteLine("4. Выход");

                // Получение выбора пользователя
                string choice = Console.ReadLine();

                // Switch-конструкция для обработки различных выборов пользователя
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
        #region Доп. методы
        /// <summary>
        /// Получение информации о клиентах по наименованию товара
        /// </summary>
        static void GetClientInfoByProduct(ExcelManager excelManager)
        {
            // Запрос у пользователя наименования товара
            Console.Write("Введите наименование товара: ");
            string productName = Console.ReadLine();

            // Получение данных из файла Excel
            var products = excelManager.GetProducts();
            var clients = excelManager.GetClients();
            var orders = excelManager.GetOrders();

            // Поиск товара по его наименованию (без учета регистра)
            var product = products.FirstOrDefault(p => p.Name.Equals(productName, StringComparison.OrdinalIgnoreCase));

            // Обработка случая, когда товар не найден
            if (product == null)
            {
                Console.WriteLine("Товар не найден.");
                return;
            }

            // Поиск заказов, связанных с указанным товаром, и объединение их с информацией о клиентах
            var clientOrders = orders.Where(o => o.ProductCode == product.Code)
                .Join(clients,
                    o => o.ClientCode,
                    c => c.Code,
                    (o, c) => new { Order = o, Client = c });

            // Итерация по заказам клиентов и вывод информации
            foreach (var clientOrder in clientOrders)
            {
                Console.WriteLine($"Клиент: {clientOrder.Client.Name}");
                Console.WriteLine($"Количество: {clientOrder.Order.Quantity}");
                Console.WriteLine($"Цена: {product.Price * clientOrder.Order.Quantity} рублей");
                Console.WriteLine($"Дата заказа: {clientOrder.Order.Date:d}");
                Console.WriteLine();
            }
        }
        /// <summary>
        /// Обновление контактного лица клиента
        /// </summary>
        static void UpdateClientContactPerson(ExcelManager excelManager)
        {
            // Запрос у пользователя названия организации и нового контактного лица
            Console.Write("Введите название организации: ");
            string clientName = Console.ReadLine();

            Console.Write("Введите ФИО нового контактного лица: ");
            string newContactPerson = Console.ReadLine();

            // Обновление контактного лица клиента в файле Excel
            excelManager.UpdateClientContactPerson(clientName, newContactPerson);
        }

        /// <summary>
        /// Определение золотого клиента
        /// </summary>
        static void GetGoldenClient(ExcelManager excelManager)
        {
            // Запрос у пользователя года и месяца
            Console.Write("Введите год: ");
            int year = int.Parse(Console.ReadLine());

            Console.Write("Введите месяц (1-12): ");
            int month = int.Parse(Console.ReadLine());

            // Получение данных из файла Excel
            var orders = excelManager.GetOrders();
            var clients = excelManager.GetClients();

            // Поиск клиента с наибольшим количеством заказов за указанный год и месяц
            var goldenClient = orders.Where(o => o.Date.Year == year && o.Date.Month == month)
                .GroupBy(o => o.ClientCode)
                .OrderByDescending(g => g.Count())
                .FirstOrDefault();

            // Обработка случая, когда заказы за указанный период не найдены
            if (goldenClient == null)
            {
                Console.WriteLine("Нет заказов за указанный период.");
                return;
            }

            // Получение информации о золотом клиенте из списка клиентов
            var client = clients.First(c => c.Code == goldenClient.Key);
            Console.WriteLine($"Золотой клиент: {client.Name}");
            Console.WriteLine($"Количество заказов: {goldenClient.Count()}");
        }
        #endregion
    }
}
