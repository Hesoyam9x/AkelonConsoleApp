using ClosedXML.Excel;

namespace AkelonConsoleApp
{
    class Context
    {
        public List<Product> Products = new();
        public List<Client> Clients = new();
        public List<Order> Orders = new();

        private XLWorkbook workbook;

        private IXLRangeRows rows2;
        private IXLWorksheet worksheet2;
        public Context(string path)
        {
            workbook = new(path);
            WorkBookToLists(workbook);
        }

        private void WorkBookToLists(XLWorkbook workbook)
        {
            var worksheet1 = workbook.Worksheet(1);
            var rows = worksheet1.RangeUsed().RowsUsed();
            int ind = 0;
            foreach (var row in rows)
            {
                ind++;
                if (ind == 1) continue;
                Product product = new()
                {
                    Id = (int)row.Cell(1).Value,
                    Name = (string)row.Cell(2).Value,
                    Unit = (string)row.Cell(3).Value,
                    Price = (double)row.Cell(4).Value
                };
                Products.Add(product);
            }
            ind = 0;
            worksheet2 = workbook.Worksheet(2);
            rows2 = worksheet2.RangeUsed().RowsUsed();
            foreach (var row in rows2)
            {
                ind++;
                if (ind == 1) continue;
                Client client = new()
                {
                    Id = (int)row.Cell(1).Value,
                    CompanyName = (string)row.Cell(2).Value,
                    Address = (string)row.Cell(3).Value,
                    ContactPerson = (string)row.Cell(4).Value
                };
                Clients.Add(client);
            }
            ind = 0;
            var worksheet3 = workbook.Worksheet(3);
            var rows3 = worksheet3.RangeUsed().RowsUsed();
            foreach (var row in rows3)
            {
                ind++;
                if (ind == 1) continue;
                Order order = new()
                {
                    Id = (int)row.Cell(1).Value,
                    ProductId = (int)row.Cell(2).Value,
                    ClientId = (int)row.Cell(3).Value,
                    OrderNumber = (int)row.Cell(4).Value,
                    NeedCount = (int)row.Cell(5).Value,
                    Date = (DateTime)row.Cell(6).Value
                };
                Orders.Add(order);
            }
        }
        
        public List<String> ClientInfo(string productName)
        {
            List<string> strs = new();
            string resp = "";
            Product product = Products.Find(p => p.Name == productName);
            if (product != null)
            {
                List<Order> orderCl = Orders.Where(o => product.Id == o.ProductId).ToList();
                orderCl.ForEach(order =>
                {
                    string compName = "";
                    Clients.ForEach(client =>
                    {
                        if (client.Id == order.ClientId) compName = client.CompanyName;
                    });
                    resp = $"Клиент: {compName}. " + $"Количество: {order.NeedCount}шт. " + $"Цена: {product.Price}руб. " + $"Дата заказа: {order.Date.ToShortDateString()}";
                    strs.Add(resp);
                });
            }
            else
            {
                resp = "Нет такого товара";
                strs.Add(resp);
            }
            return strs;
        }

        public void EditClientInfo(string companyName)
        {
            Client client = Clients.Where(c => c.CompanyName.Contains(companyName)).FirstOrDefault();
            if(client != null)
            {
                Console.WriteLine($"Организация: {client.CompanyName}. Контактное лицо: {client.ContactPerson}");
                Console.WriteLine("Изменить контактное лицо? (Y/n)");

                string select = Console.ReadLine();
                if (select == "Y" || select == "y" || select == "")
                    Console.WriteLine("Введите имя нового контактного лица:");
                else return;

                string contactPerson = Console.ReadLine();
                Console.WriteLine($"{client.CompanyName}. Контактное лицо: {contactPerson}");

                int indexRow = 1;
                foreach (var row in rows2)
                {
                    string cell = row.Cell(2).Value.ToString();
                    if (cell.Contains(companyName)) break;
                    indexRow++;

                }
                IXLCell name = worksheet2.Cell(indexRow, 4);
                name.Value = contactPerson;
                workbook.Save();
                Console.WriteLine();
            }
            else
            {
                Console.WriteLine("Организации с таким названием нет в списке");
                Console.WriteLine();
            }
        }

        public KeyValuePair<string, int> GoldenClient(int month, int year)
        {
            List<Order> selectOrders = Orders
                .Where(o => o.Date.Month == month && o.Date.Year == year)
                .ToList();
            if (selectOrders.Count() == 0)
            {
                KeyValuePair<string, int> noResult = new();
                return noResult;
            }
            List<int> clientsId = selectOrders
                .GroupBy(x => x.ClientId)
                .Select(x => x.First().ClientId)
                .ToList();

            Dictionary<string, int> sortList = new();

            clientsId.ForEach(id =>
            {
                string name = Clients
                    .Where(c => c.Id == id)
                    .Select(n => n.CompanyName)
                    .FirstOrDefault();
                int sum = 0;
                selectOrders.ForEach(order =>
                {
                    if(id == order.ClientId) sum += order.NeedCount;
                });
                sortList.Add(name, sum);
            });

            KeyValuePair<string, int> result = sortList
                .OrderByDescending(x => x.Value)
                .First();
            return result;
        }
    }
}
