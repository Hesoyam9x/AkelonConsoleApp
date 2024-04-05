using AkelonConsoleApp;

Console.WriteLine("1) Укажите путь до файла с данными:");
string path = Console.ReadLine();
Context ct = new Context(path);
Console.WriteLine();

Console.WriteLine("2) Введите полное наименование товара:");
string productName = Console.ReadLine();
List<String> clientsRes = ct.ClientInfo(productName);
clientsRes.ForEach(x => Console.WriteLine(x));
Console.WriteLine();


Console.WriteLine("3) Введите наименование организации:");
string companyName = Console.ReadLine();
ct.EditClientInfo(companyName);


Console.WriteLine("4) Введите год и месяц, что бы определить золотого клиента:");
Console.Write("Месяц: ");
int month = Int32.Parse(Console.ReadLine());
Console.Write("Год: ");
int year = Int32.Parse(Console.ReadLine());
KeyValuePair<string, int> goldenClient = ct.GoldenClient(month, year);
if (goldenClient.Key != null)
    Console.WriteLine($"Наименование организации: {goldenClient.Key}, кол-во заказов за месяц: {goldenClient.Value}");
else Console.WriteLine("За указанный месяц не было заказов");