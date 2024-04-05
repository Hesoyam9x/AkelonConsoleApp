namespace AkelonConsoleApp
{
    class Product
    {
        public int Id { get; set; }
        public string Name { get;  set; }
        public string Unit { get; set; }
        public double Price { get; set; }
    }
    class Client
    {
        public int Id { get; set; }
        public string CompanyName { get; set; }
        public string Address { get; set; }
        public string ContactPerson { get; set; }
    }
    class Order
    {
        public int Id { get; set; }
        public int ProductId { get; set; }
        public int ClientId { get; set; }
        public int OrderNumber { get; set; }
        public int NeedCount { get; set; }
        public DateTime Date { get; set; }
    }
}
