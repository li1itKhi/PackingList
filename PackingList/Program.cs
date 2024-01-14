using CsvHelper.Configuration.Attributes;

namespace PackingList
{
    public class Program
    {
        private static string _path = "data.csv"; //путь для CSV файла
        private static string _pathXlsx = "Packing.xlsx";// путь для эксель файла
        static void Main(string[] args)
        {
            CSVWork.ReadCsvFile(_path);
            ExcelWork.ConvertCsvToExcel(_path, _pathXlsx);
            
        }
    }

    public class PackingList : IComparable<PackingList>
    {
        [Name("Название товара")]
        public string? Name { get; set; }

        [Name("Количество")]
        public int Quantity { get; set; }

        [Name("Стоимость")]
        public double Price { get; set; }

        [Name("ФИО поставщика")]
        public string? NameSupplier { get; set; }

        [Name("ФИО получателя")]
        public string? NameRecipient { get; set; }

        [Name("Дата поставки")]
        public DateTime DateSupply { get; set; }

        public PackingList() { } //конструктор без параметров

        public PackingList(string? name, int quantity, double price, 
            string? nameSupplier, string? nameRecipient, DateTime dateSupply)
        {
            Name = name;
            Quantity = quantity;
            Price = price;
            NameSupplier = nameSupplier;
            NameRecipient = nameRecipient;
            DateSupply = dateSupply;
        }
        public int CompareTo(PackingList obj)
        {
            return DateSupply.CompareTo(obj.DateSupply);
        }
    }
}