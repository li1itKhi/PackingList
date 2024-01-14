using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection.PortableExecutable;
using System.Text;
using System.Threading.Tasks;
using CsvHelper;
using CsvHelper.Configuration;

namespace PackingList
{
    public class CSVWork
    {
        public static List<PackingList> ReadCsvFile(string path)
        {
            StreamReader reader = new StreamReader(path);

            CsvReader csvReader = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture));

            List<PackingList> packingListCSV = csvReader.GetRecords<PackingList>().ToList();

            List<PackingList> temp = new List<PackingList>(); //временный

            foreach (PackingList packingList in packingListCSV)
            {
                bool isTrue = false;

                for (int i = 0; i < temp.Count; i++)
                {
                    if (temp[i].Name == packingList.Name && temp[i].NameRecipient == packingList.NameRecipient
                        && temp[i].DateSupply == packingList.DateSupply)
                    {
                        temp[i].Quantity += packingList.Quantity;
                        isTrue = true;
                    }
                }

                if (!isTrue)
                {
                    temp.Add(packingList);
                }
            }

            temp.Sort(); //сортировка

            foreach (PackingList packList in temp) //считывание файла
            {
                Console.WriteLine($"{packList.Name}\t" +
                    $"{packList.Quantity}\t" +
                    $"{packList.Price}\t" +
                    $"{packList.NameSupplier}\t" +
                    $"{packList.NameRecipient}\t" +
                    $"{packList.DateSupply}");
            }
            return packingListCSV;
        }


    }
}
