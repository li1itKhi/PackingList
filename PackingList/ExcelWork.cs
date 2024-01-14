using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;


namespace PackingList
{
    internal class ExcelWork
    {
        public static void ConvertCsvToExcel(string csvFilePath, string xlsxFilePath)
        {
            FileInfo csvFile = new FileInfo(csvFilePath);
            FileInfo newFile = new FileInfo(xlsxFilePath);

            using (ExcelPackage package = new ExcelPackage(newFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Накладные");

                string[] csvLines = File.ReadAllLines(csvFilePath);
                int rowIndex = 1;

                foreach (string line in csvLines)
                {
                    string[] fields = line.Split(','); // разделитель - запятая
                    int columnIndex = 1;

                    foreach (string field in fields)
                    {
                        worksheet.Cells[rowIndex, columnIndex].Value = field;
                        columnIndex++;
                    }

                    rowIndex++;
                }

                package.Save();
            }
        }

        public void GetData(string excelFilePath)
        {
            List<PackingList> goodsList = new List<PackingList>();

            FileInfo existingFile = new FileInfo(excelFilePath);

            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++) //первая строка - заголовки
                {
                    PackingList goods = new PackingList(
                        worksheet.Cells[row, 1].Value.ToString(),
                        int.Parse(worksheet.Cells[row, 2].Value.ToString()),
                        double.Parse(worksheet.Cells[row, 3].Value.ToString()),
                        worksheet.Cells[row, 4].Value.ToString(),
                        worksheet.Cells[row, 5].Value.ToString(),
                        DateTime.Parse(worksheet.Cells[row, 6].Value.ToString())
                    );
                    goodsList.Add(goods);
                }

                FormWaybill(goodsList);
            }

        }

        public void FormWaybill(List<PackingList> data)
        {
            Dictionary<string, List<PackingList>> compliances = new Dictionary<string, List<PackingList>>();

            HashSet<DateTime> dates = new HashSet<DateTime>();

            //Группировка товаров по получателям
            foreach (var item in data)
            {
                if (!compliances.ContainsKey(item.NameRecipient!))
                {
                    compliances.Add(item.NameRecipient, new List<PackingList>());
                }
                
                compliances[item.NameRecipient].Add(item);

                dates.Add(item.DateSupply);
            }

            //Формирование накладной для каждого получателя на каждую уникальную дату
            foreach (var date in dates)
            {
                foreach (var receiver in compliances.Keys)
                {
                    List<PackingList> temp = new List<PackingList>();

                    foreach (var item in data)
                    {
                        if (item.NameRecipient == receiver && item.DateSupply == date)
                        {
                            temp.Add(item);
                        }
                    }
                    WriteWaybill(temp);
                   
                    temp.Clear();
                }
            }
        }

        public void WriteWaybill(List<PackingList> temp)
        {
            Random random = new Random();//для генерации номера товарной накладной

            double summ = 0; //для подсчета суммы пришедших товаров

            // Преобразуем дату и получателя в строку для формирования имени файла
            string fileName = $"Накладная от {temp[0].DateSupply:yyyy.MM.dd}. Получатель - {temp[0].NameRecipient}.docx"; //!!!!!!!!!!!!!!!!

            // Имя файла-образца
            string templateFilePath = "Товарная накладная.docx";

            // Копирование файла-образца с новым именем
            File.Copy(templateFilePath, fileName);

            Word.Application wordApp = new Word.Application(); //открываем приложение

            Word.Document doc = wordApp.Documents.Open(fileName); //файл для заполнения

            doc.Content.Find.Execute("<NUM>", ReplaceWith: $"{random.Next(100000, 1000000)}");
            doc.Content.Find.Execute("<DATE>", ReplaceWith: $"{temp[0].DateSupply.ToShortTimeString()}");
            doc.Content.Find.Execute("<PROVIDER>", ReplaceWith: $"{temp[0].NameSupplier}");
            doc.Content.Find.Execute("<RECIPIENT>", ReplaceWith: $"{temp[0].NameRecipient}");
            doc.Content.Find.Execute("<RECIPIENT>", ReplaceWith: $"{temp[0].NameRecipient}");

            Word.Table table = doc.Tables[1]; //таблица

            for (int i = 1; i < temp.Count; i++)
            {
                table.Rows.Add();
                table.Cell(i + 1, 1).Range.Text = i.ToString();
                table.Cell(i + 1, 2).Range.Text = temp[i].Name;
                table.Cell(i + 1, 3).Range.Text = temp[i].Quantity.ToString();
                table.Cell(i + 1, 4).Range.Text = temp[i].Price.ToString();

                double total = temp[i].Quantity * temp[i].Price; //подсчет суммы

                table.Cell(i + 1, 5).Range.Text = total.ToString();

                summ += total; //общее
            }

            doc.Content.Find.Execute("<SUMM>", ReplaceWith: $"{summ}");
            doc.Content.Find.Execute("<SUMM>", ReplaceWith: $"{summ}");
            doc.Content.Find.Execute("<QUANTITY>", ReplaceWith: $"{temp.Count}");

            doc.SaveAs2(fileName);

            Console.WriteLine($"Файл {fileName} создан");
        }
    }
}
