using HtmlAgilityPack;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace LeoSpider
{
    class Program
    {
        static async Task Main(string[] args)
        {
            string fileName = "Leo.xlsx";
            string extension = Path.GetExtension(fileName);
            if (extension == ".xlsx" || extension == ".xls")
            {
                IWorkbook wb;
                using (var fs = new FileStream(fileName, FileMode.Open, FileAccess.Read))
                {
                    if (extension == ".xlsx")
                    {
                        wb = new XSSFWorkbook(fs);
                    }
                    else
                    {
                        wb = new HSSFWorkbook(fs);
                    }
                }

                var sheet = wb.GetSheetAt(0);
                for (int i = sheet.FirstRowNum; i <= sheet.LastRowNum; i++)
                {
                    var row = sheet.GetRow(i);
                    if (row == null) continue;
                    var cell = row.GetCell(0);
                    string company = cell.ToString();
                    string phoneNumber = await GetPhoneNumber(company);
                    row.CreateCell(1).SetCellValue(phoneNumber);
                    Console.WriteLine($"{company} {phoneNumber}");
                }

                using (var fs = new FileStream(fileName, FileMode.Create, FileAccess.Write))
                {
                    wb.Write(fs);
                }

                Process.Start("Leo.xlsx");
            }
            else
            {
                Console.WriteLine("File invalid");
            }
        }

        private static async Task<string> GetPhoneNumber(string company)
        {
            using (var client = new HttpClient())
            {
                var html = await client.GetStringAsync($"https://www.google.com/search?q={company} 電話");
                var doc = new HtmlDocument();
                doc.LoadHtml(html);
                var phoneNumber = doc.DocumentNode.SelectNodes(@"//*[@id=""main""]/div[3]/div/div[3]/div/div/div/div/div/div/div/div/div")?.FirstOrDefault()?.InnerText ?? "Not found";
                return phoneNumber;
            }
        }
    }
}
