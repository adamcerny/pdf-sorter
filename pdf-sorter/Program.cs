using CsvHelper;
using iText.Kernel.Pdf;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace pdf_sorter
{
    class Program
    {
        static void Main(string[] args)
        {

            var builder = new ConfigurationBuilder()
             .SetBasePath(Directory.GetCurrentDirectory())
             .AddJsonFile("appsettings.json");

            var config = builder.Build();

            string csvPath = config["csvPath"];
            string sourcePdfPath = config["sourcePdfPath"];
            string destinationPdfPath = config["destinationPdfPath"];
            List<CsvData> records = GetCsvRecords(csvPath);

            if (!IsCsvValid(records))
            {
                Console.WriteLine("CSV has bad data. Press 'y' if you would like to abort job.");
                ConsoleKeyInfo cki = Console.ReadKey();
                if (cki.Key.ToString().ToLower() == "y")
                    return;
            }

            if (File.Exists(destinationPdfPath))
            {
                Console.WriteLine("Destination pdf already exists. Press 'y' if you would like to overwrite it.");
                ConsoleKeyInfo cki = Console.ReadKey();
                if (cki.Key.ToString().ToLower() == "y")
                    File.Delete(destinationPdfPath);
            }

            CreateOrderedPdf(sourcePdfPath, destinationPdfPath, records);

        }

        private static void CreateOrderedPdf(string sourcePdfPath, string destinationPdfPath, List<CsvData> records)
        {
            Console.WriteLine("Loading PDF...");
            PdfDocument newPdfDoc = new PdfDocument(new PdfWriter(destinationPdfPath).SetSmartMode(true));
            Console.WriteLine("PDF Loaded");

            PdfDocument srcDoc = new PdfDocument(new PdfReader(sourcePdfPath));
            var pdfReader = new PdfReader(sourcePdfPath);

            foreach (var record in records.OrderBy(r => r.Date))
            {
                Console.WriteLine($"Copying pages {record.PageFrom} to {record.PageTo} from {record.Date:MM/dd/yy}");
                srcDoc.CopyPagesTo(record.PageFrom, record.PageTo, newPdfDoc);
            }

            newPdfDoc.Close();
        }

        private static List<CsvData> GetCsvRecords(string csvPath)
        {
            List<CsvData> records;

            using (var csvReader = new StreamReader(csvPath))
            using (var csv = new CsvReader(csvReader, CultureInfo.InvariantCulture))
            {
                Console.WriteLine("Loading CSV...");
                records = csv.GetRecords<CsvData>().ToList();
                Console.WriteLine("CSV Loaded");
            }

            return records;
        }

        private static bool IsCsvValid(List<CsvData> records)
        {
            bool validCsv = true;

            for (int i = 0; i < records.Count; i++)
            {
                if (i > 0)
                {
                    if (records[i].PageTo < records[i].PageFrom ||
                        records[i].PageFrom - records[i - 1].PageTo != 1)
                    {
                        Console.WriteLine($"Bad data found on row {i + 2}");
                        validCsv = false;
                    }
                }
            }

            return validCsv;
        }
    }
}
