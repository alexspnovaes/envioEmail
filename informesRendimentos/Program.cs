using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace informesRendimentos
{
    class Program
    {
        static void Main(string[] args)
        {
            var path = @"c:\excel";
            if (Directory.Exists(path))
            {
                ProcessDirectory(path);
            }

           
        }

        private static void ProcessDirectory(string path)
        {
            var ext = new List<string> { "xls", "xlsx"};
            string[] fileEntries =
                Directory
                .GetFiles(path)
                .Where(s => ext
                    .Contains(
                        Path
                        .GetExtension(s)
                        .TrimStart('.')
                        .ToLowerInvariant()))
                .ToArray();
            foreach (string filePath in fileEntries)
                ProcessFile(filePath);
        }

        private static void ProcessFile(string filePath)
        {
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

                using var reader = ExcelReaderFactory.CreateReader(stream);
                do
                {
                    while (reader.Read())
                    {
                         reader.GetDouble(0);
                    }
                } while (reader.NextResult());
                
                var result = reader.AsDataSet();
                // The result of each spreadsheet is in result.Tables
            }
        }
    }
}
