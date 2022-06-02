using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace Matriz_KeyWords_Generator
{
    class Program
    {
        static string RelativePath;
        static void Main(string[] args)
        {
            Console.WriteLine("Starting process.");
            SetRelativePath();

            List<string> wordsList = GetWordsFromCSV();
            int countList = wordsList.Count;
            DirectoryInfo di = new DirectoryInfo($"{RelativePath}Arquivos");
            FileInfo[] files = di.GetFiles("*.pdf");

            int[,] table_AND = new int[countList, countList];
            int[,] table_NOT = new int[countList, countList];

            int totalFiles = files.Length;
            int count = 0;

            foreach (var file in files)
            {
                var pdfContent = ExtractTextFromPDF(file.FullName).ToLower();

                var text = TreatText(pdfContent);

                for (int i = 0; i < countList; i++)
                    if (ContainsInfo(text, wordsList[i]))
                    {
                        for (int j = 0; j < countList; j++)
                        {
                            bool containsInDoc = ContainsInfo(text, wordsList[j]);

                            // se contêm em ambos dará match nas palavras e adicionará no AND
                            if (containsInDoc && i != j)
                                table_AND[i, j] = table_AND[i, j] + 1;
                            // se contêm no primeiro e não contêm no segundo adicionará no NOT
                            else if (!containsInDoc && i != j)
                                table_NOT[i, j] = table_NOT[i, j] + 1;
                        }
                    }
                    else
                    {
                        for (int j = 0; j < countList; j++)
                        {
                            bool containsInDoc = ContainsInfo(text, wordsList[j]);

                            // se contêm em ambos dará match nas palavras
                            if (containsInDoc && i != j)
                                table_NOT[i, j] = table_NOT[i, j] + 1;
                        }
                    }

                count++;
                Console.Write("\r{0}   ", (count * 100 / totalFiles + "%"));
            }

            ExportExcel(table_AND, wordsList, "MatrizKeywords_AND");

            ExportExcel(table_NOT, wordsList, "MatrizKeywords_NOT");
        }

        private static string TreatText(string text)
        {
            var internalTtext = text;

            internalTtext = internalTtext.Replace("-", "").Replace("\n", "").Replace("\r", "");

            return internalTtext;
        }

        private static void SetRelativePath()
        {
#if DEBUG
            RelativePath = @"C:\Users\gabri\projetos\Matrix_KeyWords_Generator\";
#else
            RelativePath = AppDomain.CurrentDomain.BaseDirectory;
#endif
        }
        private static bool ContainsInfo(string pdfContent, string word)
        {
            var compareInfo = CultureInfo.InvariantCulture.CompareInfo;

            //Identifica se um significado tem mais de uma palavra separada por vírgula no arquivo.
            foreach (var wordItem in word.Split(','))
            {
                string treatedWord = wordItem.TrimStart().TrimEnd();

                if (compareInfo.IndexOf(pdfContent, treatedWord, CompareOptions.IgnoreNonSpace) > -1)
                    return true;
            }


            return false;
        }

        private static List<string> GetWordsFromCSV()
        {
            List<string> wordsList = new List<string>();
            string[] words = File.ReadAllLines($"{RelativePath}Words.csv");

            foreach (var word in words)
                wordsList.Add(word);
            return wordsList;

        }
        private static string ExtractTextFromPDF(string filePath)
        {
            //depois implementar leitura ocr para identificar artigos escaneados
            //https://itextpdf.com/en/blog/technical-notes/how-use-itext-pdfocr-recognize-text-scanned-documents

            string pageContent = "";

            PdfReader pdfReader = new PdfReader(filePath);
            PdfDocument pdfDoc = new PdfDocument(pdfReader);

            for (int page = 1; page <= pdfDoc.GetNumberOfPages(); page++)
            {
                try
                {
                    ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();

                    pageContent += PdfTextExtractor.GetTextFromPage(pdfDoc.GetPage(page), strategy) + " ";
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }

            }
            pdfDoc.Close();
            pdfReader.Close();

            return pageContent;
        }

        private static void ExportExcel(int[,] table, List<string> wordsList, string fileName)
        {
            DateTime date = DateTime.Now;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var file = new FileInfo(fileName: $"{RelativePath}Arquivos_Gerados\\{fileName}_{date.Day}_{date.Hour}_{date.Year}_{date.Hour}_{date.Minute}_{date.Second}" + ".xlsx");

            using var package = new ExcelPackage(file);

            var ws = package.Workbook.Worksheets.Add("MatrizKeywords");

            //set columns header
            for (int i = 1; i <= wordsList.Count; i++)
                ws.Cells[1, i + 1].Value = wordsList[i - 1];


            //set rows header
            for (int i = 1; i <= wordsList.Count; i++)
                ws.Cells[i + 1, 1].Value = wordsList[i - 1];


            //set results
            for (int i = 1; i <= wordsList.Count; i++)
                for (int j = 1; j <= wordsList.Count; j++)
                    ws.Cells[i + 1, j + 1].Value = table[i - 1, j - 1];

            package.Save();
        }
    }

}
