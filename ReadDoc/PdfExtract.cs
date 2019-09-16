using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;

namespace ReadDoc
{
    public class PdfExtract
    {
        public static string ReadPDFFile(string fileName)
        {
            StringBuilder text = new StringBuilder();

            if (File.Exists(fileName))
            {
                PdfReader pdfReader = new PdfReader(fileName);

                for (int page = 1; page <= pdfReader.NumberOfPages; page++)
                {
                    ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
                    string currentText = PdfTextExtractor.GetTextFromPage(pdfReader, page);
                    string[] lines = currentText.Split('\n');
                    foreach (string line in lines)
                    {
                        string[] words = line.Split('\n');
                        foreach (string wrd in words)
                        {

                        }
                        text.Append(" "+line);
                    }
                  
                }
                pdfReader.Close();
            }
            return text.ToString();
        }

        public static bool CheckTextInFile(string pdfContent,string vText)
        {
            string newPdfCOntent = pdfContent.Replace("  ", " ");
            return newPdfCOntent.Contains(vText.Trim().Replace("  ", " "));
        }

    }
}
