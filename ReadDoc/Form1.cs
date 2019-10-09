using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Email;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ReadDoc
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            DataBind();
        }

        static List<OpenXmlElement> cElements = new List<OpenXmlElement>();

        string folderPath = @"C:\Users\akari\Downloads\RedPDF\ReadDoc\Documents\6letters";
        string pdffolderPath = @"C:\Users\akari\Downloads\RedPDF\ReadDoc\Documents\dest";

        private void DataBind()
        {

            DirectoryInfo di = new DirectoryInfo(folderPath);
            FileInfo[] files = di.GetFiles("*.docx");
            foreach (var item in files)
            {
                comboBox1.Items.Add(item.Name);
            }

            DirectoryInfo di1 = new DirectoryInfo(pdffolderPath);
            FileInfo[] files1 = di1.GetFiles("*.pdf");
            ddDest.Items.Clear();
            foreach (var item1 in files1)
            {
                ddDest.Items.Add(item1.Name);
            }

        }
        private void button1_Click(object sender, EventArgs e)
        {
            //DialogResult result = openFileDialog1.ShowDialog();
            //if (result == DialogResult.OK) // Test result.
            //{
            //Do whatever you want
            //openFileDialog1.FileName .....
            DirectoryInfo di = new DirectoryInfo(folderPath);
            DirectoryInfo di1 = new DirectoryInfo(pdffolderPath);
            string selectedFileName = comboBox1.SelectedItem.ToString();

            List<string> nTexts = new List<string>();
            cElements.Clear();
            lstResult.Items.Clear();
            int pageCount = 0;
            Dictionary<int, string> pageviseContent = new Dictionary<int, string>();
            try
            {
                var filepath = System.IO.Path.Combine(di.FullName, selectedFileName);// EGR_Exh25B.docx");
                lstResult.Items.Add("Extracting data from " + filepath);
                string selectedCondition = ddCondition.SelectedItem.ToString();
                switch (selectedCondition.Trim())
                {
                    case "Get Text By Color":
                        pageCount = ExtractColorData(nTexts, pageCount, pageviseContent, filepath);
                        break;
                    case "Get Text Between Tags":
                       // GetTagsData(nTexts, pageCount, pageviseContent, filepath);
                         ExtractTagsData(nTexts, pageCount, pageviseContent, filepath);
                        break;
                    case "IF Condition":
                        pageCount = GetTextBySearch(nTexts, pageCount, pageviseContent, filepath);
                        break;
                    case "Get Static Text":
                      var  ldata = ExtractDataFromDoc_V1(nTexts, pageCount, pageviseContent, filepath);
                        var pdfFilepath = System.IO.Path.Combine(di1.FullName, ddDest.SelectedItem.ToString());
                        string pdfText = PdfExtract.ReadPDFFile(pdfFilepath);
                        foreach (var item in ldata)
                        {
                            string ss = item.Text;//.Replace('.',' ');
                          var ddd=  PdfExtract.CheckTextInFile(pdfText, ss);
                            lstResult.Items.Add("Text:-" + item.Text + " result:- " + ddd.ToString());
                        }
                      

                        break;
                    default:
                        break;
                }

            }
            catch (Exception ex)
            {

                lstResult.Items.Add("Error While Extracting data   " + ex.Message);
            }

            //  }
        }



        private int ExtractDataFromDoc(List<string> nTexts, int pageCount, Dictionary<int, string> pageviseContent, string filepath)
        {
            WordprocessingDocument wordDocument = WordprocessingDocument.Open(filepath, true);
            Body body = wordDocument.MainDocumentPart.Document.Body;
            if (wordDocument.ExtendedFilePropertiesPart.Properties.Pages.Text != null)
            {
                pageCount = Convert.ToInt32(wordDocument.ExtendedFilePropertiesPart.Properties.Pages.Text);
            }
            int i = 1;



            StringBuilder pageContentBuilder = new StringBuilder();
            foreach (OpenXmlElement element in body.ChildElements)
            {

                if (element.InnerXml.IndexOf("<w:br w:type=\"page\" />", StringComparison.OrdinalIgnoreCase) < 0)
                {
                    //pageContentBuilder.Append(element.InnerText );
                    string result = DataExtractUtilities.RemoveTagsFromData(element.InnerText);

                    pageContentBuilder.Append(result);
                    if (i > 1)
                    {
                        if (!string.IsNullOrWhiteSpace(result.Trim()))
                        {
                            if (element.HasChildren)
                            {
                                nTexts.AddRange(DataExtractUtilities.GetTextWithoutColor(element));
                            }
                            else
                            {
                                if (!DataExtractUtilities.IsTextHasColor(element))
                                {
                                    nTexts.AddRange(DataExtractUtilities.GetNonStrikeTextWithOutTags(element));
                                }
                            }
                            //  Console.WriteLine(element.InnerText);


                        }
                    }
                }
                else
                {
                    pageviseContent.Add(i, pageContentBuilder.ToString());
                    i++;
                    pageContentBuilder = new StringBuilder();
                }
                if (body.LastChild == element && pageContentBuilder.Length > 0)
                {
                    pageviseContent.Add(i, pageContentBuilder.ToString());
                }
            }
            int tagStart = 0;
            foreach (var ntextItem in nTexts)
            {
                // Console.WriteLine(ntextItem);
                if (ntextItem == ">")
                {
                    tagStart = 0;
                }
                else if (ntextItem == "<")
                {
                    tagStart = 1;
                }
                else
                {
                    if (tagStart != 1)
                    {
                        if (ntextItem.Trim() != "," && ntextItem.Trim() != "." && ntextItem.Trim() != ":" && ntextItem.Trim() != "")
                        {
                            lstResult.Items.Add(ntextItem);
                        }
                    }
                }

            }

            return pageCount;
        }

        private List<ExtractText> ExtractDataFromDoc_V1(List<string> nTexts, int pageCount, Dictionary<int, string> pageviseContent, string filepath)
        {
            List<ExtractText> result = new List<ExtractText>();
            lstResult.Items.Clear();
            string StartTag = "<";
            string EndTag = ">";
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filepath, true))
            {
                Body body = wordDocument.MainDocumentPart.Document.Body;
                if (wordDocument.ExtendedFilePropertiesPart.Properties.Pages.Text != null)
                {
                    pageCount = Convert.ToInt32(wordDocument.ExtendedFilePropertiesPart.Properties.Pages.Text);
                }
                int i = 0;

                StringBuilder pageContentBuilder = new StringBuilder();
                List<OpenXmlElement> cEles = new List<OpenXmlElement>();
                foreach (OpenXmlElement element in body.ChildElements)
                {

                    if (element.InnerXml.IndexOf("<w:br w:type=\"page\" />", StringComparison.OrdinalIgnoreCase) > 0)
                    {
                        i =1;
                        // string result = DataExtractUtilities.GetTagData(element.InnerText);
                       // cEles = DataExtractUtilities.GetAllChildElements(element, cElements);
                    }
                    
                    if (i == 1)
                    {
                        cEles = DataExtractUtilities.GetAllChildElements(element, cElements);
                    }
                    
                }

                int startTagCount = 0;
                string sb = "";
                foreach (var elem in cEles)
                {
                    List<string> rItems = new List<string>();
                    string resultData = elem.InnerText; //DataExtractUtilities.GetTagsData(elem);


                    if (!string.IsNullOrWhiteSpace(resultData))
                    {
                        if (IsTextField(elem))
                        {
                            string color = DataExtractUtilities.GetTextColor(elem);
                            bool IsTextStrike = DataExtractUtilities.IsTextStrike(elem);
                            if (string.IsNullOrWhiteSpace(color)  && !IsTextStrike)
                            {
                                //lstResult.Items.Add(elem.InnerText);
                                result.Add(new ExtractText()
                                {
                                    Text = elem.InnerText,
                                    Type = ExtractType.Text.ToString()
                                });


                            }
                        }

                    }
                }

            }

            return result;
        }

        private void GetTagsData(List<string> nTexts, int pageCount, Dictionary<int, string> pageviseContent, string filepath)
        {
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filepath, true))
            {
                Body body = wordDocument.MainDocumentPart.Document.Body;
                List<string> s1 = body.InnerText.Split('<').ToList();

                int i = 0;
                foreach (var item in s1)
                {
                    if (i > 0)
                    {
                        string f1 = item.Split('>')[0];
                        lstResult.Items.Add(f1);
                    }
                    i++;
                }
            }
        }


        private int ExtractTagsData(List<string> nTexts, int pageCount, Dictionary<int, string> pageviseContent, string filepath)
        {
            lstResult.Items.Clear();
            string  StartTag = "<";
            string  EndTag = ">";
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filepath, true))
            {
                Body body = wordDocument.MainDocumentPart.Document.Body;
                if (wordDocument.ExtendedFilePropertiesPart.Properties.Pages.Text != null)
                {
                    pageCount = Convert.ToInt32(wordDocument.ExtendedFilePropertiesPart.Properties.Pages.Text);
                }
                int i = 1;

                StringBuilder pageContentBuilder = new StringBuilder();
                List<OpenXmlElement> cEles = new List<OpenXmlElement>();
                foreach (OpenXmlElement element in body.ChildElements)
                {

                    if (element.InnerXml.IndexOf("<w:br w:type=\"page\" />", StringComparison.OrdinalIgnoreCase) < 0)
                    {
                        // string result = DataExtractUtilities.GetTagData(element.InnerText);
                        cEles = DataExtractUtilities.GetAllChildElements(element, cElements);
                    }
                    else
                    {
                        pageviseContent.Add(i, pageContentBuilder.ToString());
                        i++;
                        pageContentBuilder = new StringBuilder();
                    }
                    if (body.LastChild == element && pageContentBuilder.Length > 0)
                    {
                        pageviseContent.Add(i, pageContentBuilder.ToString());
                    }
                }

                int startTagCount = 0;
                string sb = "";
                foreach (var elem in cEles)
                {
                    List<string> rItems = new List<string>();
                    string resultData = elem.InnerText; //DataExtractUtilities.GetTagsData(elem);


                    if (!string.IsNullOrWhiteSpace(resultData))
                    {
                        if (IsTextField(elem) )
                        {
                            string color = DataExtractUtilities.GetTextColor(elem);
                            bool IsTextStrike = DataExtractUtilities.IsTextStrike(elem);
                            if (color == "green" && !IsTextStrike)
                            {
                                if(resultData.IndexOf('<') != -1)
                                {
                                    startTagCount++;
                                }
                                if (resultData.IndexOf('>') != -1)
                                {
                                    sb += resultData;
                                    startTagCount =0;
                                }
                                if(resultData.IndexOf('<') != -1 && resultData.IndexOf('>') != -1)
                                {
                                    sb = resultData;
                                }
                            
                                if (startTagCount == 0)
                                {
                                    lstResult.Items.Add(sb);
                                    sb = "";
                                }
                                else
                                {
                                    sb += resultData;
                                }
                                
                            }
                        }
                       
                    }
                } 

            }

            return pageCount;
        }

         

        private bool IsTextField(OpenXmlElement ele)
        {
            bool isText = false;
            switch (ele.LocalName)
            {
                case "t":
                    isText = true;
                    break;
                default:
                    break;
            }
            return isText;
        }

        private int ExtractColorData(List<string> nTexts, int pageCount, Dictionary<int, string> pageviseContent, string filepath)
        {

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filepath, true))
            {
                Body body = wordDocument.MainDocumentPart.Document.Body;
                if (wordDocument.ExtendedFilePropertiesPart.Properties.Pages.Text != null)
                {
                    pageCount = Convert.ToInt32(wordDocument.ExtendedFilePropertiesPart.Properties.Pages.Text);
                }
                int i = 1;

                StringBuilder pageContentBuilder = new StringBuilder();
                foreach (OpenXmlElement element in body.ChildElements)
                {

                    if (element.InnerXml.IndexOf("<w:br w:type=\"page\" />", StringComparison.OrdinalIgnoreCase) < 0)
                    {
                        //pageContentBuilder.Append(element.InnerText );
                        string result = DataExtractUtilities.GetTagData(element.InnerText);
                        var isMatchColor = DataExtractUtilities.IsMatchColor(element, "green");
                        if (isMatchColor)
                            lstResult.Items.Add("Green Color Text --> " + result);

                        if (i > 1)
                        {
                            if (!string.IsNullOrWhiteSpace(result.Trim()) && isMatchColor)
                            {
                                //  Console.WriteLine(element.InnerText);
                                nTexts.Add(element.InnerText);
                            }
                        }
                    }
                    else
                    {
                        pageviseContent.Add(i, pageContentBuilder.ToString());
                        i++;
                        pageContentBuilder = new StringBuilder();
                    }
                    if (body.LastChild == element && pageContentBuilder.Length > 0)
                    {
                        pageviseContent.Add(i, pageContentBuilder.ToString());
                    }
                }

                // Console.WriteLine("pageContentBuilder=", pageContentBuilder.ToString());
                foreach (var ntextItem in nTexts)
                {
                    // Console.WriteLine(ntextItem);
                    if (ntextItem.Trim() != "," && ntextItem.Trim() != "." && ntextItem.Trim() != ":" && ntextItem.Trim() != "")
                    {
                        lstResult.Items.Add(ntextItem);
                    }
                }
            }
            return pageCount;
        }

        private int GetTextBySearch(List<string> nTexts, int pageCount, Dictionary<int, string> pageviseContent, string filepath)
        {
            StringBuilder sb = new StringBuilder();
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filepath, true))
            {
                Body body = wordDocument.MainDocumentPart.Document.Body;
                if (wordDocument.ExtendedFilePropertiesPart.Properties.Pages.Text != null)
                {
                    pageCount = Convert.ToInt32(wordDocument.ExtendedFilePropertiesPart.Properties.Pages.Text);
                    // Console.WriteLine("*** pageCount --> " + pageCount);
                }
                int i = 1;
                int ifCount = 0;
                StringBuilder pageContentBuilder = new StringBuilder();
                foreach (OpenXmlElement element in body.ChildElements)
                {
                    Console.WriteLine("***InnerText  --> " + element.InnerText);

                    String[] startCondition = new String[] { "[IF BUSINESSTYPE = ‘MAPD’" };
                    String[] endCondition = new String[] { " END IF]" };
                    var isMatchCondition = DataExtractUtilities.IsMatchCondition(element, startCondition);
                    if (isMatchCondition)
                    {
                        ifCount++;
                        sb.AppendFormat(startCondition + DataExtractUtilities.ExtractConditionText(element, startCondition));
                        lstResult.Items.Add(startCondition[0].ToString() + DataExtractUtilities.ExtractConditionText(element, startCondition));
                    }


                    var isMatchConditionEnd = DataExtractUtilities.IsMatchCondition(element, endCondition);
                    if (isMatchConditionEnd)
                    {
                        sb.Append(DataExtractUtilities.ExtractEndConditionText(element, endCondition) + endCondition);
                        ifCount = 0;
                        lstResult.Items.Add(sb.ToString());
                        //Console.WriteLine("*** 2. Condition Text --> " + element.InnerText);
                    }
                    //if (ifCount == 0)
                    //    Console.WriteLine("111 ***---- Text --> " + sb.ToString());
                    /*
                    if (element.InnerXml.IndexOf("<w:br w:type=\"page\" />", StringComparison.OrdinalIgnoreCase) < 0)
                    {
                        //pageContentBuilder.Append(element.InnerText );
                        string result =  element.InnerText;
                        var isMatchCondition = DataExtractUtilities.IsMatchCondition(element, "[IF ");
                         if (isMatchCondition)
                        Console.WriteLine("*** 1. Condition Text --> " + result);


                        var isMatchConditionEnd = DataExtractUtilities.IsMatchCondition(element, " END IF]");
                        if (isMatchConditionEnd)
                        {
                            Console.WriteLine("*** 2. Condition Text --> " + element.InnerText);
                        }

                            if (i > 1)
                        {
                            if (!string.IsNullOrWhiteSpace(result.Trim()) && isMatchCondition)
                            {
                                //  Console.WriteLine(element.InnerText);
                                nTexts.Add(element.InnerText);
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine("*** Page Number --> " + i);
                        pageviseContent.Add(i, pageContentBuilder.ToString());
                        i++;
                        pageContentBuilder = new StringBuilder();
                    }
                    if (body.LastChild == element && pageContentBuilder.Length > 0)
                    {
                        pageviseContent.Add(i, pageContentBuilder.ToString());
                    }*/
                }

                // Console.WriteLine("pageContentBuilder=", pageContentBuilder.ToString());
                foreach (var ntextItem in nTexts)
                {
                    // Console.WriteLine(ntextItem);
                    if (ntextItem.Trim() != "," && ntextItem.Trim() != "." && ntextItem.Trim() != ":" && ntextItem.Trim() != "")
                    {
                        lstResult.Items.Add(ntextItem);
                    }
                }
            }
            return 0;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            string basepath = Class1.GetBasePath();

            string selectedPdfFileName = ddDest.SelectedItem.ToString();
            string pdfFile = System.IO.Path.Combine(pdffolderPath, selectedPdfFileName);
            ExtractTextFromPdf(pdfFile);
        }
        public static string ExtractTextFromPdf(string path)
        {
            //            ITextExtractionStrategy its = new iTextSharp.text.pdf.parser.LocationTextExtractionStrategy();
            ITextExtractionStrategy its = new iTextSharp.text.pdf.parser.SimpleTextExtractionStrategy();

            using (PdfReader reader = new PdfReader(path))
            {
                StringBuilder text = new StringBuilder();

                for (int i = 1; i <= reader.NumberOfPages; i++)
                {
                    string thePage = PdfTextExtractor.GetTextFromPage(reader, i, its);
                    string[] theLines = thePage.Split('\n');
                    foreach (var theLine in theLines)
                    {
                        text.Append(" "+theLine);
                    }
                }
                return text.ToString();
            }
        }
    }

    public class ExtractText
    {
        public string Text { get; set; }
        public string Type { get; set; }
    }
    public enum ExtractType
    {
        Text,
        Tag,
        condition
    }
}
