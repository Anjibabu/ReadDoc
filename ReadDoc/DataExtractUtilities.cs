using System.Xml.Linq;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using System;
using System.Collections.Generic;

namespace ReadDoc
{
    public class DataExtractUtilities
    {
        public static string pattern = @"<(.*?)>";

        public static void GetPageNumbers(string filePath)
        {
            using (var document = WordprocessingDocument.Open(filePath, false))
            {
                var paragraphInfos = new List<string>();

                var paragraphs = document.MainDocumentPart.Document.Descendants<Paragraph>();

                int pageIdx = 1;
                foreach (var paragraph in paragraphs)
                {
                    var run = paragraph.GetFirstChild<Run>();

                    if (run != null)
                    {
                        var lastRenderedPageBreak = run.GetFirstChild<LastRenderedPageBreak>();
                        var pageBreak = run.GetFirstChild<Break>();
                        if (lastRenderedPageBreak != null || pageBreak != null)
                        {
                            pageIdx++;
                        }
                    }
                    Console.WriteLine("page number " + pageIdx, paragraph.InnerText);
                    // paragraphInfos.Add(info);
                }

            }
        }


        public static List<string> GetNonStrikeTextWithOutTags(OpenXmlElement element)
        {
            List<string> normalText = new List<string>();
            var childElements = element.ChildElements;
            foreach (var cElement in childElements)
            {
                if (cElement.FirstChild != null)
                {
                    if (!string.IsNullOrWhiteSpace(cElement.InnerText))
                    {
                        var localName = cElement.FirstChild.LocalName;
                        switch (localName)
                        {
                            case "t":
                                //Console.WriteLine("*****" + cElement.InnerText + "*****");
                                normalText.Add(cElement.InnerText);
                                break;
                            case "rPr":
                                if (((RunProperties)cElement.FirstChild).Strike == null)
                                {
                                    string result = RemoveTagsFromData(cElement.InnerText);

                                    normalText.Add(result);
                                }
                                break;
                        }

                    }
                }
            }

            return normalText;
        }

        public static List<string> GetNonStrikeTextWithOutTagsWithOutColor(OpenXmlElement element)
        {
            List<string> normalText = new List<string>();
            var childElements = element.ChildElements;
            foreach (var cElement in childElements)
            {
                if (cElement.FirstChild != null)
                {
                    if (!string.IsNullOrWhiteSpace(cElement.InnerText))
                    {
                        var localName = cElement.FirstChild.LocalName;
                        switch (localName)
                        {
                            case "t":
                                //Console.WriteLine("*****" + cElement.InnerText + "*****");
                                normalText.Add(cElement.InnerText);
                                break;
                            case "rPr":
                                if (((RunProperties)cElement.FirstChild).Strike == null)
                                {
                                    string result = RemoveTagsFromData(cElement.InnerText);

                                    normalText.Add(result);
                                }
                                break;
                        }

                    }
                }
            }

            return normalText;
        }

        public static List<string> GetNonStrikeTextWithTags(OpenXmlElement element)
        {
            List<string> normalText = new List<string>();
            var childElements = element.ChildElements;
            foreach (var cElement in childElements)
            {
                if (cElement.FirstChild != null)
                {
                    if (!string.IsNullOrWhiteSpace(cElement.InnerText))
                    {
                        var localName = cElement.FirstChild.LocalName;
                        switch (localName)
                        {
                            case "t":
                                //Console.WriteLine("*****" + cElement.InnerText + "*****");
                                normalText.Add(cElement.InnerText);
                                break;
                            case "rPr":
                                if (((RunProperties)cElement.FirstChild).Strike == null)
                                {
                                    string result = cElement.InnerText;

                                    normalText.Add(result);
                                }
                                break;
                        }

                    }
                }
            }

            return normalText;
        }

        public static int GetPageNumber(OpenXmlElement elem, OpenXmlElement root)
        {
            int pageNbr = 1;
            var tmpElem = elem;
            while (tmpElem != root)
            {
                var sibling = tmpElem.PreviousSibling();
                while (sibling != null)
                {
                    pageNbr += sibling.Descendants<LastRenderedPageBreak>().Count();
                    sibling = sibling.PreviousSibling();
                }
                tmpElem = tmpElem.Parent;
            }
            return pageNbr;
        }

        public static string RemoveTagsFromData(string inputData)
        {
            Regex rgx = new Regex(pattern);
            return rgx.Replace(inputData, "");
        }

        public static string GetTagData(string inputData)
        {
            Regex rgx = new Regex(pattern);

            if (rgx.IsMatch(inputData))
            {
                //Console.WriteLine("Tagas Data-->" + inputData);
                return inputData;
            }
            else
            {
                return "";
            }
        }

        public static string[] GetTagsData(string inputData, List<string> rItems)
        {
            Regex rgx = new Regex(pattern);

            foreach (Match match in Regex.Matches(inputData, pattern))
            {
                rItems.Add(match.Groups[1].Value);
            }

            return rItems.ToArray();
        }

        public static string  GetTagsData(OpenXmlElement element )
        {
            var inputData = element.InnerText;
            Regex rgx = new Regex(pattern);

            if (rgx.IsMatch(inputData))
            {
                //Console.WriteLine("Tagas Data-->" + inputData);
                return inputData;
            }
            else
            {
                return "";
            }
        }
        public static string[] GetTagsListData(OpenXmlElement element, List<string> rItems)
        {
            var childElements = element.ChildElements;
            foreach (var cElement in childElements)
            {
                if (cElement.HasChildren)
                {
                    GetTagsListData(cElement, rItems);
                }
                else
                {
                    GetTagsData(cElement.InnerText, rItems);
                }

            }
            return rItems.ToArray();
        }

        static List<OpenXmlElement> cElements = new List<OpenXmlElement>();
        public static List<OpenXmlElement> GetAllChildElements(OpenXmlElement element)
        {
            var childElements = element.ChildElements;
            foreach (var cElement in childElements)
            {
                if (cElement.HasChildren)
                {
                    GetAllChildElements(cElement);
                }
                else
                {
                    cElements.Add(cElement);
                }

            }
            return cElements;
        }

        public static bool IsMatchColor(OpenXmlElement element, string color)
        {
            bool isMatchColor = false;
            var childElements = element.ChildElements;
            foreach (var cElement in childElements)
            {
                if (cElement.FirstChild != null)
                {
                    if (!string.IsNullOrWhiteSpace(cElement.InnerText))
                    {
                        var localName = cElement.FirstChild.LocalName;
                        switch (localName)
                        {

                            case "rPr":
                                if (((RunProperties)cElement.FirstChild).Highlight != null)
                                {
                                    if (((RunProperties)cElement.FirstChild).Highlight.Val == color)
                                    {
                                        isMatchColor = true;
                                    }
                                }
                                break;
                        }

                    }
                }
            }

            return isMatchColor;
        }

        public static bool IsTextHasColor(OpenXmlElement element)
        {
            bool isColor = false;
            var childElements = element.ChildElements;
            foreach (var cElement in childElements)
            {
                if (cElement.FirstChild != null)
                {
                    if (!string.IsNullOrWhiteSpace(cElement.InnerText))
                    {
                        var localName = cElement.FirstChild.LocalName;
                        switch (localName)
                        {

                            case "rPr":
                                if (((RunProperties)cElement.FirstChild).Highlight != null)
                                {
                                    if (((RunProperties)cElement.FirstChild).Highlight.Val != null )
                                    {
                                        isColor = true;
                                    }
                                }
                                break;
                        }

                    }
                }
            }

            return isColor;
        }

        public static List<string> GetTextWithoutColor(OpenXmlElement element)
        {
            bool isColor = false;
            List<string> text = new List<string>();
            var childElements = element.ChildElements;
            foreach (var cElement in childElements)
            {
                if (cElement.FirstChild != null)
                {
                    if (!string.IsNullOrWhiteSpace(cElement.InnerText))
                    {
                        var localName = cElement.FirstChild.LocalName;
                        switch (localName)
                        {

                            case "rPr":
                                if (((RunProperties)cElement.FirstChild).Highlight != null)
                                {
                                    if (((RunProperties)cElement.FirstChild).Highlight.Val == null)
                                    {
                                        isColor = true;
                                       
                                    }
                                }
                                else
                                {
                                    if (((RunProperties)cElement.FirstChild).Strike == null)
                                    {
                                        text.Add(cElement.InnerText);
                                    }
                                   
                                }
                                break;
                        }

                    }
                }
            }

            return text;
        }


        public static bool IsMatchCondition(OpenXmlElement element, String[] condition)
        {
            bool isMatchColor = false;

            var innertext = element.InnerText;
            if (innertext.Contains(condition[0]))
            {
                isMatchColor = true;
            }

            return isMatchColor;
        }

        public static string ExtractConditionText(OpenXmlElement element, String[] condition)
        {
            var innertext = element.InnerText;
            return innertext.Split(condition, StringSplitOptions.None)[1];
        }

        public static string ExtractEndConditionText(OpenXmlElement element, string[] condition)
        {
            var innertext = element.InnerText;
            return innertext.Split(condition, StringSplitOptions.None)[0];
        }
    }
}
