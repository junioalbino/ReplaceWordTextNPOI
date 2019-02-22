using NPOI.OpenXml4Net.OPC;
using NPOI.XWPF.UserModel;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ReplaceWordText
{
    class Program
    {
        static void Main(string[] args)
        {
            var doc = new XWPFDocument(OPCPackage.Open("input.docx"));
            doc.RemoveParagraphs("BEGIN", "END");
            doc.Write(new FileStream("output.docx", FileMode.Create));
        }
    }

    static class SWPFDocumentExtensions
    {
        public static void RemoveParagraphs(this XWPFDocument doc, string beginTag, string endTag)
        {
            var remove = false;

            int i = 0;
            while (i < doc.BodyElements.Count)
            {
                if (!(doc.BodyElements[i] is XWPFParagraph))
                {
                    i++;
                    continue;
                }

                var runText = (doc.BodyElements[i] as XWPFParagraph).Text;

                if (runText == beginTag)
                    remove = true;

                if (remove)
                    doc.RemoveBodyElement(i);

                if (runText == endTag)
                {
                    remove = false;
                    continue;
                }

                if (!remove)
                    i++;
            }
        }
    }
}