using org.pdfclown.documents;
using org.pdfclown.documents.contents;
using org.pdfclown.documents.contents.objects;
using org.pdfclown.files;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;


namespace PDF_Demo.Helper
{

    public class InvoiceTextExtraction
    {
        private List<string> _contentList;

        public void GetValueFromPdf()
        {
            _contentList = new List<string>();
            CreatePdfContent(@"C:empInvoice1.pdf");

            var index = _contentList.FindIndex(e => e == "INVOICE") + 1;
            int.TryParse(_contentList[index], out var value);
            Console.WriteLine(value);
        }


        public void CreatePdfContent(string filePath)
        {
            using (var file = new File(filePath))
            { 
                Document document = file.Document;
                foreach (var page in document.Pages)
                {
                    Extract(new ContentScanner(page));
                }
            }
        }

        private void Extract(ContentScanner level)
        {
            if (level == null)
                return;

            while (level.MoveNext())
            {
                var content = level.Current;
                switch (content)
                {
                    case ShowText text:
                        {
                            var font = level.State.Font;
                            _contentList.Add(font.Decode(text.Text));
                            break;
                        }
                    case Text _:
                    case ContainerObject _:
                        Extract(level.ChildLevel);
                        break;
                }
            }
        }
    }
}
   