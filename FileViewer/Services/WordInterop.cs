using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.IO;

namespace FileViewer.Services
{
    public class WordInterop
    {
        //localhost:20657/api/file/test3/
        public void OpenDocument()
        {
            //WordprocessingDocument wordprocessingDocument =
            //WordprocessingDocument.Open("./Files/test.docx", true);
            //WordprocessingDocument myDoc = WordprocessingDocument.Create("./Files/doc.docx", WordprocessingDocumentType.Document);
            //myDoc.AddMainDocumentPart();

            //// Create the Document DOM. 
            //myDoc.MainDocumentPart.Document =
            //  new Document(
            //    new Body(
            //      new Paragraph(
            //        new Run(
            //          new Text("Hello World!")))));
            //myDoc.MainDocumentPart.Document.Save();
            //myDoc.Close();

            using (FileStream mem = new FileStream("test.docx", FileMode.Create, FileAccess.Write))
            {
                // Create Document
                using (WordprocessingDocument wordDocument =
                    WordprocessingDocument.Create("./Files/doc.docx", WordprocessingDocumentType.Document))
                {
                    // Add a main document part. 
                    MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

                    // Create the document structure and add some text.
                    mainPart.Document = new Document();
                    Body docBody = new Body();

                    // Add your docx content here

                    Paragraph p = new Paragraph();
                    // Run 1
                    Run r1 = new Run();
                    Text t1 = new Text("Pellentesque ") { Space = SpaceProcessingModeValues.Preserve };
                    // The Space attribute preserve white space before and after your text
                    r1.Append(t1);
                    p.Append(r1);

                    wordDocument.MainDocumentPart.Document.Save();
                }
            }
        }
    }
}
