using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Google.Apis.Docs.v1;
using Google.Apis.Docs.v1.Data;
using Google.Apis.Services;

namespace FileViewer.Services
{

    public class GoogleDocsService
    {
        private DocsService docsService = null;
        public GoogleDocsService(GoogleOAuthService googleOAuthService)
        {
            docsService = new DocsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = googleOAuthService.GetCurrentCredential(),
                ApplicationName = "My Project 53610",
            });
        }

        public void test(string documentId)
        {
            DocumentsResource.GetRequest request = docsService.Documents.Get(documentId);
            Document doc = request.Execute();

            Footer f = new Footer();
            StructuralElement se = new StructuralElement();
            Paragraph p = new Paragraph();
            ParagraphElement pe = new ParagraphElement();
            AutoText at = new AutoText();

            at.Type = "PAGE_NUMBER";
            at.TextStyle = new TextStyle();


            pe.AutoText = at;
            pe.StartIndex = 1;
            pe.EndIndex = 10;
            pe.ColumnBreak = new ColumnBreak();

            p.Elements = new List<ParagraphElement>()
            {
                pe
            };

            se.Paragraph = p;
            se.StartIndex = 0;
            se.EndIndex = 0;
            se.SectionBreak = new SectionBreak();
            se.Table = new Table();
            se.TableOfContents = new TableOfContents();

            f.Content = new List<StructuralElement>() { se };
            //f.Content.Add(se);
            doc.Footers = new Dictionary<string, Footer>();
            doc.Footers.Add(new KeyValuePair<string, Footer>("test", f));

            BatchUpdateDocumentRequest budr = new BatchUpdateDocumentRequest();
            Request r = new Request();
            CreateFooterRequest cfr = new CreateFooterRequest();
            Location l = new Location();

            l.SegmentId = "test";
            l.
            cfr.SectionBreakLocation = l;
            r.CreateFooter = cfr;
            budr.Requests = new List<Request>() { r };
            docsService.Documents.BatchUpdate(budr, "1fIrVDmqz_3nJIRikd-YD4vTDwp6a_FlmDNQNqqpTu6A?id=e6258ffe-c09d-4ef9-a83f-b8fc6ad9c836");
        }
    }
}
