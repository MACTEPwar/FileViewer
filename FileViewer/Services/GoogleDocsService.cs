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

        //localhost:20657/api/file/test/1fIrVDmqz_3nJIRikd-YD4vTDwp6a_FlmDNQNqqpTu6A
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
            f.FooterId = "test";
            //f.Content.Add(se);
            //doc.Footers = new Dictionary<string, Footer>();
            //doc.Footers.Add(new KeyValuePair<string, Footer>("test", f));

            //BatchUpdateDocumentRequest budr = new BatchUpdateDocumentRequest();
            //Request r = new Request();
            //CreateFooterRequest cfr = new CreateFooterRequest();
            //Location l = new Location();

            //l.SegmentId = null;
            //l.Index = 5;
            //cfr.SectionBreakLocation = l;
            //r.CreateFooter = cfr;
            //budr.Requests = new List<Request>() { r };
            //docsService.Documents.BatchUpdate(budr, "1fIrVDmqz_3nJIRikd-YD4vTDwp6a_FlmDNQNqqpTu6A?id=e6258ffe-c09d-4ef9-a83f-b8fc6ad9c836");

            var requests = new BatchUpdateDocumentRequest()
            {
                Requests = new List<Request>()
            };

            if (doc.Footers.Count > 0)
            {
                foreach (var footer in doc.Footers)
                {
                    requests.Requests.Add(new Request()
                    {
                        DeleteFooter = new DeleteFooterRequest()
                        {
                            FooterId = footer.Key
                        }
                    });
                }

            }

            requests.Requests.Add(new Request()
            {
                CreateFooter = new CreateFooterRequest()
                {
                    Type = "DEFAULT",
                    SectionBreakLocation = new Location()
                    {
                        Index = 0,
                        SegmentId = "footer1"
                    }
                }
            });

            docsService.Documents.BatchUpdate(requests, documentId).Execute();

            DocumentsResource.GetRequest request2 = docsService.Documents.Get(documentId);
            Document doc2 = request.Execute();

            docsService.Documents.BatchUpdate(new BatchUpdateDocumentRequest()
            {
                Requests = new List<Request>() { new Request()
                {
                    InsertText = new InsertTextRequest()
                    {
                        Text = "test",
                         Location = new Location()
                         {
                             Index = 0,
                             SegmentId = doc2.Footers.First().Key
                         }
                    },
                } }
            }, documentId).Execute();

        }

        //1fIrVDmqz_3nJIRikd-YD4vTDwp6a_FlmDNQNqqpTu6A
        //localhost:20657/api/file/test2/1fIrVDmqz_3nJIRikd-YD4vTDwp6a_FlmDNQNqqpTu6A
        public void test2(string documentId)
        {
            docsService.Documents.BatchUpdate(new BatchUpdateDocumentRequest()
            {
                Requests = new List<Request>()
                {
                    new Request()
                    {
                        InsertText = new InsertTextRequest()
                        {
                            Text = "THIS INSERTED TEXT",
                            Location = new Location()
                            {
                                Index = 5,
                                SegmentId = null
                            }
                        }
                    }
                }
            }, documentId).Execute();
        }

        public void insertImage(string documentId)
        {
            DocumentsResource.GetRequest request = docsService.Documents.Get(documentId);
            Document doc2 = request.Execute();

            docsService.Documents.BatchUpdate(new BatchUpdateDocumentRequest()
            {
                Requests = new List<Request>()
                {
                    new Request()
                    {
                        InsertInlineImage = new InsertInlineImageRequest()
                        {
                            Uri = "https://www.ixbt.com/img/n1/news/2022/6/6/32a52ad-google_large.jpg",
                            Location = new Location()
                            {
                                Index = 1,
                                SegmentId = null
                            }
                        }
                    }
                }
            }, documentId).Execute();
        }
    }
}
