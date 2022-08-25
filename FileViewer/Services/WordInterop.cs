using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.Drawing;
using DocumentFormat.OpenXml.Office2010.Word.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Wordprocessing;
using FileViewer.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using A = DocumentFormat.OpenXml.Drawing;
//using V = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace FileViewer.Services
{
    public class WordInterop
    {
        //localhost:20657/api/file/test3/
        public void OpenDocument()
        {
            string filePath = @"./Files/test.docx";
            this.AddPageNumber(filePath);
        }

        public void replaceTextToImage()
        {
            string filePath = @"./Files/test.docx";

            using (WordprocessingDocument wordDocument =
            WordprocessingDocument.Open(filePath, true))
            {
                // формирую параграф с картинкой
                MainDocumentPart mainPart = wordDocument.MainDocumentPart;
                ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Png);
                Body body = wordDocument.MainDocumentPart.Document.Body;

                using (FileStream stream = new FileStream(@"./Files/test.png", FileMode.Open))
                {
                    imagePart.FeedData(stream);
                }

                Drawing p = AddImageToBody3(wordDocument, mainPart.GetIdOfPart(imagePart));

                // ищу вместо чего вставлять
                Regex regexText = new Regex("test");
                string finderText = "test";

                var texts = this.test(wordDocument.MainDocumentPart.Document);

                //foreach (var paragraph in wordDocument.MainDocumentPart.Document.Descendants<Paragraph>())
                foreach (var paragraph in texts)
                {
                    Match match = regexText.Match(paragraph.InnerText);
                    if (match.Success)
                    {
                        //create a new run and set its value to the correct text
                        //this must be done before the child runs are removed otherwise
                        //paragraph.InnerText will be empty

                        //newRun.AppendChild(new Text(paragraph.InnerText.Replace(match.Value, "some new value")));
                        //int ind = paragraph.InnerText.IndexOf(finderText);
                        string[] strs = paragraph.InnerText.Split(finderText);
                        List<Run> els = new List<Run>();

                        for (int i = 1; i < strs.Length; i++)
                        {
                            Run newRunBefore = new Run();
                            Run newRun = new Run();
                            Run newRunAfter = new Run();

                            newRunBefore.AppendChild(new Text(strs[i - 1]));
                            newRun.AppendChild(p);

                            els.Add(newRunBefore);
                            els.Add(newRun);

                            if (i == strs.Length - 1)
                            {
                                newRunAfter.AppendChild(new Text(strs[i]));
                                els.Add(newRunAfter);
                            }

                            //newRunAfter.AppendChild(new Text(strs[i]));
                        }


                        //remove any child runs
                        var parent = paragraph.Parent.Parent;
                        paragraph.Remove();
                        //parent.RemoveAllChildren<Run>();
                        //add the newly created run
                        foreach (var el in els)
                        {
                            parent.AppendChild(el);
                        }
                    }
                }

                // вставляю
                //body.InsertAfter(p, body.FirstChild);
                mainPart.Document.Save();

            }

            //using (WordprocessingDocument wordDoc =
            //WordprocessingDocument.Open(filePath, true))
            //{
            //    string docText = null;
            //    string image = null;
            //    using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
            //    {
            //        docText = sr.ReadToEnd();
            //    }

            //    using (FileStream stream = new FileStream(@"./Files/image1.png", FileMode.Open))
            //    {
            //        using (StreamReader sr = new StreamReader(stream))
            //        {
            //            image = sr.ReadToEnd();
            //        }
            //    }
            //    Regex regexText = new Regex("test");
            //    docText = regexText.Replace(docText, image);

            //    using (StreamWriter sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
            //    {
            //        sw.Write(docText);
            //    }
            //}
        }

        public List<OpenXmlElement> testFind()
        {
            string filePath = @"./Files/test.docx";

            using (WordprocessingDocument wordDocument =
            WordprocessingDocument.Open(filePath, true))
            {
                return this.test(wordDocument.MainDocumentPart.Document);
            }
        }

        public List<OpenXmlElement> test(Document doc)
        {
            string finder = "test";
            List<OpenXmlElement> result = new List<OpenXmlElement>();

            foreach (var paragraph in doc.Descendants<Paragraph>())
            {
                result.AddRange(this.find(finder, paragraph));
            }

            return result;

        }

        public List<OpenXmlElement> find(string str, OpenXmlElement element)
        {
            List<OpenXmlElement> result = new List<OpenXmlElement>();

            {
                //var innerText = element.InnerText;
                //var innerXml = element.InnerXml;
                //var texts = element.Descendants<Text>();
                if (element.LocalName == "t")
                {
                    result.Add(element);
                }
                //result.AddRange(texts);
            }

            foreach (var el in element.ChildElements)
            {
                List<OpenXmlElement> tempResult = this.find(str, el);
                result.AddRange(tempResult);
            }

            return result;
        }

        public void replaceText()
        {
            string filePath = @"./Files/test.docx";
            using (WordprocessingDocument wordDoc =
            WordprocessingDocument.Open(filePath, true))
            {
                string docText = null;
                using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                {
                    docText = sr.ReadToEnd();
                }
                Regex regexText = new Regex("test");
                docText = regexText.Replace(docText, "text");

                using (StreamWriter sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
                {
                    sw.Write(docText);
                }
            }
        }

        public void insertImage()
        {
            string filePath = @"./Files/test.docx";
            this.TTt(filePath);
        }

        public void Test()
        {
            string filePath = @"./Files/test.docx";
            //using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(filePath, true))
            {
                // Insert other code here. 
                //MainDocumentPart mainPart = wordprocessingDocument.MainDocumentPart;
                //ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Png);
                //using (FileStream stream = new FileStream(@"./Files/image1.png", FileMode.Open))
                //{
                //    imagePart.FeedData(stream);
                //}
                //AddImageToBody(wordprocessingDocument, mainPart.GetIdOfPart(imagePart));
                //AddImageToBody2(wordprocessingDocument, mainPart.GetIdOfPart(imagePart));


            }

            //this.TTt(filePath);
            using (WordprocessingDocument package = WordprocessingDocument.Open(filePath, true))
            {
                InsertCustomWatermark(package, @"./Files/image1.png");
            }
        }



















        private void InsertCustomWatermark(WordprocessingDocument package, string p)
        {
            SetWaterMarkPicture(p);
            MainDocumentPart mainDocumentPart1 = package.MainDocumentPart;
            if (mainDocumentPart1 != null)
            {
                mainDocumentPart1.DeleteParts(mainDocumentPart1.HeaderParts);
                HeaderPart headPart1 = mainDocumentPart1.AddNewPart<HeaderPart>();
                GenerateHeaderPart1Content(headPart1);
                string rId = mainDocumentPart1.GetIdOfPart(headPart1);
                ImagePart image = headPart1.AddNewPart<ImagePart>("image/jpeg", "rId999");
                GenerateImagePart1Content(image);
                IEnumerable<SectionProperties> sectPrs = mainDocumentPart1.Document.Body.Elements<SectionProperties>();
                foreach (var sectPr in sectPrs)
                {
                    sectPr.RemoveAllChildren<HeaderReference>();
                    sectPr.PrependChild<HeaderReference>(new HeaderReference() { Id = rId });
                }
            }
        }

        // для watermark
        private void GenerateHeaderPart1Content(HeaderPart headerPart1)
        {
            Header header1 = new Header();
            Paragraph paragraph2 = new Paragraph();
            Run run1 = new Run();
            Picture picture1 = new Picture();
            Shape shape1 = new Shape() { Id = "WordPictureWatermark75517470", Style = "position:absolute;left:0;text-align:left;margin-left:0;margin-top:0;width:415.2pt;height:456.15pt;z-index:-251656192;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin", OptionalString = "_x0000_s2051", AllowInCell = false, Type = "#_x0000_t75" };
            ImageData imageData1 = new ImageData() { Gain = "19661f", BlackLevel = "22938f", Title = "??", RelationshipId = "rId999" };
            shape1.Append(imageData1);
            picture1.Append(shape1);
            run1.Append(picture1);
            paragraph2.Append(run1);
            header1.Append(paragraph2);
            headerPart1.Header = header1;
        }

        // для watermark
        private void GenerateImagePart1Content(ImagePart imagePart1)
        {
            System.IO.Stream data = GetBinaryDataStream(imagePart1Data);
            imagePart1.FeedData(data);
            data.Close();
        }

        // для watermark
        private static string imagePart1Data = "";

        // для watermark
        private static System.IO.Stream GetBinaryDataStream(string base64String)
        {
            return new System.IO.MemoryStream(System.Convert.FromBase64String(base64String));
        }

        // для watermark
        public static void SetWaterMarkPicture(string file)
        {
            FileStream inFile;
            try
            {
                inFile = new FileStream(file, FileMode.Open, FileAccess.Read);
                byte[] byteArray = new byte[inFile.Length];
                long byteRead = inFile.Read(byteArray, 0, (int)inFile.Length);
                inFile.Close();
                imagePart1Data = Convert.ToBase64String(byteArray, 0, byteArray.Length);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }








































        private void TTt(string path)
        {
            string filepath = path;
            // Open a WordprocessingDocument based on a filepath.
            Dictionary<int, string> pageWithContent = new Dictionary<int, string>();
            int pageCount = 0;
            using (WordprocessingDocument wordDocument =
                WordprocessingDocument.Open(filepath, true))
            {
                // Assign a reference to the existing document body.  
                Body body = wordDocument.MainDocumentPart.Document.Body;
                //if (wordDocument.ExtendedFilePropertiesPart.Properties.Pages.Text != null)
                //{
                //    pageCount = Convert.ToInt32(wordDocument.ExtendedFilePropertiesPart.Properties.Pages.Text);
                //}
                //int i = 1;
                //StringBuilder pageContentBuilder = new StringBuilder();
                //foreach (var element in body.ChildElements)
                //{
                //    if (element.InnerXml.IndexOf("<w:br w:type=\"page\" />", StringComparison.OrdinalIgnoreCase) < 0)
                //    {
                //        pageContentBuilder.Append(element.InnerText);
                //    }
                //    else
                //    {
                //        pageWithContent.Add(i, pageContentBuilder.ToString());
                //        i++;
                //        pageContentBuilder = new StringBuilder();
                //    }
                //    if (body.LastChild == element && pageContentBuilder.Length > 0)
                //    {
                //        pageWithContent.Add(i, pageContentBuilder.ToString());
                //    }
                //}

                MainDocumentPart mainPart = wordDocument.MainDocumentPart;
                ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Png);
                using (FileStream stream = new FileStream(@"./Files/image1.png", FileMode.Open))
                {
                    imagePart.FeedData(stream);
                }

                Drawing newD = AddImageToBody3(wordDocument, mainPart.GetIdOfPart(imagePart));

                Paragraph p = new Paragraph(new Run(newD));

                body.InsertAfter(p, body.FirstChild);
                mainPart.Document.Save();

                //foreach(var page in pageWithContent){

                //}
            }
        }

        private Drawing FormImage(WordprocessingDocument wordDoc, ImageOptions options = null)
        {
            if (options == null)
            {
                options = new ImageOptions();
            }

            MainDocumentPart mainPart = wordDoc.MainDocumentPart;
            ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Png);

            using (FileStream stream = new FileStream(@"./Files/test.png", FileMode.Open))
            {
                imagePart.FeedData(stream);
            }

            string relationshipId = mainPart.GetIdOfPart(imagePart);

            // new graphicFrameLocks
            var graphicFrameLocks = new A.GraphicFrameLocks()
            {
                NoChangeAspect = true
            };
            graphicFrameLocks.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            // new useLocalDpi
            var useLocalDpi = new UseLocalDpi() { Val = false };
            useLocalDpi.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            //new picture
            var picture = new PIC.Picture(
                new PIC.NonVisualPictureProperties(
                    new PIC.NonVisualDrawingProperties()
                    {
                        Id = (UInt32Value)1U,
                        Name = "image1.png"
                    },
                    new PIC.NonVisualPictureDrawingProperties()
                ),
                new PIC.BlipFill(
                    new A.Blip(
                        new A.BlipExtensionList(
                            new A.BlipExtension(
                                useLocalDpi
                            )
                            {
                                Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"
                            }
                        )
                    )
                    { Embed = relationshipId },
                    new A.Stretch(
                        new A.FillRectangle()
                    )
                ),
                new PIC.ShapeProperties(
                    new A.Transform2D(
                        new A.Offset() { X = 0L, Y = 0L },
                        new A.Extents() { Cx = 1080000L, Cy = 1080000L }
                    ),
                    new A.PresetGeometry(
                        new A.AdjustValueList()
                    ) { Preset = A.ShapeTypeValues.Rectangle }
                )
            );
            picture.AddNamespaceDeclaration("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture");

            // new graphic
            var graphic = new A.Graphic(
                new A.GraphicData(picture)
                {
                    Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"
                }
            );
            graphic.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            //new drawing
            return new Drawing(
                new DW.Anchor(
                    new DW.SimplePosition()
                    {
                        X = 0L,
                        Y = 0L
                    },
                    new DW.HorizontalPosition(
                        new DW.PositionOffset()
                        {
                            Text = "10"
                        })
                    {
                        RelativeFrom = DW.HorizontalRelativePositionValues.Character
                    },
                    new DW.VerticalPosition(
                        new DW.PositionOffset()
                        {
                            Text = "8890"
                        })
                    { RelativeFrom = DW.VerticalRelativePositionValues.Paragraph },
                    new DW.Extent()
                    {
                        Cx = options.Width * 360000L,
                        Cy = options.Height * 360000L
                    },
                    new DW.EffectExtent()
                    {
                        LeftEdge = 0L,
                        TopEdge = 0L,
                        RightEdge = 0L,
                        BottomEdge = 0L
                    },
                    new DW.WrapNone(),
                    new DW.DocProperties()
                    {
                        Id = (UInt32Value)1U,
                        Name = "Рисунок 1"
                    },
                    new DW.NonVisualGraphicFrameDrawingProperties(graphicFrameLocks),
                    graphic,
                    new RelativeWidth(
                        new PercentageWidth()
                        {
                            Text = "0"
                        }
                    )
                    {
                        ObjectId = SizeRelativeHorizontallyValues.Page
                    },
                    new RelativeHeight(
                        new PercentageWidth()
                        {
                            Text = "0"
                        }
                    )
                    {
                        RelativeFrom = SizeRelativeVerticallyValues.Page
                    }
                )
                {
                    DistanceFromTop = (UInt32Value)0U,
                    DistanceFromBottom = (UInt32Value)0U,
                    DistanceFromLeft = (UInt32Value)114300U,
                    DistanceFromRight = (UInt32Value)114300U,
                    SimplePos = false,
                    RelativeHeight = (UInt32Value)251658240U,
                    BehindDoc = true,
                    Locked = false,
                    LayoutInCell = true,
                    AllowOverlap = true,
                    EditId = "0B572E2B",
                    AnchorId = "4DC14A96"
                    //HorizontalPosition = new DW.HorizontalPosition()
                    //{
                    //    RelativeFrom = DW.HorizontalRelativePositionValues.Character
                    //}
                });
        }

        private Drawing AddImageToBody3(WordprocessingDocument wordDoc, string relationshipId, ImageOptions options = null)
        {
            if (options == null)
            {
                options = new ImageOptions();
            }

            DW.Anchor anchor1 = new DW.Anchor() { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)114300U, DistanceFromRight = (UInt32Value)114300U, SimplePos = false, RelativeHeight = (UInt32Value)251658240U, BehindDoc = true, Locked = false, LayoutInCell = true, AllowOverlap = true, EditId = "0B572E2B", AnchorId = "4DC14A96" };
            DW.SimplePosition simplePosition1 = new DW.SimplePosition() { X = 0L, Y = 0L };

            //DW.HorizontalPosition horizontalPosition1 = new DW.HorizontalPosition() { RelativeFrom = DW.HorizontalRelativePositionValues.Margin };
            //DW.HorizontalAlignment horizontalAlignment1 = new DW.HorizontalAlignment();
            //horizontalAlignment1.Text = "right";

            // Говорю что позицию хочу задавать в символах.
            DW.HorizontalPosition horizontalPosition1 = new DW.HorizontalPosition() { RelativeFrom = DW.HorizontalRelativePositionValues.Character };
            DW.PositionOffset horizontalAlignment1 = new DW.PositionOffset();
            // отступаю 10 символов
            horizontalAlignment1.Text = "10";

            horizontalPosition1.Append(horizontalAlignment1);

            DW.VerticalPosition verticalPosition1 = new DW.VerticalPosition() { RelativeFrom = DW.VerticalRelativePositionValues.Paragraph };
            DW.PositionOffset positionOffset1 = new DW.PositionOffset();
            positionOffset1.Text = "8890";

            verticalPosition1.Append(positionOffset1);
            //DW.Extent extent1 = new DW.Extent() { Cx = 5940425L, Cy = 4455160L };
            //DW.Extent extent1 = new DW.Extent() { Cx = 1000000L, Cy = 1000000L };
            DW.Extent extent1 = new DW.Extent() { Cx = options.Width * 360000L, Cy = options.Height * 360000L };
            DW.EffectExtent effectExtent1 = new DW.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L };
            DW.WrapNone wrapNone1 = new DW.WrapNone();
            DW.DocProperties docProperties1 = new DW.DocProperties() { Id = (UInt32Value)1U, Name = "Рисунок 1" };

            DW.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new DW.NonVisualGraphicFrameDrawingProperties();

            A.GraphicFrameLocks graphicFrameLocks1 = new A.GraphicFrameLocks() { NoChangeAspect = true };
            graphicFrameLocks1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            nonVisualGraphicFrameDrawingProperties1.Append(graphicFrameLocks1);

            A.Graphic graphic1 = new A.Graphic();
            graphic1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.GraphicData graphicData1 = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" };

            PIC.Picture picture1 = new PIC.Picture();
            picture1.AddNamespaceDeclaration("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture");

            PIC.NonVisualPictureProperties nonVisualPictureProperties1 = new PIC.NonVisualPictureProperties();
            PIC.NonVisualDrawingProperties nonVisualDrawingProperties1 = new PIC.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "image1.png" };
            PIC.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties1 = new PIC.NonVisualPictureDrawingProperties();

            nonVisualPictureProperties1.Append(nonVisualDrawingProperties1);
            nonVisualPictureProperties1.Append(nonVisualPictureDrawingProperties1);

            PIC.BlipFill blipFill1 = new PIC.BlipFill();

            A.Blip blip1 = new A.Blip() { Embed = relationshipId };

            A.BlipExtensionList blipExtensionList1 = new A.BlipExtensionList();

            A.BlipExtension blipExtension1 = new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };

            UseLocalDpi useLocalDpi1 = new UseLocalDpi() { Val = false };
            useLocalDpi1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            blipExtension1.Append(useLocalDpi1);

            blipExtensionList1.Append(blipExtension1);

            blip1.Append(blipExtensionList1);

            A.Stretch stretch1 = new A.Stretch();
            A.FillRectangle fillRectangle1 = new A.FillRectangle();

            stretch1.Append(fillRectangle1);

            blipFill1.Append(blip1);
            blipFill1.Append(stretch1);

            PIC.ShapeProperties shapeProperties1 = new PIC.ShapeProperties();

            A.Transform2D transform2D1 = new A.Transform2D();
            A.Offset offset1 = new A.Offset() { X = 0L, Y = 0L };
            //A.Extents extents1 = new A.Extents() { Cx = 5940425L, Cy = 4455160L };
            A.Extents extents1 = new A.Extents() { Cx = 1080000L, Cy = 1080000L };

            transform2D1.Append(offset1);
            transform2D1.Append(extents1);

            A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);

            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);

            picture1.Append(nonVisualPictureProperties1);
            picture1.Append(blipFill1);
            picture1.Append(shapeProperties1);

            graphicData1.Append(picture1);

            graphic1.Append(graphicData1);

            RelativeWidth relativeWidth1 = new RelativeWidth() { ObjectId = SizeRelativeHorizontallyValues.Page };
            PercentageWidth percentageWidth1 = new PercentageWidth();
            percentageWidth1.Text = "0";

            relativeWidth1.Append(percentageWidth1);

            RelativeHeight relativeHeight1 = new RelativeHeight() { RelativeFrom = SizeRelativeVerticallyValues.Page };
            PercentageHeight percentageHeight1 = new PercentageHeight();
            percentageHeight1.Text = "0";

            relativeHeight1.Append(percentageHeight1);

            anchor1.Append(simplePosition1);
            anchor1.Append(horizontalPosition1);
            anchor1.Append(verticalPosition1);
            anchor1.Append(extent1);
            anchor1.Append(effectExtent1);
            anchor1.Append(wrapNone1);
            anchor1.Append(docProperties1);
            anchor1.Append(nonVisualGraphicFrameDrawingProperties1);
            anchor1.Append(graphic1);
            anchor1.Append(relativeWidth1);
            anchor1.Append(relativeHeight1);

            IEnumerable<SectionProperties> sections = wordDoc.MainDocumentPart.Document.Body.Elements<SectionProperties>();

            return new Drawing(anchor1);

            //return new Paragraph(new Run(new Drawing(anchor1)));

            //wordDoc.MainDocumentPart.Document.Body.InsertAfter(new Paragraph(new Run(new Drawing(anchor1))), wordDoc.MainDocumentPart.Document.Body.FirstChild);

            //wordDoc.MainDocumentPart.Document.Save();
        }

        private void AddImageToBody2(WordprocessingDocument wordDoc, string relationshipId)
        {
            DW.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new DW.NonVisualGraphicFrameDrawingProperties();

            A.GraphicFrameLocks graphicFrameLocks1 = new A.GraphicFrameLocks() { NoChangeAspect = true };
            graphicFrameLocks1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            nonVisualGraphicFrameDrawingProperties1.Append(graphicFrameLocks1);


            var element = new Drawing(new DW.Anchor(
                new DW.SimplePosition() { X = 0L, Y = 0L },
                new DW.Extent() { Cx = 990000L, Cy = 792000L },
                new DW.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                new DW.WrapNone(),
                new DW.DocProperties() { Id = (UInt32Value)1U, Name = "image1" },
                nonVisualGraphicFrameDrawingProperties1,
                  new A.Graphic(
                             new A.GraphicData(
                                 new PIC.Picture(
                                     new PIC.NonVisualPictureProperties(
                                         new PIC.NonVisualPictureProperties(),
                                         new PIC.NonVisualDrawingProperties()
                                         {
                                             Id = (UInt32Value)0U,
                                             Name = "image1.png"
                                         },
                                         new PIC.NonVisualPictureDrawingProperties()),
                                     new PIC.BlipFill(
                                         new A.Blip(
                                             new A.BlipExtensionList(
                                                 new A.BlipExtension()
                                                 {
                                                     Uri =
                                                        "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                 })
                                         )
                                         {
                                             Embed = relationshipId,
                                             CompressionState =
                                             A.BlipCompressionValues.Print
                                         },
                                         new A.Stretch(
                                             new A.FillRectangle())),
                                     new PIC.ShapeProperties(
                                         new A.Transform2D(
                                             new A.Offset() { X = 0L, Y = 0L },
                                             new A.Extents() { Cx = 990000L, Cy = 792000L }),
                                         new A.PresetGeometry(
                                             new A.AdjustValueList()
                                         )
                                         { Preset = A.ShapeTypeValues.Rectangle }))
                             )
                             { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
            )
            {
                DistanceFromTop = (UInt32Value)0U,
                DistanceFromBottom = (UInt32Value)0U,
                DistanceFromLeft = (UInt32Value)0U,
                DistanceFromRight = (UInt32Value)0U,
                SimplePos = false,
                RelativeHeight = (UInt32Value)251658240U,
                BehindDoc = true,
                Locked = false,
                LayoutInCell = true,
                AllowOverlap = true
            });
            wordDoc.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(element)));


            wordDoc.MainDocumentPart.Document.Save();
        }

        private void AddImageToBody(WordprocessingDocument wordDoc, string relationshipId)
        {
            // Define the reference of the image.
            var element =
                 new Drawing(
                     //new DW.Anchor(
                     new DW.Inline(
                         new DW.WrapTight() { WrapText = DW.WrapTextValues.BothSides },
                         new DW.Extent() { Cx = 990000L, Cy = 792000L },
                         new DW.EffectExtent()
                         {
                             LeftEdge = 0L,
                             TopEdge = 0L,
                             RightEdge = 0L,
                             BottomEdge = 0L
                         },
                         new DW.DocProperties()
                         {
                             Id = (UInt32Value)1U,
                             Name = "image1"
                         },
                         new DW.NonVisualGraphicFrameDrawingProperties(
                             new A.GraphicFrameLocks() { NoChangeAspect = true, }),
                         new A.Graphic(
                             new A.GraphicData(
                                 new PIC.Picture(
                                     new PIC.NonVisualPictureProperties(
                                         new PIC.NonVisualDrawingProperties()
                                         {
                                             Id = (UInt32Value)0U,
                                             Name = "image1.png"
                                         },
                                         new PIC.NonVisualPictureDrawingProperties()),
                                     new PIC.BlipFill(
                                         new A.Blip(
                                             new A.BlipExtensionList(
                                                 new A.BlipExtension()
                                                 {
                                                     Uri =
                                                        "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                 })
                                         )
                                         {
                                             Embed = relationshipId,
                                             CompressionState =
                                             A.BlipCompressionValues.Print
                                         },
                                         new A.Stretch(
                                             new A.FillRectangle())),
                                     new PIC.ShapeProperties(
                                         new A.Transform2D(
                                             new A.Offset() { X = 0L, Y = 0L },
                                             new A.Extents() { Cx = 990000L, Cy = 792000L }),
                                         new A.PresetGeometry(
                                             new A.AdjustValueList()
                                         )
                                         { Preset = A.ShapeTypeValues.Rectangle }))
                             )
                             { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                     )
                     {
                         DistanceFromTop = (UInt32Value)0U,
                         DistanceFromBottom = (UInt32Value)0U,
                         DistanceFromLeft = (UInt32Value)0U,
                         DistanceFromRight = (UInt32Value)0U,
                         EditId = "50D07946"
                     });

            // Append the reference to body, the element should be in a Run.
            wordDoc.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(element)));
        }

        public void OpenAndAddTextToWordDocument(string filepath, string txt)
        {
            // Open a WordprocessingDocument for editing using the filepath.
            WordprocessingDocument wordprocessingDocument =
                WordprocessingDocument.Open(filepath, true);

            // Assign a reference to the existing document body.
            Body body = wordprocessingDocument.MainDocumentPart.Document.Body;

            // Add new text.
            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());
            run.AppendChild(new Text(txt));

            // Close the handle explicitly.
            wordprocessingDocument.Close();
        }

        public void AddPageNumber(string filepath)
        {
            // Open a WordprocessingDocument for editing using the filepath.
            WordprocessingDocument package =
                WordprocessingDocument.Open(filepath, true);

            var mainPart = package.MainDocumentPart;

            mainPart.DeleteParts(mainPart.FooterParts);

            FooterPart footerPart = mainPart.AddNewPart<FooterPart>();

            string footerPartId = mainPart.GetIdOfPart(footerPart);

            GenerateFooterPartContentFirst(footerPart);

            IEnumerable<SectionProperties> sections = mainPart.Document.Body.Elements<SectionProperties>();

            foreach (var section in sections)
            {
                // Delete existing references to headers and footers
                section.RemoveAllChildren<FooterReference>();
                section.RemoveAllChildren<TitlePage>();

                // Create the new header and footer reference node
                section.PrependChild<FooterReference>(new FooterReference() { Id = footerPartId, Type = HeaderFooterValues.First });
                section.PrependChild<FooterReference>(new FooterReference() { Id = footerPartId, Type = HeaderFooterValues.Default });
                section.PrependChild<TitlePage>(new TitlePage());
            }

            mainPart.Document.Save();

            package.Close();
        }

        private void GenerateFooterPartContentFirst(FooterPart part)
        {
            Footer footer =
              new Footer(
               new Paragraph(
                new ParagraphProperties(
                 //new ParagraphStyleId() { Val = "Footer" },
                 new Justification() { Val = JustificationValues.Center }),
                new Run(
                 new Text() { Text = "Сторiнка ", Space = SpaceProcessingModeValues.Preserve }),
                new Run(
                 new SimpleField() { Instruction = "Page" }),
                new Run(
                 new Text() { Text = " з ", Space = SpaceProcessingModeValues.Preserve }),
                new Run(
                 new SimpleField() { Instruction = "NUMPAGES" })
               ));

            part.Footer = footer;
        }
    }
}
