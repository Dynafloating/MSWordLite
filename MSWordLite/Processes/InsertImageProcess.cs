using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using MSWordLite.Orders;
using System;
using System.IO;
using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace MSWordLite.Processes
{
    class InsertImageProcess : OrderProcess<InsertImageOrder>
    {
        long widthInEMU { get; set; }
        long heightInEMU { get; set; }

        const decimal INCH_TO_CM = 2.54M;
        const decimal CM_TO_EMU = 360000M;
        const int DPI = 300;

        public override OrderResult Initialize(Document document)
        {
            Initializer.Bookmarks(document);
            widthInEMU = Convert.ToInt64(Order.Width * INCH_TO_CM * CM_TO_EMU / DPI);
            heightInEMU = Convert.ToInt64(Order.Height * INCH_TO_CM * CM_TO_EMU / DPI);

            return new OrderResult(success: true);
        }

        public override OrderResult Process(Document document)
        {
            if (document.WordBookmarks.ContainsKey(Order.BookmarkKey))
            {
                var contentType = _getContentType(Order.ContentType);
                using (var stream  = new MemoryStream(Order.ImageContent))
                {
                    var run = GenerateImageRun(document.WordDocument, stream, contentType, widthInEMU, heightInEMU);
                    document.WordBookmarks[Order.BookmarkKey].InsertRun(run);
                }
            }

            return new OrderResult(success: true);
        }

        static ImagePartType _getContentType(string contentType)
        {
            var adaptedContentType = contentType.Split('/').Last().ToLower();
            return Enum.GetValues(typeof(ImagePartType))
                .Cast<ImagePartType>()
                .Where(o => o.ToString().ToLower() == adaptedContentType)
                .FirstOrDefault();
        }

        static Run GenerateImageRun(
            WordprocessingDocument wordDoc, Stream content, ImagePartType contentType, long widthInEMU, long heightInEMU
            )
        {
            var mainPart = wordDoc.MainDocumentPart;
            var imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
            var relationshipId = mainPart.GetIdOfPart(imagePart);
            imagePart.FeedData(content);

            var imageId = Guid.NewGuid().ToString();
            var imageName = $"IMG_{imageId}";
            var fileName = $"{imageId}.{contentType.ToString()}";

            // Define the reference of the image.
            var element =
                 new Drawing(
                     new DW.Inline(
                         //Size of image, unit = EMU(English Metric Unit)
                         //1 cm = 360000 EMUs
                         new DW.Extent() { Cx = widthInEMU, Cy = heightInEMU },
                         new DW.EffectExtent()
                         {
                             LeftEdge = 0L,
                             TopEdge = 0L,
                             RightEdge = 0L,
                             BottomEdge = 0L
                         },
                         new DW.DocProperties()
                         {
                             Id = 1U,
                             Name = imageName
                         },
                         new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks() { NoChangeAspect = true }),
                         new A.Graphic(
                             new A.GraphicData(
                                 new PIC.Picture(
                                     new PIC.NonVisualPictureProperties(
                                         new PIC.NonVisualDrawingProperties()
                                         {
                                             Id = 0U,
                                             Name = fileName
                                         },
                                         new PIC.NonVisualPictureDrawingProperties()),
                                     new PIC.BlipFill(
                                         new A.Blip(
                                             new A.BlipExtensionList(
                                                 new A.BlipExtension()
                                                 {
                                                     Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                 })
                                         )
                                         {
                                             Embed = relationshipId,
                                             CompressionState =
                                             A.BlipCompressionValues.Print
                                         },
                                         new A.Stretch(new A.FillRectangle())),
                                     new PIC.ShapeProperties(
                                         new A.Transform2D(
                                             new A.Offset() { X = 0L, Y = 0L },
                                             new A.Extents()
                                             {
                                                 Cx = widthInEMU,
                                                 Cy = heightInEMU
                                             }),
                                         new A.PresetGeometry(new A.AdjustValueList())
                                         {
                                             Preset = A.ShapeTypeValues.Rectangle
                                         }))
                             )
                             { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                     )
                     {
                         DistanceFromTop = 0U,
                         DistanceFromBottom = 0U,
                         DistanceFromLeft = 0U,
                         DistanceFromRight = 0U,
                         EditId = "50D07946"
                     });

            return new Run(element);
        }
    }
}
