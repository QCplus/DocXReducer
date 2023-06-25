using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace DocxReducerTests
{
    internal static class AssertExtensions
    {
        public static void AllRunStylesDefined(this Assert assert, WordprocessingDocument doc)
        {
            var runProperties = doc.MainDocumentPart.RootElement.Descendants<RunProperties>();
            var styles = doc.MainDocumentPart.StyleDefinitionsPart.Styles;

            foreach (var rPr in runProperties)
            {
                string runStyleId = rPr.RunStyle.Val;

                Assert.IsTrue(
                    styles.Elements<Style>().Where(s => s.StyleId == runStyleId).Any(),
                    $"Paragraph has undefined style with id {runStyleId}");
            }
        }

        public static void HaveCustomStyles(this Assert assert, Styles styles, int expectedStylesCount)
        {
            Assert.AreEqual(expectedStylesCount,
                styles.Elements<Style>().Where(s => s.Default == false).Count());
        }
    }
}
