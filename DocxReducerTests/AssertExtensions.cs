using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace DocxReducerTests
{
    internal static class AssertExtensions
    {
        public static void AllRunStylesDefined(this Assert assert, Paragraph paragraph, IEnumerable<string> styleIds)
        {
            foreach (var rPr in paragraph.Descendants<RunProperties>())
            {
                string runStyleId = rPr.RunStyle.Val;

                if (!styleIds.Contains(runStyleId))
                    Assert.Fail($"Paragraph has undefined style with id {runStyleId}");
            }
        }
    }
}
