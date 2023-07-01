using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxReducer.Core;

namespace DocxReducerTests.CoreTests
{
    [TestClass]
    public class StylesManagerTests
    {
        private TestDataGenerator DataGenerator { get; }

        private Styles DocStyles { get; set; }

        private StylesManager StylesManager { get; set; }

        public StylesManagerTests()
        {
            DataGenerator = new TestDataGenerator();
        }

        [TestInitialize]
        public void TestInit()
        {
            DocStyles = new Styles();

            StylesManager = new StylesManager(DocStyles);
        }

        [TestMethod]
        public void CreateStyle_CreatesNewStyleInDocument()
        {
            var runProperties = new RunProperties(@"
                <w:rPr xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                    <w:color w:val=""2F5496"" w:themeColor=""accent1"" w:themeShade=""BF"" />
                    <w:sz w:val=""32"" />
                    <w:szCs w:val=""32"" />
                </w:rPr>");

            string styleId = StylesManager.CreateStyle(runProperties).StyleId;

            Assert.IsTrue(
                DocStyles.Descendants<Style>().Where(s => s.StyleId == styleId).Any());
        }
    }
}
