using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using DocumentFormat.OpenXml.Wordprocessing;

using DocxReducer.Core;

namespace DocxReducerTests.Core
{
    [TestClass]
    public class RunStyleCreatorTests
    {
        private TestDataGenerator DataGenerator { get; }

        private Styles DocStyles { get; set; }

        private RunStylesManager StyleCreator { get; set; }

        public RunStyleCreatorTests()
        {
            DataGenerator = new TestDataGenerator();
        }

        [TestInitialize]
        public void TestInit()
        {
            DocStyles = new Styles();

            StyleCreator = new RunStylesManager(DocStyles);
        }

        //[TestMethod]
        //public void TestCreateGlobalRunStyle()
        //{
        //    var styleCreator = new RunStyleCreator(

        //    var rPrHash = 123;
        //    Styles.Append(new Style()
        //    {
        //        StyleId = Processor.GenerateNewStyleId(rPrHash)
        //    });

        //    var rPr = DataGenerator.GenerateRunProperties();

        //    Processor.CreateGlobalRunStyle(rPr, rPrHash);

        //    var styles = Styles.Descendants<Style>().ToList();

        //    Assert.IsTrue(styles.Count() == 2);
        //    Assert.AreNotEqual(styles[0].StyleId.Value, styles[1].StyleId.Value);
        //}

        private Style CreateRunStyle(Run run)
        {
            return StyleCreator.CreateStyleForRun(run.RunProperties, StyleCreator.GenerateHash(run.RunProperties));
        }

        private bool IsStyleReplacementWorthIt(Run run)
        {
            return StyleCreator.IsStyleReplacementWorthIt(run.RunProperties, CreateRunStyle(run));
        }

        [TestMethod]
        public void TestIsStyleReplacementWorthIt()
        {
            var rPr = new RunProperties(new NoProof());
            var run = DataGenerator.GenerateRun(rPr, " ");
            Assert.IsFalse(IsStyleReplacementWorthIt(run));

            var run2 = DataGenerator.GenerateRun("TEXT");
            Assert.IsTrue(IsStyleReplacementWorthIt(run2));

            var run3 = new Run();
            run3.RunProperties = new RunProperties();
            Assert.IsFalse(IsStyleReplacementWorthIt(run3));
        }

        [TestMethod]
        public void TestGenerateStyleId()
        {
            int propertiesHash = 1234;
            int secondPropertiesHash = int.Parse("12" + propertiesHash.ToString());
            var styleId = StyleCreator.GenerateStyleId(propertiesHash);

            DocStyles.Append(new Style()
            {
                StyleId = styleId
            });

            Assert.AreNotEqual(styleId, StyleCreator.GenerateStyleId(secondPropertiesHash));
        }
    }
}
