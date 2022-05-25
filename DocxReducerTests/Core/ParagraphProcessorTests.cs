using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

using DocxReducer.Core;

namespace DocxReducerTests.Core
{
    [TestClass]
    public class ParagraphProcessorTests
    {
        private TestDataGenerator DataGenerator { get; }

        public ParagraphProcessorTests()
        {
            DataGenerator = new TestDataGenerator();
        }

        private Run[] GenerateRuns()
        {
            var runs = new Run[]
            {
                DataGenerator.GenerateRun("THIS IS"),
                DataGenerator.GenerateRun(" A"),
                DataGenerator.GenerateRun(" "),
                DataGenerator.GenerateRun(DataGenerator.GenerateRunPropertiesBold(), "TEST"),
                DataGenerator.GenerateRun(" DOCUMENT.")
            };

            return runs;
        }

        private Paragraph GenerateParagraph()
        {
            Paragraph paragraph1 = new Paragraph();
            paragraph1.Append(DataGenerator.GenerateParagraphProperties());

            foreach (var r in GenerateRuns())
                paragraph1.Append(r);

            return paragraph1;
        }

        private bool IsTextPreserved(Run run)
        {
            var text = run.GetFirstChild<Text>();

            return text.Space != null && text.Space.Value == SpaceProcessingModeValues.Preserve;
        }

        [TestMethod]
        public void ProcessTest()
        {
            var processor = new ParagraphProcessor(new Styles(), true);

            var par = GenerateParagraph();

            processor.Process(par);

            var runs = par.Elements<Run>().ToList();
            Assert.AreEqual(3, runs.Count);

            Assert.AreEqual("THIS IS A ", runs[0].InnerText);
            Assert.IsTrue(IsTextPreserved(runs[0]));
            Assert.AreEqual("TEST", runs[1].InnerText);
            Assert.IsTrue(IsTextPreserved(runs[2]));
            Assert.AreEqual(" DOCUMENT.", runs[2].InnerText);
        }

        [TestMethod]
        public void ProcessWithoutRunsTest()
        {
            var processor = new ParagraphProcessor(new Styles(), true);

            var par = new Paragraph();
            par.Append(DataGenerator.GenerateParagraphProperties());

            processor.Process(par);
        }

        [TestMethod]
        public void ProcessRunsWithTabBetween()
        {
            var processor = new ParagraphProcessor(new Styles(), false);

            var par = new Paragraph();
            par.Append(
                DataGenerator.GenerateRun("Text"),
                DataGenerator.GenerateRun(new TabChar(), new Text() { Text = "t" }),
                DataGenerator.GenerateRun("1"));

            processor.Process(par);

            Assert.AreEqual(3, par.Elements<Run>().Count());
        }
    }
}
