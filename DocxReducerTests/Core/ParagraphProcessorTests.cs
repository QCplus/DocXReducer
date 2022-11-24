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

        private Styles DocStyles { get; set; }

        private ParagraphProcessor ParProcessor { get; set; }

        public ParagraphProcessorTests()
        {
            DataGenerator = new TestDataGenerator();
        }

        [TestInitialize]
        public void TestInit()
        {
            DocStyles = new Styles();
            ParProcessor = new ParagraphProcessor(DocStyles, true);
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
            Paragraph paragraph1 = new Paragraph(
                DataGenerator.GenerateParagraphProperties());

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
            var par = GenerateParagraph();

            ParProcessor.Process(par);

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
            var par = new Paragraph();
            par.Append(DataGenerator.GenerateParagraphProperties());

            ParProcessor.Process(par);
        }

        [TestMethod]
        public void ProcessRunsWithTabBetween()
        {
            var par = new Paragraph(
                DataGenerator.GenerateRun("TEXT"),
                DataGenerator.GenerateRun(new TabChar()),
                DataGenerator.GenerateRun("1"));

            ParProcessor.Process(par);

            /* Should be
             *  <w:r>
			 *      <w:rPr>
			 *          ...
			 *      </w:rPr>
			 *      <w:t>TEXT</w:t>
			 *      <w:tab/>
			 *      <w:t>1</w:t>
			 *  </w:r>
             */

            Assert.AreEqual(1, par.ChildElements.Count);

            var run = par.GetFirstChild<Run>();
            Assert.AreEqual(4,                  run.ChildElements.Count);
            Assert.AreEqual("TEXT",             run.ElementAt(1).InnerText);
            Assert.AreEqual(typeof(TabChar),    run.ElementAt(2).GetType());
            Assert.AreEqual("1",                run.ElementAt(3).InnerText);
        }

        [TestMethod]
        public void ProcessRunsWithTabText()
        {
            var par = new Paragraph(
                DataGenerator.GenerateRun("TEXT"),
                DataGenerator.GenerateRun(
                    new TabChar(),
                    new Text("1"))
                );

            ParProcessor.Process(par);

            /*
             *  <w:r>
             *      <w:rPr>
             *          ...
             *      </w:rPr>
             *      <w:t>TEXT</w:t>
             *      <w:tab/>
             *      <w:t>1</w:t>
             *  </w:r>
             */

            Assert.AreEqual(1, par.ChildElements.Count);

            var run = par.GetFirstChild<Run>();
            Assert.AreEqual(4,                  run.ChildElements.Count);
            Assert.AreEqual("TEXT",             run.ElementAt(1).InnerText);
            Assert.AreEqual(typeof(TabChar),    run.ElementAt(2).GetType());
            Assert.AreEqual("1",                run.ElementAt(3).InnerText);
        }

        [TestMethod]
        public void FirstRunWithoutText()
        {
            var par = new Paragraph(
                DataGenerator.GenerateRun(),
                DataGenerator.GenerateRun("TEXT"));

            ParProcessor.Process(par);

            Assert.AreEqual(1, par.ChildElements.Count);
            Assert.AreEqual("TEXT", par.GetFirstChild<Run>().InnerText);
        }

        [TestMethod]
        public void SecondRunWithoutText()
        {
            var par = new Paragraph(
                DataGenerator.GenerateRun("TEXT"),
                DataGenerator.GenerateRun());

            ParProcessor.Process(par);

            Assert.AreEqual(1, par.ChildElements.Count);
            Assert.AreEqual("TEXT", par.GetFirstChild<Run>().InnerText);
        }
    }
}
