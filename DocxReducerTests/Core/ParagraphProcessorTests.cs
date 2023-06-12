using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxReducer.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OMath = DocumentFormat.OpenXml.Math.OfficeMath;

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
            ParProcessor = new ParagraphProcessor(DocStyles,
                new DocxReducer.Options.ReducerOptions()
                {
                    CreateNewStyles = true,
                    DeleteBookmarks = true
                });
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
        public void Process_TabCharBetweenTwoRuns()
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
            Assert.AreEqual(4, run.ChildElements.Count);
            Assert.AreEqual("TEXT", run.ElementAt(1).InnerText);
            Assert.AreEqual(typeof(TabChar), run.ElementAt(2).GetType());
            Assert.AreEqual("1", run.ElementAt(3).InnerText);
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
            Assert.AreEqual(4, run.ChildElements.Count);
            Assert.AreEqual("TEXT", run.ElementAt(1).InnerText);
            Assert.AreEqual(typeof(TabChar), run.ElementAt(2).GetType());
            Assert.AreEqual("1", run.ElementAt(3).InnerText);
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

        private Paragraph Process(string xml)
        {
            var par = new Paragraph(xml);

            ParProcessor.Process(par);

            return par;
        }

        [TestMethod]
        public void Process_MathFormulaKeepPositionBetweenRuns()
        {
            var par = Process(@"
                <w:p w:rsidR=""008D5083"" xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                  <w:r>
                    <w:t xml:space=""preserve"">Complexity is </w:t>
                  </w:r>
                  <m:oMath xmlns:m=""http://schemas.openxmlformats.org/officeDocument/2006/math"">
                    <m:r>
                      <m:t>O</m:t>
                    </m:r>
                  </m:oMath>
                  <w:r>
                    <w:t xml:space=""preserve""> complexity</w:t>
                  </w:r>
                </w:p>");

            Assert.AreEqual(3, par.ChildElements.Count);

            Assert.AreEqual(typeof(OMath), par.ElementAt(1).GetType());
        }

        [TestMethod]
        public void Process_AllBookmarksWereRemoved()
        {
            var par = Process(@"
                <w:p w14:paraId=""0334E79D"" w14:textId=""1C39F1F7"" xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"" xmlns:w14=""http://schemas.microsoft.com/office/word/2010/wordml"">
                  <w:bookmarkStart w:name=""_Toc517125172"" w:id=""13"" />
                  <w:bookmarkStart w:name=""_Toc517125298"" w:id=""14"" />
                  <w:r>
                    <w:t>TEXT</w:t>
                  </w:r>
                  <w:bookmarkEnd w:id=""13"" />
                  <w:bookmarkEnd w:id=""14"" />
                </w:p>");

            Assert.AreEqual(0, par.Descendants<BookmarkStart>().Count());
            Assert.AreEqual(0, par.Descendants<BookmarkEnd>().Count());
        }

        [TestMethod]
        public void Process_AllRsidsWereRemoved()
        {
            var par = Process(@"
                <w:p w:rsidRPr=""00D20A94"" w:rsidR=""00D20A94"" xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                  <w:r w:rsidRPr=""00D20A94"">
                    <w:t>TEXT</w:t>
                  </w:r>
                  <w:r w:rsidRPr=""00D20A94"">
                    <w:rPr>
                      <w:color w:val=""2B91AF"" />
                      <w:sz w:val=""19"" />
                    </w:rPr>
                    <w:t xml:space=""preserve""> TEXT</w:t>
                  </w:r>
                </w:p>");

            Assert.AreEqual(0, par.GetAttributes().Count());
            Assert.AreEqual(0, par.Descendants<Run>().Where(r => r.HasAttributes).Count());
        }
    }
}
