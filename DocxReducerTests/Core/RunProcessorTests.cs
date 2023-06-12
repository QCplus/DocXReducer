using System;
using System.Linq;
using System.Collections.Generic;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using DocumentFormat.OpenXml.Wordprocessing;

using DocxReducer.Core;

namespace DocxReducerTests.Core
{
    [TestClass]
    public class RunProcessorTests
    {
        private TestDataGenerator DataGenerator { get; }

        private RunProcessor Processor { get; set; }

        private Styles Styles { get; set; }

        public RunProcessorTests()
        {
            DataGenerator = new TestDataGenerator();
        }

        [TestInitialize]
        public void TestInit()
        {
            Styles = new Styles();

            Processor = new RunProcessor(new RunStylesManager(Styles));
        }

        [TestMethod]
        public void TestRunPropertiesAreEqual()
        {
            Assert.IsTrue(Processor.AreEqual(
                DataGenerator.GenerateRunProperties(),
                DataGenerator.GenerateRunProperties()
                ));

            Assert.IsFalse(Processor.AreEqual(
                DataGenerator.GenerateRunPropertiesBold(),
                DataGenerator.GenerateRunProperties()
                ));

            Assert.IsTrue(Processor.AreEqual(
                null,
                null
                ));

            Assert.IsFalse(Processor.AreEqual(
                DataGenerator.GenerateRunProperties(),
                null
                ));
        }

        [TestMethod]
        public void TestReplaceRunPropertiesWithStyle()
        {
            var run = DataGenerator.GenerateRun("TEXT");

            Processor.ReplaceRunPropertiesWithStyle(run);

            Assert.IsNotNull(run.RunProperties.RunStyle);
            Assert.IsTrue(run.RunProperties.ChildElements.Count == 1);
        }

        [TestMethod]
        public void Merge_SecondTextWithSpace()
        {
            var paragraph = new Paragraph(@"
                <w:p w:rsidP=""009B6E00"" xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                    <w:r w:rsidP=""009B6E3A"">
                        <w:t>TEST</w:t>
                    </w:r>
                    <w:r w:rsidP=""009B6E00"">
                        <w:t xml:space=""preserve""> text</w:t>
                    </w:r>
                </w:p>");
            var runs = paragraph.Elements<Run>().ToList();

            var resultedRun = Processor.MergeIfNeeded(runs[0], runs[1]);

            Assert.AreEqual(1, resultedRun.Elements<Text>().Count());
            Assert.AreEqual("TEST text", resultedRun.InnerText);
            Assert.AreEqual(1, paragraph.Elements<Run>().Count());
        }

        [TestMethod]
        public void Merge_IfFirstRunWithoutText()
        {
            var paragraph = new Paragraph(@"
                <w:p w:rsidP=""009B6E00"" xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                    <w:r w:rsidP=""009B6E3A"">
                    </w:r>
                    <w:r w:rsidP=""009B6E00"">
                        <w:t>text</w:t>
                    </w:r>
                </w:p>");
            var runs = paragraph.Elements<Run>().ToList();

            var resultedRun = Processor.MergeIfNeeded(runs[0], runs[1]);

            Assert.AreEqual(1, resultedRun.Elements<Text>().Count());
            Assert.AreEqual("text", resultedRun.InnerText);
        }
    }
}
