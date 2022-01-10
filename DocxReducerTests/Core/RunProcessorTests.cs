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

            Processor = new RunProcessor(Styles);
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
        public void TestCreateGlobalRunStyle()
        {
            var rPrHash = 123;
            Styles.Append(new Style()
            {
                StyleId = Processor.GenerateNewStyleId(rPrHash)
            });

            var rPr = DataGenerator.GenerateRunProperties();

            Processor.CreateGlobalRunStyle(rPr, rPrHash);

            var styles = Styles.Descendants<Style>().ToList();

            Assert.IsTrue(styles.Count() == 2);
            Assert.AreNotEqual(styles[0].StyleId.Value, styles[1].StyleId.Value);
        }

        [TestMethod]
        public void TestReplaceRunPropertiesWithStyle()
        {
            var run = DataGenerator.GenerateRun("TEXT");

            Processor.ReplaceRunPropertiesWithStyleIfNeeded(run);

            Assert.IsNotNull(run.RunProperties.RunStyle);
            Assert.IsTrue(run.RunProperties.ChildElements.Count == 1);
        }

        [TestMethod]
        public void TestIsStyleCreationWorthIt()
        {
            var rPr = new RunProperties(new NoProof());
            var run = DataGenerator.GenerateRun(rPr, " ");
            Assert.IsFalse(Processor.IsStyleCreationWorthIt(run));

            Assert.IsTrue(Processor.IsStyleCreationWorthIt(
                DataGenerator.GenerateRun("TEXT")
                ));

            Assert.IsFalse(Processor.IsStyleCreationWorthIt(new Run()));
        }
    }
}
