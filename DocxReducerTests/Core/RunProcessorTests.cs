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
        public void TestReplaceRunPropertiesWithStyle()
        {
            var run = DataGenerator.GenerateRun("TEXT");

            Processor.ReplaceRunPropertiesWithStyle(run);

            Assert.IsNotNull(run.RunProperties.RunStyle);
            Assert.IsTrue(run.RunProperties.ChildElements.Count == 1);
        }
    }
}
