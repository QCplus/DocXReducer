using Microsoft.VisualStudio.TestTools.UnitTesting;
using DocxReducer.Helpers;

namespace DocxReducerTests.Helpers
{
    [TestClass]
    public class IntExtensionsTests
    {
        [TestMethod]
        public void TestGetLastNDigits()
        {
            Assert.AreEqual("456", 123456.GetLastNDigits(3));
            Assert.AreEqual("0", 0.GetLastNDigits(3));
            Assert.AreEqual("76", (-9876).GetLastNDigits(2));
            Assert.AreEqual("", 9876.GetLastNDigits(0));
            Assert.AreEqual("", 9876.GetLastNDigits(-2));
            Assert.AreEqual("3647", int.MaxValue.GetLastNDigits(4));
            Assert.AreEqual("83", 83.GetLastNDigits(4));
        }
    }
}
