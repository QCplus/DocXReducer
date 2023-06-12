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

        [TestMethod]
        public void ToBase_ConvertTo36Base()
            => Assert.AreEqual("9IX", 12345.ToBase(36));

        [TestMethod]
        public void ToBase_ConvertZeroTo36Base()
            => Assert.AreEqual("0", 0.ToBase(36));

        [TestMethod]
        public void ToBase_ConvertNegativeNumber()
            => Assert.AreEqual("-9IX", (-12345).ToBase(36));

        [TestMethod]
        public void TakeFirstBits_IfAllBitsAreOnes()
            => Assert.AreEqual(0b1111111111, 0b111111111111111.TakeFirstBits(10));

        [TestMethod]
        public void TakeFirstBits()
            => Assert.AreEqual(0b10000111, 0b1010000111.TakeFirstBits(8));

        [TestMethod]
        public void TakeFirstBits_PassMoreThan32Bits()
            => Assert.AreEqual(0b101, 0b101.TakeFirstBits(100));
    }
}
