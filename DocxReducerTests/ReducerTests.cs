using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using DocxReducer;

namespace DocxReducerTests
{
    [TestClass]
    public class ReducerTests
    {
        [TestMethod]
        public void TestReduce()
        {
            var reducer = new Reducer();

            var pathToFile = Path.Combine(@"..\..\..\..\", "TestFiles", "Holocaust.docx");

            var reducedDocx = reducer.Reduce(pathToFile);

            Assert.IsNotNull(reducedDocx);
        }
    }
}
