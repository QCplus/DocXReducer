using System;
using DocumentFormat.OpenXml.Packaging;

using DocxReducer;

namespace ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0)
                throw new Exception("Specify an input file path (.docx)");
            if (args.Length > 2)
                Console.WriteLine("Unknown args was specified. Ignoring...");

            var inputFilePath = args[0];
            var outputFilePath = args.Length > 1 ? args[1] : "out.docx";

            var docx = WordprocessingDocument.Open(inputFilePath, true);

            var reducer = new Reducer();
            reducer.Reduce(docx);

            docx.SaveAs(outputFilePath);
        }
    }
}
