using DocumentFormat.OpenXml.Packaging;
using DocxReducer.DI;
using DocxReducer.Options;
using DocxReducer.Processors;
using Microsoft.Extensions.DependencyInjection;

namespace DocxReducer
{
    public class Reducer
    {
        public ReducerOptions Options { get; set; }

        public Reducer(bool deleteBookmarks = true,
                       bool createNewStyles = true)
        {
            Options = new ReducerOptions(
                deleteBookmarks: deleteBookmarks,
                createNewStyles: createNewStyles);
        }

        public void Reduce(WordprocessingDocument docx)
        {
            var docRoot = docx.MainDocumentPart.RootElement;

            var servicesProvider = ServicesFactory.CreateServiceProvider(docx, Options);

            servicesProvider.GetService<ParagraphProcessor>().ProcessAllParagraphs(docRoot);
        }

        public WordprocessingDocument Reduce(string pathToFile)
        {
            var docx = WordprocessingDocument.Open(pathToFile, true);

            Reduce(docx);

            return docx;
        }
    }
}
