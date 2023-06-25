using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocxReducer.DI;
using DocxReducer.Options;
using Microsoft.Extensions.DependencyInjection;

namespace DocxReducer
{
    public class Reducer
    {
        private ServiceProvider _serviceProvider;

        public ReducerOptions Options { get; set; }

        public Reducer(bool deleteBookmarks = true,
                       bool createNewStyles = true)
        {
            Options = new ReducerOptions(
                deleteBookmarks: deleteBookmarks,
                createNewStyles: createNewStyles);
        }

        public void Reduce(MainDocumentPart mainDocumentPart)
        {
            _serviceProvider = ServicesFactory.CreateServiceProvider(mainDocumentPart, Options);

            ElementsIterator.Iterate(
                _serviceProvider,
                mainDocumentPart.RootElement.FirstChild.ChildElements.ToList());
        }

        public void Reduce(WordprocessingDocument docx)
        {
            Reduce(docx.MainDocumentPart);
        }

        public WordprocessingDocument Reduce(string pathToFile)
        {
            var docx = WordprocessingDocument.Open(pathToFile, true);

            Reduce(docx);

            return docx;
        }
    }
}
