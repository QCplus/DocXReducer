using DocumentFormat.OpenXml;
using DocxReducer.Options;
using DocxReducer.Processors.Abstract;

namespace DocxReducer.Processors
{
    internal class BookmarksProcessor : IElementsProcessor
    {
        private readonly ReducerOptions _options;

        public BookmarksProcessor(ReducerOptions options)
        {
            _options = options;
        }

        public bool NeedProcessChildren(OpenXmlElement element) => false;

        public void Process(OpenXmlElement element)
        {
            if (_options.DeleteBookmarks)
                element.Remove();
        }
    }
}
