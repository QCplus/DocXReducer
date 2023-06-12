using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxReducer.Core;
using DocxReducer.Options;

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

        private Styles GetOrCreateNewDocStyles(WordprocessingDocument docx)
        {
            var styleDefinitions = docx.MainDocumentPart.StyleDefinitionsPart;

            var styles = styleDefinitions.Styles;
            if (styles == null)
                styles = styleDefinitions.Styles = new Styles();

            return styles;
        }

        public void Reduce(WordprocessingDocument docx)
        {
            var docRoot = docx.MainDocumentPart.RootElement;

            var styles = GetOrCreateNewDocStyles(docx);

            // For every new document paragraph processor must be new
            new ParagraphProcessor(styles, Options).ProcessAllParagraphs(docRoot);
        }

        public WordprocessingDocument Reduce(string pathToFile)
        {
            var docx = WordprocessingDocument.Open(pathToFile, true);

            Reduce(docx);

            return docx;
        }
    }
}
