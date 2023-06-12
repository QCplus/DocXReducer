using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;

using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;

using DocxReducer.Core;

namespace DocxReducer
{
    public class Reducer
    {
        public bool DeleteBookmarks { get; set; }

        public bool CreateNewStyles { get; set; }

        private TagsDestroyer TagDestroyer { get; }

        public Reducer(bool deleteBookmarks=true,
                       bool createNewStyles=true)
        {
            DeleteBookmarks = deleteBookmarks;
            CreateNewStyles = createNewStyles;

            TagDestroyer = new TagsDestroyer();
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
            TagDestroyer.RemoveProofErrors(docx);

            if (DeleteBookmarks)
                TagDestroyer.RemoveBookmarks(docx);

            var styles = GetOrCreateNewDocStyles(docx);

            // For every new document paragraph processor must be new
            new ParagraphProcessor(styles, CreateNewStyles).ProcessAllParagraphs(docx);
        }

        public WordprocessingDocument Reduce(string pathToFile)
        {
            var docx = WordprocessingDocument.Open(pathToFile, true);

            Reduce(docx);

            return docx;
        }
    }
}
