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

        public Reducer(bool deleteBookmarks=true,
                       bool createNewStyles=true)
        {
            DeleteBookmarks = deleteBookmarks;
            CreateNewStyles = createNewStyles;
        }

        private void RemoveBookmarks(WordprocessingDocument docx)
        {
            var rootElement = docx.MainDocumentPart.RootElement;

            foreach (var bm in rootElement.Descendants<BookmarkStart>())
            {
                bm.Remove();
            }

            foreach (var bm in rootElement.Descendants<BookmarkEnd>())
            {
                bm.Remove();
            }
        }

        private void RemoveProofErrors(WordprocessingDocument docx)
        {
            var proofErrors = docx.MainDocumentPart.RootElement.Descendants<ProofError>().ToList();
            foreach (var pe in proofErrors)
            {
                pe.Remove();
            }
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
            RemoveProofErrors(docx);

            if (DeleteBookmarks)
                RemoveBookmarks(docx);

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
