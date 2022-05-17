using System.Linq;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxReducer.Core
{
    internal class TagDestroyer
    {
        public void RemoveBookmarks(WordprocessingDocument docx)
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

        public void RemoveProofErrors(WordprocessingDocument docx)
        {
            var proofErrors = docx.MainDocumentPart.RootElement.Descendants<ProofError>().ToList();

            foreach (var pe in proofErrors)
            {
                pe.Remove();
            }
        }
    }
}
