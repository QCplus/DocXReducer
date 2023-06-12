using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxReducer.Core
{
    internal class TagsDestroyer
    {
        public void RemoveBookmarks(OpenXmlPartRootElement docRoot)
        {
            foreach (var bm in docRoot.Descendants<BookmarkStart>())
            {
                bm.Remove();
            }

            foreach (var bm in docRoot.Descendants<BookmarkEnd>())
            {
                bm.Remove();
            }
        }

        public void RemoveProofErrors(OpenXmlPartRootElement docRoot)
        {
            var proofErrors = docRoot.Descendants<ProofError>().ToList();

            foreach (var pe in proofErrors)
            {
                pe.Remove();
            }
        }
    }
}
