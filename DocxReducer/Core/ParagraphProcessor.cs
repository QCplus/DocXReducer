using System;
using System.Linq;
using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;

#if DEBUG
[assembly: InternalsVisibleTo("DocxReducerTests")]
#endif
namespace DocxReducer.Core
{
    // For every new document paragraph processor must be new
    internal sealed class ParagraphProcessor
    {
        private RunProcessor RunProcessor { get; }

        private bool CreateNewStyles { get; }

        public ParagraphProcessor(Styles docStyles, bool createNewStyles)
        {
            if (docStyles == null)
                throw new Exception("Document styles can't be null");

            RunProcessor = new RunProcessor(docStyles);

            CreateNewStyles = createNewStyles;
        }

        private void ReplacePropertiesWithStyles(Paragraph par)
        {
            foreach (var r in par.Elements<Run>())
                RunProcessor.ReplaceRunPropertiesWithStyle(r);
        }

        public void Process(Paragraph par)
        {
            var children = par.ChildElements.ToList();
            Run baseRun = null;

            foreach (var child in children)
            {
                if (child is Run parRun)
                {
                    if (baseRun == null)
                    {
                        baseRun = parRun;
                        continue;
                    }

                    baseRun = RunProcessor.MergeIfNeeded(baseRun, parRun);
                }
                else
                {
                    baseRun = null;
                }
            }

            // NOTE: little file in zip can be bigger than big file. Zip compression nuance?
            if (CreateNewStyles)
                ReplacePropertiesWithStyles(par);
        }

        public void ProcessAllParagraphs(WordprocessingDocument docx)
        {
            foreach (var p in docx.MainDocumentPart.RootElement.Descendants<Paragraph>())
            {
                Process(p);
            }
        }
    }
}
