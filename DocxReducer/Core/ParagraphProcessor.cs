using System;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Collections.Generic;
using System.Text;

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
                RunProcessor.ReplaceRunPropertiesWithStyleIfNeeded(r);
        }

        public void Process(Paragraph par)
        {
            var runs = par.Elements<Run>().ToList();
            if (runs.Count() <= 1)
                return;

            var baseRun = runs.FirstOrDefault();

            foreach (var r in runs.Skip(1))
            {
                baseRun = RunProcessor.MergeIfNeeded(baseRun, r);
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
