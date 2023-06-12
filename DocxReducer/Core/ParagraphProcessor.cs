using System;
using System.Linq;
using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocxReducer.Options;

#if DEBUG
[assembly: InternalsVisibleTo("DocxReducerTests")]
#endif
namespace DocxReducer.Core
{
    /// <summary>
    /// For every new document paragraph processor must be new
    /// </summary>
    internal sealed class ParagraphProcessor
    {
        private RunProcessor RunProcessor { get; }

        private ParagraphProcessorOptions Options { get; }

        public ParagraphProcessor(Styles docStyles, ReducerOptions reducerOptions)
        {
            if (docStyles == null)
                throw new Exception("Document styles can't be null");

            RunProcessor = new RunProcessor(docStyles);

            Options = new ParagraphProcessorOptions(reducerOptions);
        }

        private void ReplacePropertiesWithStyles(Paragraph par)
        {
            foreach (var r in par.Elements<Run>())
                RunProcessor.ReplaceRunPropertiesWithStyle(r);
        }

        private void RemoveIfNecessary(OpenXmlElement element)
        {
            var type = element.GetType();

            if (type == typeof(BookmarkStart) || type == typeof(BookmarkEnd))
                if (Options.DeleteBookmarks)
                    element.Remove();
            else if (type == typeof(ProofError))
                    element.Remove();
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

                    RemoveIfNecessary(child);
                }
            }

            // NOTE: little file in zip can be bigger than big file. Zip compression nuance?
            if (Options.CreateNewStyles)
                ReplacePropertiesWithStyles(par);
        }

        public void ProcessAllParagraphs(OpenXmlPartRootElement root)
        {
            foreach (var p in root.Descendants<Paragraph>())
            {
                Process(p);
            }
        }

        public void ProcessAllParagraphs(WordprocessingDocument docx)
        {
            ProcessAllParagraphs(docx.MainDocumentPart.RootElement);
        }
    }
}
