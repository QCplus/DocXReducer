﻿using System;
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

            RunProcessor = new RunProcessor(
                new RunStyleCreator(docStyles));

            Options = new ParagraphProcessorOptions(reducerOptions);
        }

        private void ReplacePropertiesWithStyles(Paragraph par)
        {
            foreach (var r in par.Elements<Run>())
                RunProcessor.ReplaceRunPropertiesWithStyle(r);
        }

        private bool NeedToBeRemoved(OpenXmlElement element)
        {
            var type = element.GetType();

            if (type == typeof(BookmarkStart) || type == typeof(BookmarkEnd))
            {
                return Options.DeleteBookmarks;
            }
            else if (type == typeof(ProofError))
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="baseRun"></param>
        /// <param name="nextElement"></param>
        /// <returns>New base run</returns>
        private Run Process(Run baseRun, OpenXmlElement nextElement)
        {
            if (nextElement is Run parRun)
            {
                nextElement.ClearAllAttributes();

                return baseRun == null
                    ? parRun
                    : RunProcessor.MergeIfNeeded(baseRun, parRun);
            }
            else
            {
                if (NeedToBeRemoved(nextElement))
                    nextElement.Remove();

                return null;
            }
        }

        public void Process(Paragraph par)
        {
            par.ClearAllAttributes();

            var children = par.ChildElements.ToList();
            Run baseRun = null;

            foreach (var child in children)
            {
                baseRun = Process(baseRun, child);
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
