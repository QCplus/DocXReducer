using System;
using System.Linq;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Text;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.RegularExpressions;

#if DEBUG
[assembly: InternalsVisibleTo("DocxReducerTests")]
#endif
namespace DocxReducer.Core
{
    // For every new document run processor must be new
    internal sealed class RunProcessor
    {
        public const int EMPTY_RUN_STYLE_LENGTH = 51;

        public const int EMPTY_RUN_PROPERTIES_STYLE_LENGTH = 20;

        private Styles DocStyles { get; }

        private RunStyleCreator RunStyleCreator { get; }

        public RunProcessor(Styles docStyles)
        {
            RunStyleCreator = new RunStyleCreator(docStyles);
        }

        public bool AreEqual(RunProperties rPr1, RunProperties rPr2)
        {
            // NOTE: If XML attributes order is different but properties are the same?
            return (rPr1 == null && rPr2 == null) || !(rPr1 == null ^ rPr2 == null) && rPr1.InnerXml == rPr2.InnerXml;
        }

        internal void MergeRunToFirst(Run run1, Run run2)
        {
            var runText = run2.GetFirstChild<Text>();

            run1.GetFirstChild<Text>().Text += runText.Text;

            if (runText.Space != null && runText.Space.Value == SpaceProcessingModeValues.Preserve)
                run1.GetFirstChild<Text>().Space = SpaceProcessingModeValues.Preserve;
        }

        /// <summary>
        /// Merge second run to first if it's possible
        /// </summary>
        /// <param name="run1"></param>
        /// <param name="run2"></param>
        /// <returns>New base run</returns>
        public Run MergeIfNeeded(Run run1, Run run2)
        {
            if (run1.GetFirstChild<Text>() == null || run2.GetFirstChild<Text>() == null)
                return run2;

            if (AreEqual(run1.RunProperties, run2.RunProperties))
            {
                MergeRunToFirst(run1, run2);

                run2.Remove();

                return run1;
            }

            return run2;
        }

        public void ReplaceRunPropertiesWithStyle(Run run)
        {
            if (run == null || run.RunProperties == null || run.Elements<RunStyle>().Count() > 0)
                return;

            RunStyleCreator.ReplaceRunPropertiesWithStyleIfNeeded(run);
        }
    }
}
