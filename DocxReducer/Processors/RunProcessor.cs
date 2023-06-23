﻿using System.Linq;
using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxReducer.Core;

#if DEBUG
[assembly: InternalsVisibleTo("DocxReducerTests")]
#endif
namespace DocxReducer.Processors
{
    // For every new document run processor must be new
    internal sealed class RunProcessor
    {
        private RunStylesManager RunStyleCreator { get; }

        public RunProcessor(RunStylesManager runStyleCreator)
        {
            RunStyleCreator = runStyleCreator;
        }

        public bool AreEqual(RunProperties rPr1, RunProperties rPr2)
        {
            // NOTE: If XML attributes order is different but properties are the same?
            return (rPr1 == null && rPr2 == null) || !(rPr1 == null ^ rPr2 == null) && rPr1.InnerXml == rPr2.InnerXml;
        }

        private Text MoveChildren(Run target, Run source)
        {
            var children = source.ChildElements.Where(elem => 
                elem.GetType() != typeof(Text) && 
                elem.GetType() != typeof(RunProperties)
                ).ToList();

            var runText = source.GetFirstChild<Text>();

            source.RemoveAllChildren();

            target.Append(children);

            return runText;
        }

        internal void MergeRunToFirst(Run targetRun, Run sourceRun)
        {
            Text sourceRunText = MoveChildren(targetRun, sourceRun);

            if (sourceRunText != null)
            {
                var targetRunLastChild = targetRun.LastChild;

                if (targetRunLastChild == null || targetRunLastChild.GetType() != typeof(Text))
                    targetRun.Append(sourceRunText);
                else
                    (targetRunLastChild as Text).Text += sourceRunText.Text;

                if (sourceRunText?.Space != null && sourceRunText.Space.Value == SpaceProcessingModeValues.Preserve)
                    targetRun.GetFirstChild<Text>().Space = SpaceProcessingModeValues.Preserve;
            }
        }

        internal bool HasExtraElement(Run run)
        {
            return run.Elements().Where(
                t => typeof(Text) != t.GetType() && typeof(RunProperties) != t.GetType()
                ).Count() > 0;
        }

        internal bool CanMerge(Run run1, Run run2)
        {
            return AreEqual(run1.RunProperties, run2.RunProperties);
        }

        /// <summary>
        /// Merge second run to first if it's possible
        /// </summary>
        /// <param name="run1"></param>
        /// <param name="run2"></param>
        /// <returns>New base run</returns>
        public Run MergeIfNeeded(Run run1, Run run2)
        {
            if (CanMerge(run1, run2))
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