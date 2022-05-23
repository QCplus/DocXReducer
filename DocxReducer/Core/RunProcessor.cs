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

        private Dictionary<int, string> _globalRunStylesIds = new Dictionary<int, string>();

        private Dictionary<int, RunProperties> _previousRunProperties = new Dictionary<int, RunProperties>();

        public RunProcessor(Styles docStyles)
        {
            DocStyles = docStyles;
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
            if (AreEqual(run1.RunProperties, run2.RunProperties))
            {
                MergeRunToFirst(run1, run2);

                run2.Remove();

                return run1;
            }

            return run2;
        }

        private StyleRunProperties ConvertRunPropertiesToStyle(RunProperties properties)
        {
            var styleProperties = new StyleRunProperties();
            foreach (var c in properties.Elements())
                styleProperties.Append(c.CloneNode(true));

            return styleProperties;
        }

        internal Style ConvertRunPropertiesToStyle(RunProperties runProperties, string styleId)
        {
            var styleRunProperties = ConvertRunPropertiesToStyle(runProperties);

            return new Style()
            {
                StyleId = styleId,
                Type = StyleValues.Character,
                //CustomStyle = true,
                StyleRunProperties = styleRunProperties
            };
        }

        /// <summary>
        /// Generates StyleId for hash, assuming that StyleId wasn't generated for this hash before
        /// </summary>
        /// <param name="runPropertiesHash"></param>
        /// <returns></returns>
        internal string GenerateNewStyleId(int runPropertiesHash)
        {
            var styleId = $"s{runPropertiesHash}";
            var createdStylesIds = DocStyles.Descendants<Style>().Where(t => t.StyleId.HasValue).Select(t => t.StyleId.Value);

            while (createdStylesIds.Contains(styleId))
                styleId += "1";

            return styleId;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="runProperties"></param>
        /// <param name="generatedHash"></param>
        /// <returns>Created style id</returns>
        internal string CreateGlobalRunStyle(RunProperties runProperties, int generatedHash)
        {
            var styleId = GenerateNewStyleId(generatedHash);

            DocStyles.Append(
                ConvertRunPropertiesToStyle(runProperties, styleId));

            _globalRunStylesIds[generatedHash] = styleId;

            return styleId;
        }

        private string DeleteXmlnsAttr(string xml)
        {
            return Regex.Replace(xml, "xmlns:[^=]*=\"[^\"]*\"", "");
        }

        private int CalcXmlLengthWithStyle(RunProperties runProperties, int propertiesHash)
        {
            var propertiesLength = DeleteXmlnsAttr(runProperties.OuterXml).Length;

            return propertiesLength + EMPTY_RUN_STYLE_LENGTH + 2 * propertiesHash.ToString().Length + EMPTY_RUN_PROPERTIES_STYLE_LENGTH;
        }

        /// <summary>
        /// Checks if style creation for run will reduce length of xml
        /// </summary>
        /// <param name="run"></param>
        /// <returns></returns>
        public bool IsStyleCreationWorthIt(Run run, int propertiesHash)
        {
            if (run.RunProperties == null)
                return false;

            var rPrLength = DeleteXmlnsAttr(run.RunProperties.InnerXml).Length;

            // Check if there were properties with this hash
            // If so then add sum of lengths to rPrLength

            return rPrLength > CalcXmlLengthWithStyle(run.RunProperties, propertiesHash);
        }

        public void ReplaceRunPropertiesWithStyleIfNeeded(Run run)
        {
            if (run == null || run.RunProperties == null || run.Elements<RunStyle>().Count() > 0)
                return;

            var propertiesHash = run.RunProperties.InnerXml.GetHashCode();

            if (!IsStyleCreationWorthIt(run, propertiesHash))
                return;

            if (!_globalRunStylesIds.TryGetValue(propertiesHash, out string styleId))
                styleId = CreateGlobalRunStyle(run.RunProperties, propertiesHash);

            run.RunProperties.RemoveAllChildren();
            run.RunProperties.RunStyle = new RunStyle() { Val = styleId };
        }
    }
}
