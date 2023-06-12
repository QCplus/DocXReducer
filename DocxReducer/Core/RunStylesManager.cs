using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxReducer.Helpers;

namespace DocxReducer.Core
{
    internal class RunStylesManager
    {
        private Styles DocStyles { get; }

        private Dictionary<int, string> _globalRunStylesIds = new Dictionary<int, string>();

        public RunStylesManager(Styles docStyles)
        {
            DocStyles = docStyles;
        }

        private void PasteRunStyle(Run run, string styleId)
        {
            run.RunProperties.RemoveAllChildren();
            run.RunProperties.RunStyle = new RunStyle() { Val = styleId };
        }

        private void AppendStyle(Style runStyle, int propertiesHash)
        {
            if (!_globalRunStylesIds.ContainsKey(propertiesHash))
            {
                _globalRunStylesIds[propertiesHash] = runStyle.StyleId.Value;
                DocStyles.Append(runStyle);
            }
        }

        private void ReplaceRunPropertiesWithStyle(Run run, Style runStyle, int propertiesHash)
        {
            var styleId = runStyle.StyleId.Value;
            if (styleId == null)
                throw new Exception("StyleId for run is null");

            AppendStyle(runStyle, propertiesHash);

            PasteRunStyle(run, styleId);
        }

        private string DeleteXmlnsAttr(string xml)
        {
            return Regex.Replace(xml, "xmlns:[^=]*=\"[^\"]*\"", "");
        }

        private void CloneChildrenToTarget(OpenXmlElement source, OpenXmlElement target)
        {
            foreach (var c in source.Elements())
                target.Append(c.CloneNode(true));
        }

        internal string GenerateStyleId(int runPropertiesHash)
        {
            if (_globalRunStylesIds.TryGetValue(runPropertiesHash, out string styleId))
                return styleId;

            var createdStylesIds = DocStyles.Descendants<Style>().Where(t => t.StyleId.HasValue).Select(t => t.StyleId.Value);
            int charsPoolCount = IntExtensions.CHARS_FOR_BASE_CONVERSION.Length;

            for (int maxBits = 10; maxBits <= 32; maxBits++)
            {
                styleId = runPropertiesHash.TakeFirstBits(maxBits).ToBase(charsPoolCount);

                if (!createdStylesIds.Contains(styleId))
                    return styleId;
            }

            throw new Exception("Max int bits overflow");
        }

        internal Style CreateStyleForRun(RunProperties runProperties, int propertiesHash)
        {
            var styleRunProperties = new StyleRunProperties();
            CloneChildrenToTarget(runProperties, styleRunProperties);

            return new Style()
            {
                StyleId = GenerateStyleId(propertiesHash),
                Type = StyleValues.Character,
                StyleRunProperties = styleRunProperties
            };
        }

        internal bool IsStyleReplacementWorthIt(RunProperties runProperties, Style runStyle)
        {
            return runProperties != null
                && runStyle.InnerXml.Length < runProperties.InnerXml.Length + runStyle.StyleId.Value.Length;
        }

        internal int GenerateHash(RunProperties runProperties)
        {
            return runProperties.InnerXml.GetHashCode();
        }

        public void ReplaceRunPropertiesWithStyleIfNeeded(Run run)
        {
            var runProperties = run.RunProperties;
            var propertiesHash = GenerateHash(runProperties);

            var runStyle = CreateStyleForRun(runProperties, propertiesHash);

            if (IsStyleReplacementWorthIt(runProperties, runStyle))
                ReplaceRunPropertiesWithStyle(run, runStyle, propertiesHash);
        }
    }
}
