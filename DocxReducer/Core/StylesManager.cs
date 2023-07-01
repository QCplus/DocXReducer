using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxReducer.Extensions;
using DocxReducer.Helpers;

#if DEBUG
[assembly: InternalsVisibleTo("DocxReducerTests")]
#endif
namespace DocxReducer.Core
{
    internal class StylesManager
    {
        public const int AVERAGE_STYLE_ID_LEN = 3;

        private Styles DocStyles { get; }

        private Dictionary<int, string> _globalRunStylesIds = new Dictionary<int, string>();

        private readonly HashSet<string> _docStyleIds;

        public StylesManager(Styles docStyles)
        {
            DocStyles = docStyles;
            _docStyleIds = new HashSet<string>(DocStyles.Descendants<Style>().Where(t => t.StyleId.HasValue).Select(t => t.StyleId.Value));
        }

        private int GetHash(OpenXmlElement properties)
        {
            return properties.InnerXml.GetHashCode();
        }

        private void AppendStyle(Style style, int propertiesHash)
        {
            _globalRunStylesIds[propertiesHash] = style.StyleId.Value;

            DocStyles.Append(style);
        }

        internal bool IsStyleReplacementWorthIt(OpenXmlElement properties)
        {
            return properties != null;
        }

        internal string GenerateStyleId(int runPropertiesHash)
        {
            if (_globalRunStylesIds.TryGetValue(runPropertiesHash, out string styleId))
                return styleId;

            int charsPoolCount = IntExtensions.CHARS_FOR_BASE_CONVERSION.Length;

            for (int maxBits = 10; maxBits <= 32; maxBits++)
            {
                styleId = runPropertiesHash.TakeFirstBits(maxBits).ToBase(charsPoolCount);

                if (!_docStyleIds.Contains(styleId))
                    return styleId;
            }

            throw new Exception("Max int bits overflow");
        }

        private Style CreateNewGlobalStyle(RunProperties runProperties, int propertiesHash)
        {
            var styleRunProperties = new StyleRunProperties();
            runProperties.CloneChildrenTo(styleRunProperties);

            var newStyle = new Style()
            {
                StyleId = GenerateStyleId(propertiesHash),
                Type = StyleValues.Character,
                StyleRunProperties = styleRunProperties,
            };

            AppendStyle(newStyle, propertiesHash);

            return newStyle;
        }

        public Style CreateStyle(RunProperties runProperties)
        {
            int propertiesHash = GetHash(runProperties);

            if (_globalRunStylesIds.TryGetValue(propertiesHash, out string styleId))
            {
                return DocStyles.Descendants<Style>().Where(s => s.StyleId.Value == styleId).First();
            }
            else
            {
                return CreateNewGlobalStyle(runProperties, propertiesHash);
            }
        }
    }
}
