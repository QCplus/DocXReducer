using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxReducer.Core;
using DocxReducer.Helpers;
using DocxReducer.Options;
using DocxReducer.Processors.Abstract;

namespace DocxReducer.Processors
{
    internal class RunPropertiesProcessor : IElementsProcessor
    {
        private readonly StylesManager _stylesManager;

        private readonly ReducerOptions _options;

        public RunPropertiesProcessor(StylesManager stylesManager, ReducerOptions options)
        {
            _stylesManager = stylesManager;
            _options = options;
        }

        public bool NeedProcessChildren(OpenXmlElement element) => false;

        private void PasteStyle(RunProperties runProperties, string styleId)
        {
            runProperties.RemoveAllChildren();
            runProperties.RunStyle = new RunStyle() { Val = styleId };
        }

        private void ReplaceRunPropertiesWithStyle(RunProperties runProperties)
        {
            Style runStyle = _stylesManager.CreateStyle(runProperties);

            PasteStyle(runProperties, runStyle.StyleId.Value);
        }

        public void Process(OpenXmlElement element)
        {
            var rPr = (RunProperties)element;

            if (rPr.Elements<RunStyle>().Any())
                return;

            if (_options.CreateNewStyles && _stylesManager.IsStyleReplacementWorthIt(rPr))
                ReplaceRunPropertiesWithStyle(rPr);
        }
    }
}
