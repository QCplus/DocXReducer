using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocxReducer.Processors.Abstract;

namespace DocxReducer.Processors
{
    public class ProcessorCollection
    {
        private readonly Dictionary<Type, IElementProcessor> _processors;

        public ProcessorCollection()
        {

        }

        public IElementProcessor this[OpenXmlElement element]
        {
            get => _processors[element.GetType()];
        }

        public bool TryGetValue(OpenXmlElement element, out IElementProcessor processor)
        {
            var type = element.GetType();

            return _processors.TryGetValue(type, out processor);
        }

        /// <summary>
        /// Don't throw exception on unprocessable element
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="element"></param>
        public void ProcessSafe(WordprocessingDocument doc, OpenXmlElement element)
        {
            if (TryGetValue(element, out IElementProcessor processor))
            {
                processor.Process(doc, element);
            }
        }
    }
}
