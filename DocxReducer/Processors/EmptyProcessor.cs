using DocumentFormat.OpenXml;
using DocxReducer.Processors.Abstract;

namespace DocxReducer.Processors
{
    public class EmptyProcessor : IElementsProcessor
    {
        public bool NeedProcessChildren(OpenXmlElement element) => true;

        public void Process(OpenXmlElement element) { }
    }
}
