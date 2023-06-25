using DocumentFormat.OpenXml;

namespace DocxReducer.Processors.Abstract
{
    internal interface IElementsProcessor
    {
        bool NeedProcessChildren(OpenXmlElement element);

        void Process(OpenXmlElement element);
    }
}
