using DocumentFormat.OpenXml;

namespace DocxReducer.Processors.Abstract
{
    public interface IElementsProcessor
    {
        bool NeedProcessChildren(OpenXmlElement element);

        void Process(OpenXmlElement element);
    }
}
