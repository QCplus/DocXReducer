using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace DocxReducer.Processors.Abstract
{
    public interface IElementProcessor
    {
        void Process(WordprocessingDocument doc, OpenXmlElement element);
    }
}
