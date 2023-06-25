using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxReducer.Processors.Abstract;

#if DEBUG
[assembly: InternalsVisibleTo("DocxReducerTests")]
#endif
namespace DocxReducer.Processors
{
    internal class ParagraphsProcessor : IElementsProcessor
    {
        public bool NeedProcessChildren(OpenXmlElement element) => true;

        public void Process(OpenXmlElement element)
        {
            if (element.GetType() != typeof(Paragraph))
                throw new System.Exception($"OpenXmlElement should be {nameof(Paragraph)}");

            var par = (Paragraph)element;

            par.ClearAllAttributes();
        }
    }
}
