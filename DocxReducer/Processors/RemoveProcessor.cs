using DocumentFormat.OpenXml;
using DocxReducer.Processors.Abstract;

namespace DocxReducer.Processors
{
    internal class RemoveProcessor : IElementsProcessor
    {
        private static RemoveProcessor _instance;

        public static RemoveProcessor Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new RemoveProcessor();
                return _instance;
            }
        }

        public bool NeedProcessChildren(OpenXmlElement element) => false;

        public void Process(OpenXmlElement element)
        {
            element.Remove();
        }

        private RemoveProcessor() { }
    }
}
