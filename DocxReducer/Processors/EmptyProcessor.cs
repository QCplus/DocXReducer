using DocumentFormat.OpenXml;
using DocxReducer.Processors.Abstract;

namespace DocxReducer.Processors
{
    internal class EmptyProcessor : IElementsProcessor
    {
        private static EmptyProcessor _instance;

        public static EmptyProcessor Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new EmptyProcessor();
                return _instance;
            }
        }

        public bool NeedProcessChildren(OpenXmlElement element) => true;

        public void Process(OpenXmlElement element) { }

        private EmptyProcessor() { }
    }
}
