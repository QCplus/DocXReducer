using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml;
using DocxReducer.Extensions;

#if DEBUG
[assembly: InternalsVisibleTo("DocxReducerTests")]
#endif
namespace DocxReducer
{
    internal class ElementsIterator
    {
        private readonly IServiceProvider _serviceProvider;

        private ElementsIterator(IServiceProvider serviceProvider)
        {
            _serviceProvider = serviceProvider;
        }

        private void Iterate(List<OpenXmlElement> elements)
        {
            foreach (OpenXmlElement e in elements)
            {
                var elementType = e.GetType();
                var processor = _serviceProvider.GetProcessor(elementType);

                if (e.HasChildren && processor.NeedProcessChildren(e))
                {
                    Iterate(e.ChildElements.ToList());
                }

                processor.Process(e);
            }
        }

        public static void Iterate(IServiceProvider serviceProvider, List<OpenXmlElement> elements)
        {
            new ElementsIterator(serviceProvider).Iterate(elements);
        }
    }
}
