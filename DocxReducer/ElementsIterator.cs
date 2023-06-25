using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml;
using DocxReducer.Extensions;
using DocxReducer.Processors.Abstract;
using Microsoft.Extensions.DependencyInjection;

#if DEBUG
[assembly: InternalsVisibleTo("DocxReducerTests")]
#endif
namespace DocxReducer
{
    internal class ElementsIterator
    {
        private readonly ServiceProvider _serviceProvider;

        private ElementsIterator(ServiceProvider serviceProvider)
        {
            _serviceProvider = serviceProvider;
        }

        private void Iterate(List<OpenXmlElement> elements)
        {
            foreach (OpenXmlElement e in elements)
            {
                var elementType = e.GetType();

                if (_serviceProvider.TryGetProcessor(elementType, out IElementsProcessor processor))
                {
                    if (e.HasChildren && processor.NeedProcessChildren(e))
                    {
                        Iterate(e.ChildElements.ToList());
                    }

                    processor.Process(e);
                }
                else if (e.HasChildren)
                {
                    Iterate(e.ChildElements.ToList());
                }
            }
        }

        public static void Iterate(ServiceProvider serviceProvider, List<OpenXmlElement> elements)
        {
            new ElementsIterator(serviceProvider).Iterate(elements);
        }
    }
}
