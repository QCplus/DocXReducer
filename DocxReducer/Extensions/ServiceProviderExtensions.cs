using System;
using DocxReducer.Processors;
using DocxReducer.Processors.Abstract;
using Microsoft.Extensions.DependencyInjection;

namespace DocxReducer.Extensions
{
    internal static class ServiceProviderExtensions
    {
        public static bool TryGetService<T>(this ServiceProvider serviceProvider, Type type, out T service)
        {
            service = (T)serviceProvider.GetService(type);

            return service != null;
        }

        public static bool TryGetProcessor(this ServiceProvider serviceProvider, Type type, out IElementsProcessor processor)
        {
            return serviceProvider.TryGetService(type, out processor);
        }

        public static IElementsProcessor GetProcessor(this ServiceProvider serviceProvider, Type type)
        {
            if (serviceProvider.TryGetProcessor(type, out IElementsProcessor processor))
            {
                return processor;
            }

            return EmptyProcessor.Instance;
        }
    }
}
