using System;
using DocxReducer.Processors;
using DocxReducer.Processors.Abstract;

namespace DocxReducer.Extensions
{
    internal static class ServiceProviderExtensions
    {
        public static bool TryGetService<T>(this IServiceProvider serviceProvider, Type type, out T service)
        {
            service = (T)serviceProvider.GetService(type);

            return service != null;
        }

        public static bool TryGetProcessor(this IServiceProvider serviceProvider, Type type, out IElementsProcessor processor)
        {
            return serviceProvider.TryGetService(type, out processor);
        }

        public static IElementsProcessor GetProcessor(this IServiceProvider serviceProvider, Type type)
        {
            if (serviceProvider.TryGetProcessor(type, out IElementsProcessor processor))
            {
                return processor;
            }

            return EmptyProcessor.Instance;
        }
    }
}
