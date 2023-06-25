using System;
using DocumentFormat.OpenXml;
using DocxReducer.Processors.Abstract;
using Microsoft.Extensions.DependencyInjection;

namespace DocxReducer.Extensions
{
    internal static class ServiceCollectionExtensions
    {
        public static IServiceCollection AddProcessor<T>(this IServiceCollection services, Func<IServiceProvider, IElementsProcessor> factory)
            where T : OpenXmlElement
        {
            return services.AddSingleton(typeof(T), implementationFactory: factory);
        }

        public static IServiceCollection AddProcessor<T>(this IServiceCollection services, IElementsProcessor instance)
        {
            return services.AddSingleton(typeof(T), sp => instance);
        }
    }
}
