using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxReducer.Core;
using DocxReducer.Extensions;
using DocxReducer.Options;
using DocxReducer.Processors;
using Microsoft.Extensions.DependencyInjection;

#if DEBUG
[assembly: InternalsVisibleTo("DocxReducerTests")]
#endif
namespace DocxReducer.DI
{
    internal static class ServicesFactory
    {
        public static ServiceProvider CreateServiceProvider(MainDocumentPart mainDocumentPart, ReducerOptions reducerOptions)
        {
            return new ServiceCollection()
                .AddSingleton(reducerOptions)
                .AddSingleton<StylesManager>(new StylesManager(mainDocumentPart))

                .AddProcessor<Run>(sp => new RunsProcessor())
                .AddProcessor<Paragraph>(sp => new ParagraphsProcessor())
                .AddProcessor<RunProperties>(sp => new RunPropertiesProcessor(sp.GetService<StylesManager>(), sp.GetService<ReducerOptions>()))

                .AddProcessor<BookmarkStart>(sp => new BookmarksProcessor(sp.GetService<ReducerOptions>()))
                .AddProcessor<BookmarkEnd>(sp => new BookmarksProcessor(sp.GetService<ReducerOptions>()))
                .AddProcessor<ProofError>(RemoveProcessor.Instance)

                .BuildServiceProvider();
        }
    }
}
