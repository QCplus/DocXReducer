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
        private static Styles GetOrCreateNewDocStyles(MainDocumentPart mainDocumentPart)
        {
            var styleDefinitions = mainDocumentPart.StyleDefinitionsPart;

            var styles = styleDefinitions.Styles;
            if (styles == null)
                styles = styleDefinitions.Styles = new Styles();

            return styles;
        }

        public static ServiceProvider CreateServiceProvider(MainDocumentPart mainDocumentPart, ReducerOptions reducerOptions)
        {
            return new ServiceCollection()
                .AddSingleton(reducerOptions)
                .AddSingleton<StylesManager>(new StylesManager(GetOrCreateNewDocStyles(mainDocumentPart)))
                .AddProcessor<Run>(sp => new RunsProcessor())
                .AddProcessor<Paragraph>(sp => new ParagraphsProcessor())
                .AddProcessor<RunProperties>(sp => new RunPropertiesProcessor(sp.GetService<StylesManager>()))

                .AddProcessor<BookmarkStart>(RemoveProcessor.Instance)
                .AddProcessor<BookmarkEnd>(RemoveProcessor.Instance)
                .AddProcessor<ProofError>(RemoveProcessor.Instance)

                .BuildServiceProvider();
        }
    }
}
