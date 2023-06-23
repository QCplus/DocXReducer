using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxReducer.Core;
using DocxReducer.Options;
using DocxReducer.Processors;
using Microsoft.Extensions.DependencyInjection;

namespace DocxReducer.DI
{
    internal static class ServicesFactory
    {
        private static Styles GetOrCreateNewDocStyles(WordprocessingDocument docx)
        {
            var styleDefinitions = docx.MainDocumentPart.StyleDefinitionsPart;

            var styles = styleDefinitions.Styles;
            if (styles == null)
                styles = styleDefinitions.Styles = new Styles();

            return styles;
        }

        public static ServiceProvider CreateServiceProvider(WordprocessingDocument docx, ReducerOptions reducerOptions)
        {
            return new ServiceCollection()
                .AddSingleton(reducerOptions)
                .AddSingleton<RunStylesManager>(new RunStylesManager(GetOrCreateNewDocStyles(docx)))
                .AddSingleton<ParagraphProcessor>()
                .BuildServiceProvider();
        }
    }
}
