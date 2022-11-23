using System;
using System.Collections.Generic;
using System.Text;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxReducerTests
{
    internal class TestDataGenerator
    {
        public const string DEF_FONT_SIZE = "24";

        private RunFonts GenerateCommonRunFonts()
        {
            return new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
        }

        public ParagraphProperties GenerateParagraphProperties()
        {
            ParagraphProperties paragraphProperties = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties = new ParagraphMarkRunProperties();

            paragraphMarkRunProperties.Append(GenerateCommonRunFonts());
            paragraphMarkRunProperties.Append(new FontSize() { Val = DEF_FONT_SIZE });
            paragraphMarkRunProperties.Append(new FontSizeComplexScript() { Val = DEF_FONT_SIZE });
            paragraphMarkRunProperties.Append(new Languages() { Val = "en-US" });

            paragraphProperties.Append(paragraphMarkRunProperties);

            return paragraphProperties;
        }

        public RunProperties GenerateRunProperties()
        {
            RunProperties runProperties = new RunProperties();
            runProperties.Append(GenerateCommonRunFonts());
            runProperties.Append(new FontSize() { Val = DEF_FONT_SIZE });
            runProperties.Append(new FontSizeComplexScript() { Val = DEF_FONT_SIZE });
            runProperties.Append(new Languages() { Val = "en-US" });

            return runProperties;
        }

        public RunProperties GenerateRunPropertiesBold()
        {
            RunProperties runPropertiesBold = GenerateRunProperties();

            runPropertiesBold.Append(new Bold());
            runPropertiesBold.Append(new BoldComplexScript());

            return runPropertiesBold;
        }

        public Run GenerateRun(RunProperties rp, string text)
        {
            var run = new Run();
            run.Append(rp);

            if (!string.IsNullOrEmpty(text))
            {
                var runText = new Text() { Text = text };
                if (text.StartsWith(' ') || text.EndsWith(' '))
                    runText.Space = SpaceProcessingModeValues.Preserve;
                run.Append(runText);
            }

            return run;
        }

        public Run GenerateRun(string text)
        {
            return GenerateRun(
                GenerateRunProperties(),
                text);
        }

        public Run GenerateRun(params OpenXmlElement[] children)
        {
            var run = new Run();

            foreach (var child in children)
                run.Append(child);

            run.RunProperties = GenerateRunProperties();

            return run;
        }
    }
}
