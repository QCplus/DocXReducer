using System;
using System.Collections.Generic;
using System.Text;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxReducerTests
{
    class TestDataGenerator
    {
        private RunFonts GenerateUsualRunFonts()
        {
            return new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", ComplexScript = "Times New Roman" };
        }

        public ParagraphProperties GenerateParagraphProperties()
        {
            ParagraphProperties paragraphProperties = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties = new ParagraphMarkRunProperties();

            paragraphMarkRunProperties.Append(GenerateUsualRunFonts());
            paragraphMarkRunProperties.Append(new FontSize() { Val = "24" });
            paragraphMarkRunProperties.Append(new FontSizeComplexScript() { Val = "24" });
            paragraphMarkRunProperties.Append(new Languages() { Val = "en-US" });

            paragraphProperties.Append(paragraphMarkRunProperties);

            return paragraphProperties;
        }

        public RunProperties GenerateRunProperties()
        {
            RunProperties runProperties = new RunProperties();
            runProperties.Append(GenerateUsualRunFonts());
            runProperties.Append(new FontSize() { Val = "24" });
            runProperties.Append(new FontSizeComplexScript() { Val = "24" });
            runProperties.Append(new Languages() { Val = "en-US" });

            return runProperties;
        }

        public RunProperties GenerateRunPropertiesBold()
        {
            RunProperties runPropertiesBold = new RunProperties();
            runPropertiesBold.Append(GenerateUsualRunFonts());
            runPropertiesBold.Append(new Bold());
            runPropertiesBold.Append(new BoldComplexScript());
            runPropertiesBold.Append(new FontSize() { Val = "24" });
            runPropertiesBold.Append(new FontSizeComplexScript() { Val = "24" });
            runPropertiesBold.Append(new Languages() { Val = "en-US" });

            return runPropertiesBold;
        }

        public Run GenerateRun(RunProperties rp, string text)
        {
            var run = new Run();
            run.Append(rp);

            var runText = new Text() { Text = text };
            if (text.StartsWith(' ') || text.EndsWith(' '))
                runText.Space = SpaceProcessingModeValues.Preserve;
            run.Append(runText);

            return run;
        }

        public Run GenerateRun(string text)
        {
            return GenerateRun(
                GenerateRunProperties(),
                text);
        }
    }
}
