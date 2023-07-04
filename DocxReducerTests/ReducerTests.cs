using System.Diagnostics;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxReducer;
using DocxReducer.Options;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OMath = DocumentFormat.OpenXml.Math.OfficeMath;

namespace DocxReducerTests
{
    [TestClass]
    public class ReducerTests
    {
        private static WordprocessingDocument CreateDoc(string bodyInnerXml)
        {
            WordprocessingDocument docx = null;
            using (var stream = new FileStream(@"Content\Empty.docx", FileMode.Open))
            {
                docx = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
            }

            docx.AddMainDocumentPart();
            docx.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
            docx.MainDocumentPart.Document = new Document(
                new Body($@"
                    <w:body xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
                        {bodyInnerXml}
                    </w:body>")
                );

            return docx;
        }

        private WordprocessingDocument Reduce(WordprocessingDocument docx, ReducerOptions options = null)
        {
            var reducer = new Reducer(options == null ? new ReducerOptions() : options);

            reducer.Reduce(docx);

            return docx;
        }

        private WordprocessingDocument Reduce(string bodyInnerXml, ReducerOptions options = null)
        {
            return Reduce(
                CreateDoc(bodyInnerXml),
                options);
        }

        private Body ReduceReturnBody(string bodyInnerXml, ReducerOptions options = null)
        {
            return (Body)Reduce(
                bodyInnerXml,
                options
                ).MainDocumentPart.RootElement.FirstChild;
        }

        private Body ReduceRunsReturnBody(string parInnerXml)
        {
            return ReduceReturnBody($"<w:p>{parInnerXml}</w:p>");
        }

        [TestMethod]
        public void Reduce_ReducingFromFileNotFails()
        {
            var reducer = new Reducer();

            var pathToFile = Path.Combine(@"..\..\..\..\", "TestFiles", "Holocaust.docx");

            var reducedDocx = reducer.Reduce(pathToFile);

            Assert.IsNotNull(reducedDocx);
        }

        [TestMethod]
        public void Reduce_AllRunsWithStylesWereCombined()
        {
            var xml = @"
                <w:p>
					<w:r>
						<w:rPr>
							<w:rFonts w:ascii=""Times New Roman"" w:hAnsi=""Times New Roman"" w:cs=""Times New Roman""/>
							<w:sz w:val=""24""/>
							<w:szCs w:val=""24""/>
							<w:lang w:val=""en-US""/>
						</w:rPr>
						<w:t>THIS IS</w:t>
					</w:r>
					<w:r>
						<w:rPr>
							<w:rFonts w:ascii=""Times New Roman"" w:hAnsi=""Times New Roman"" w:cs=""Times New Roman""/>
							<w:sz w:val=""24""/>
							<w:szCs w:val=""24""/>
							<w:lang w:val=""en-US""/>
						</w:rPr>
						<w:t xml:space=""preserve""> A</w:t>
					</w:r>
					<w:r>
						<w:rPr>
							<w:rFonts w:ascii=""Times New Roman"" w:hAnsi=""Times New Roman"" w:cs=""Times New Roman""/>
							<w:sz w:val=""24""/>
							<w:szCs w:val=""24""/>
							<w:lang w:val=""en-US""/>
						</w:rPr>
						<w:t xml:space=""preserve""> TEST.</w:t>
					</w:r>
                </w:p>";

            var par = ReduceReturnBody(xml).FirstChild;

            Assert.AreEqual(1, par.ChildElements.Count);
            Assert.AreEqual("THIS IS A TEST.", par.FirstChild.InnerText);
        }

        [TestMethod]
        public void Reduce_TabCharBetweenRunsSaved()
        {
            var xml = @"
	            <w:r>
	            	<w:t>TEXT</w:t>
	            </w:r>
	            <w:r>
	            	<w:tab/>
	            </w:r>
	            <w:r>
	            	<w:t>1</w:t>
	            </w:r>";

            var par = ReduceRunsReturnBody(xml).FirstChild;

            Assert.AreEqual(1, par.ChildElements.Count);

            var run = par.GetFirstChild<Run>();

            Assert.AreEqual(3, run.ChildElements.Count);
            Assert.AreEqual("TEXT", run.ElementAt(0).InnerText);
            Assert.AreEqual(typeof(TabChar), run.ElementAt(1).GetType());
            Assert.AreEqual("1", run.ElementAt(2).InnerText);
        }

        [TestMethod]
        public void Reduce_FirstRunWithoutTextWasDeleted()
        {
            var xml = @"
                <w:r>
                	<w:rPr>
                		<w:rFonts w:ascii=""Times New Roman"" w:hAnsi=""Times New Roman"" w:cs=""Times New Roman""/>
                		<w:sz w:val=""24""/>
                		<w:szCs w:val=""24""/>
                		<w:lang w:val=""en-US""/>
                	</w:rPr>
                </w:r>
                <w:r>
                	<w:rPr>
                		<w:rFonts w:ascii=""Times New Roman"" w:hAnsi=""Times New Roman"" w:cs=""Times New Roman""/>
                		<w:sz w:val=""24""/>
                		<w:szCs w:val=""24""/>
                		<w:lang w:val=""en-US""/>
                	</w:rPr>
                	<w:t>TEXT</w:t>
                </w:r>";

            var par = ReduceRunsReturnBody(xml).FirstChild;

            Assert.AreEqual(1, par.ChildElements.Count);
            Assert.AreEqual("TEXT", par.GetFirstChild<Run>().InnerText);
        }

        [TestMethod]
        public void Reduce_LastRunWithoutTextWasDeleted()
        {
            var xml = @"
                <w:r>
                	<w:rPr>
                		<w:rFonts w:ascii=""Times New Roman"" w:hAnsi=""Times New Roman"" w:cs=""Times New Roman""/>
                		<w:sz w:val=""24""/>
                		<w:szCs w:val=""24""/>
                		<w:lang w:val=""en-US""/>
                	</w:rPr>
                    <w:t>TEXT</w:t>
                </w:r>
                <w:r>
                	<w:rPr>
                		<w:rFonts w:ascii=""Times New Roman"" w:hAnsi=""Times New Roman"" w:cs=""Times New Roman""/>
                		<w:sz w:val=""24""/>
                		<w:szCs w:val=""24""/>
                		<w:lang w:val=""en-US""/>
                	</w:rPr>
                </w:r>";

            var par = ReduceRunsReturnBody(xml).FirstChild;

            Assert.AreEqual(1, par.ChildElements.Count);
            Assert.AreEqual("TEXT", par.GetFirstChild<Run>().InnerText);
        }

        [TestMethod]
        public void Reduce_MathFormulaKeepPositionBetweenRuns()
        {
            var xml = @"
                  <w:r>
                    <w:t xml:space=""preserve"">Complexity is </w:t>
                  </w:r>
                  <m:oMath xmlns:m=""http://schemas.openxmlformats.org/officeDocument/2006/math"">
                    <m:r>
                      <m:t>O</m:t>
                    </m:r>
                  </m:oMath>
                  <w:r>
                    <w:t xml:space=""preserve""> complexity</w:t>
                  </w:r>";

            var par = ReduceRunsReturnBody(xml).FirstChild;

            Assert.AreEqual(3, par.ChildElements.Count);

            Assert.AreEqual(typeof(OMath), par.ElementAt(1).GetType());
        }

        [TestMethod]
        public void Reduce_AllBookmarksWereRemoved()
        {
            var xml = @"
                <w:p w14:paraId=""0334E79D"" w14:textId=""1C39F1F7"" xmlns:w14=""http://schemas.microsoft.com/office/word/2010/wordml"">
                  <w:bookmarkStart w:name=""_Toc517125172"" w:id=""13"" />
                  <w:bookmarkStart w:name=""_Toc517125298"" w:id=""14"" />
                  <w:r>
                    <w:t>TEXT</w:t>
                  </w:r>
                  <w:bookmarkEnd w:id=""13"" />
                  <w:bookmarkEnd w:id=""14"" />
                </w:p>";

            var par = ReduceReturnBody(bodyInnerXml: xml).FirstChild;

            Assert.AreEqual(0, par.Descendants<BookmarkStart>().Count());
            Assert.AreEqual(0, par.Descendants<BookmarkEnd>().Count());
        }

        [TestMethod]
        public void Reduce_AllRsidsWereRemoved()
        {
            var xml = @"
                <w:p w:rsidRPr=""00D20A94"" w:rsidR=""00D20A94"">
                  <w:r w:rsidRPr=""00D20A94"">
                    <w:t>TEXT</w:t>
                  </w:r>
                  <w:r w:rsidRPr=""00D20A94"">
                    <w:rPr>
                      <w:color w:val=""2B91AF"" />
                      <w:sz w:val=""19"" />
                    </w:rPr>
                    <w:t xml:space=""preserve""> TEXT</w:t>
                  </w:r>
                </w:p>";

            var par = ReduceReturnBody(bodyInnerXml: xml).FirstChild;

            Assert.AreEqual(0, par.GetAttributes().Count());
            Assert.AreEqual(0, par.Descendants<Run>().Where(r => r.HasAttributes).Count());
        }

        [TestMethod]
        public void Reduce_StylesWereSavedBetweenParagraphs()
        {
            var xml = @"
                <w:p>
                  <w:r>
                    <w:rPr>
                      <w:color w:val=""2B95FF"" />
                      <w:sz w:val=""18"" />
                    </w:rPr>
                    <w:t>TEXT</w:t>
                  </w:r>
                  <w:r>
                    <w:rPr>
                      <w:color w:val=""2B91AF"" />
                      <w:sz w:val=""19"" />
                    </w:rPr>
                    <w:t xml:space=""preserve""> TEXT</w:t>
                  </w:r>
                </w:p>";

            var doc = Reduce(bodyInnerXml: xml + xml);

            Assert.That.HaveCustomStyles(doc, expectedStylesCount: 2);
            Assert.That.AllRunStylesDefined(doc);
        }

        [TestMethod]
        public void Reduce_AllProofErrorsWereRemoved()
        {
            var xml = @"
                <w:p>
                  <w:r>
                    <w:t xml:space=""preserve"">TEST TEXT </w:t>
                  </w:r>
                  <w:proofErr w:type=""spellStart"" />
                  <w:r>
                    <w:t>INVALID TEXT</w:t>
                  </w:r>
                  <w:proofErr w:type=""spellEnd"" />
                  <w:r>
                    <w:t>, TEST TEXT</w:t>
                  </w:r>
                </w:p>";

            var par = ReduceReturnBody(bodyInnerXml: xml).FirstChild;

            Assert.IsTrue(!par.Elements<ProofError>().Any(), "ProofErrors must be removed");
        }

        [TestMethod]
        public void Reduce_RunPropertiesWereReplacedWithStyle()
        {
            var xml = @"
                <w:r>
                	<w:rPr>
                		<w:rFonts w:ascii=""Times New Roman"" w:hAnsi=""Times New Roman"" w:cs=""Times New Roman""/>
                		<w:sz w:val=""24""/>
                		<w:szCs w:val=""24""/>
                		<w:lang w:val=""en-US""/>
                	</w:rPr>
                    <w:t>TEXT</w:t>
                </w:r>";

            var run = (Run)ReduceRunsReturnBody(xml).FirstChild.FirstChild;

            Assert.IsNotNull(run.RunProperties.RunStyle, "Run properies should have style");
            Assert.IsTrue(run.RunProperties.ChildElements.Count == 1);
        }

        [TestMethod]
        public void Reduce_RunWithSpaceWasMerged()
        {
            var xml = @"
                    <w:r w:rsidP=""009B6E3A"">
                        <w:t>TEST</w:t>
                    </w:r>
                    <w:r w:rsidP=""009B6E00"">
                        <w:t xml:space=""preserve""> text</w:t>
                    </w:r>";
            
            var par = ReduceRunsReturnBody(xml).FirstChild;

            Assert.AreEqual(1, par.Descendants<Text>().Count());
            Assert.AreEqual(1, par.Descendants<Run>().Count());
            Assert.AreEqual("TEST text", par.InnerText);
        }

        [TestMethod]
        public void Reduce_IgnoreBookmarksIfOptionSpecified()
        {
            var xml = @"
                <w:p w14:paraId=""0334E79D"" w14:textId=""1C39F1F7"" xmlns:w14=""http://schemas.microsoft.com/office/word/2010/wordml"">
                  <w:bookmarkStart w:name=""_Toc517125172"" w:id=""13"" />
                  <w:bookmarkStart w:name=""_Toc517125298"" w:id=""14"" />
                  <w:r>
                    <w:t>TEXT</w:t>
                  </w:r>
                  <w:bookmarkEnd w:id=""13"" />
                  <w:bookmarkEnd w:id=""14"" />
                </w:p>";

            var par = ReduceReturnBody(
                bodyInnerXml: xml,
                new ReducerOptions(deleteBookmarks: false, true)
                ).FirstChild;

            Assert.AreEqual(2, par.Descendants<BookmarkStart>().Count());
            Assert.AreEqual(2, par.Descendants<BookmarkEnd>().Count());
        }
    }
}
