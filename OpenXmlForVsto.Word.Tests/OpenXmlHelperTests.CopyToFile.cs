using DocumentFormat.OpenXml.Packaging;
using NUnit.Framework;
using System;
using System.Linq;
using Interop = Microsoft.Office.Interop.Word;
using OpenXml = DocumentFormat.OpenXml.Wordprocessing;
using Range = Microsoft.Office.Interop.Word.Range;

namespace OpenXmlForVsto.Word.Tests {
    [TestFixture]
    public partial class OpenXmlHelperTests {
        [Test]
        public void CopyToFile_Test_CreatesFileForValidRange() {
            Range range = _application.Documents.Add().Range();

            _tmpFile = new OpenXmlHelper().CopyToFile(range);

            FileAssert.Exists(_tmpFile);
        }

        [Test]
        public void CopyToFile_Test_ThrowsNREForNullRange() {
            Assert.Throws<ArgumentNullException>(() => new OpenXmlHelper().CopyToFile(null));
        }

        [Test]
        public void CopyToFile_Test_CreatesValidFile() {
            Range range = _application.Documents.Add().Range();

            _tmpFile = new OpenXmlHelper().CopyToFile(range);

            Assert.DoesNotThrow(() => {
                using (WordprocessingDocument.Open(_tmpFile, true)) { }
            });
        }

        [Test]
        public void CopyToFile_Test_CopiesCorrectData() {
            Range range = _application.Documents.Add().Range();
            const string expected = "test";
            range.Text = expected;

            _tmpFile = new OpenXmlHelper().CopyToFile(range);

            using (WordprocessingDocument doc = WordprocessingDocument.Open(_tmpFile, true)) {
                string result = doc.MainDocumentPart
                    .Document
                    .Body
                    .Elements<OpenXml.Paragraph>()
                    .First()
                    .InnerText;

                Assert.That(result, Is.EqualTo(expected));
            }
        }

        [Test]
        public void CopyToFile_Test_CopiesCorrectStyle() {
            Range range = _application.Documents.Add().Range();

            range.Font.Bold = 1;
            range.Font.Italic = 1;
            range.Font.Color = Interop.WdColor.wdColorGold;
            range.Font.Shading.BackgroundPatternColor = Interop.WdColor.wdColorDarkRed;

            _tmpFile = new OpenXmlHelper().CopyToFile(range);

            using (WordprocessingDocument doc = WordprocessingDocument.Open(_tmpFile, true)) {
                OpenXml.ParagraphMarkRunProperties properties = doc.MainDocumentPart
                    .Document
                    .Body
                    .Elements<OpenXml.Paragraph>()
                    .First()
                    .ParagraphProperties
                    .ParagraphMarkRunProperties;

                Assert.IsTrue(properties.Elements<OpenXml.Bold>().Any());
                Assert.IsTrue(properties.Elements<OpenXml.Italic>().Any());
                Assert.That(properties.Elements<OpenXml.Color>().Single().Val.Value, Is.EqualTo("FFCC00"));
                Assert.That(properties.Elements<OpenXml.Shading>().Single().Fill.Value, Is.EqualTo("800000"));
            }
        }

        [Test]
        public void CopyToFile_Test_CopiesCorrectHyperlinksToBookmarks() {
            Range range = _application.Documents.Add().Range();
            const string bookmarkName = "bookmark1";

            Interop.Paragraph paragraph1 = range.Paragraphs.Add();
            paragraph1.Range.Text = "bookmark";
            Interop.Bookmark bookmark = paragraph1.Range.Bookmarks.Add(bookmarkName);

            range.Paragraphs.Add();
            Interop.Paragraph paragraph2 = range.Paragraphs.Add();
            paragraph2.Range.Text = "link";
            range.Hyperlinks.Add(paragraph2.Range, SubAddress: bookmark);

            _tmpFile = new OpenXmlHelper().CopyToFile(range);

            using (WordprocessingDocument doc = WordprocessingDocument.Open(_tmpFile, true)) {
                OpenXml.Body body = doc.MainDocumentPart
                    .Document
                    .Body;

                OpenXml.BookmarkStart bookmarkStart = body.Descendants<OpenXml.BookmarkStart>().First();
                OpenXml.Hyperlink hyperlink = body.Descendants<OpenXml.Hyperlink>().First();

                Assert.That(bookmarkStart.Name.Value, Is.EqualTo(bookmarkName));
                Assert.That(hyperlink.Anchor.Value, Is.EqualTo(bookmarkName));
            }
        }

        [Test]
        public void CopyToFile_Test_CopiesCorrectHyperlinks_WithMissingBookmarks() {
            Range range = _application.Documents.Add().Range();
            const string bookmarkName = "bookmark1";

            Interop.Paragraph paragraph1 = range.Paragraphs.Add();
            paragraph1.Range.Text = "bookmark";
            Interop.Bookmark bookmark = paragraph1.Range.Bookmarks.Add(bookmarkName);

            range.Paragraphs.Add();
            Interop.Paragraph paragraph2 = range.Paragraphs.Add();
            paragraph2.Range.Text = "link";
            range.Hyperlinks.Add(paragraph2.Range, SubAddress: bookmark);

            _tmpFile = new OpenXmlHelper().CopyToFile(paragraph2.Range);
            _application.DisplayAlerts = Interop.WdAlertLevel.wdAlertsNone;
            _application.Windows[1].Close(Interop.WdSaveOptions.wdDoNotSaveChanges);

            using (WordprocessingDocument doc = WordprocessingDocument.Open(_tmpFile, true)) {
                OpenXml.Hyperlink hyperlink = doc.MainDocumentPart
                    .Document
                    .Body
                    .Descendants<OpenXml.Hyperlink>()
                    .First();

                Assert.That(hyperlink.Anchor.Value, Is.EqualTo(bookmarkName));
            }
        }

        [Test]
        public void CopyToFile_Test_ClipboardIsEmpty() {
            Range range = _application.Documents.Add().Range();
            const string expected = "test";
            range.Text = expected;

            _tmpFile = new OpenXmlHelper().CopyToFile(range);

            Assert.True(IsClipboardEmpty());
        }
    }
}