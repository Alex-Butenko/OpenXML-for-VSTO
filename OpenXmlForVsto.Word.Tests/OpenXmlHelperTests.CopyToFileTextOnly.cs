using DocumentFormat.OpenXml.Packaging;
using NUnit.Framework;
using System;
using System.Linq;
using System.Runtime.InteropServices;
using Interop = Microsoft.Office.Interop.Word;
using OpenXml = DocumentFormat.OpenXml.Wordprocessing;
using Range = Microsoft.Office.Interop.Word.Range;

namespace OpenXmlForVsto.Word.Tests {
    [TestFixture]
    public partial class OpenXmlHelperTests {
        [Test]
        public void CopyToFileTextOnly_Test_CreatesFileForValidRange() {
            Range range = _application.Documents.Add().Range();
            range.Text = "test";

            _tmpFile = new OpenXmlHelper().CopyToFileTextOnly(range);

            FileAssert.Exists(_tmpFile);
        }

        [Test]
        public void CopyToFileTextOnly_Test_ThrowsNREForNullRange() {
            Assert.Throws<ArgumentNullException>(() => new OpenXmlHelper().CopyToFileTextOnly(null));
        }

        [Test]
        public void CopyToFileTextOnly_Test_COMExceptionOnEmptyRange() {
            Range range = _application.Documents.Add().Range();
            range.Text = "";

            Assert.Throws<COMException>(() => new OpenXmlHelper().CopyToFileTextOnly(range));
        }

        [Test]
        public void CopyToFileTextOnly_Test_CreatesValidFile() {
            Range range = _application.Documents.Add().Range();
            range.Text = "test";

            _tmpFile = new OpenXmlHelper().CopyToFileTextOnly(range);

            Assert.DoesNotThrow(() => {
                using (WordprocessingDocument.Open(_tmpFile, true)) { }
            });
        }

        [Test]
        public void CopyToFileTextOnly_Test_CopiesCorrectData() {
            Range range = _application.Documents.Add().Range();
            const string expected = "test";
            range.Text = expected;

            _tmpFile = new OpenXmlHelper().CopyToFileTextOnly(range);

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
        public void CopyToFileTextOnly_Test_CopiesWithoutStyle() {
            Range range = _application.Documents.Add().Range();

            const string expected = "test";
            range.Text = expected;
            range.Font.Bold = 1;
            range.Font.Italic = 1;
            range.Font.Color = Interop.WdColor.wdColorGold;
            range.Font.Shading.BackgroundPatternColor = Interop.WdColor.wdColorDarkRed;

            _tmpFile = new OpenXmlHelper().CopyToFileTextOnly(range);

            using (WordprocessingDocument doc = WordprocessingDocument.Open(_tmpFile, true)) {
                OpenXml.Paragraph paragraph = doc.MainDocumentPart
                    .Document
                    .Body
                    .Elements<OpenXml.Paragraph>()
                    .First();

                Assert.IsNull(paragraph.ParagraphProperties);
                Assert.That(paragraph.InnerText, Is.EqualTo(expected));
            }
        }

        [Test]
        public void CopyToFileTextOnly_Test_DoesNotKeepHyperlinksAndBookmarks() {
            Range range = _application.Documents.Add().Range();
            const string bookmarkName = "bookmark1";

            Interop.Paragraph paragraph1 = range.Paragraphs.Add();
            paragraph1.Range.Text = "bookmark";
            Interop.Bookmark bookmark = paragraph1.Range.Bookmarks.Add(bookmarkName);

            range.Paragraphs.Add();
            Interop.Paragraph paragraph2 = range.Paragraphs.Add();
            paragraph2.Range.Text = "link";
            range.Hyperlinks.Add(paragraph2.Range, SubAddress: bookmark);

            _tmpFile = new OpenXmlHelper().CopyToFileTextOnly(range);

            using (WordprocessingDocument doc = WordprocessingDocument.Open(_tmpFile, true)) {
                OpenXml.Body body = doc.MainDocumentPart
                    .Document
                    .Body;

                CollectionAssert.IsEmpty(body.Descendants<OpenXml.BookmarkStart>());
                CollectionAssert.IsEmpty(body.Descendants<OpenXml.Hyperlink>());
            }
        }

        [Test]
        public void CopyToFileTextOnly_Test_ClipboardIsEmpty() {
            Range range = _application.Documents.Add().Range();
            range.Text = "test";

            _tmpFile = new OpenXmlHelper().CopyToFileTextOnly(range);

            Assert.True(IsClipboardEmpty());
        }
    }
}