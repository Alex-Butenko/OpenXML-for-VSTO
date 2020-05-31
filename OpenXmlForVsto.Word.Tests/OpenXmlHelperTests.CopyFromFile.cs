using Microsoft.Office.Interop.Word;
using NUnit.Framework;
using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using Xceed.Words.NET;
using Hyperlink = Xceed.Document.NET.Hyperlink;
using Range = Microsoft.Office.Interop.Word.Range;

namespace OpenXmlForVsto.Word.Tests {
    [TestFixture]
    public partial class OpenXmlHelperTests {
        [Test]
        public void CopyFromFile_Test_ThrowsNREForNullFile() {
            Range range = _application.Documents.Add().Range();
            Assert.Throws<ArgumentNullException>(() => new OpenXmlHelper().CopyFromFile(null, range));
        }

        [Test]
        public void CopyFromFile_Test_ThrowsNREForNullRange() {
            string anyExistingFileName = Assembly.GetAssembly(typeof(OpenXmlHelperTests)).Location;
            Assert.Throws<ArgumentNullException>(() => new OpenXmlHelper().CopyFromFile(anyExistingFileName, null));
        }

        [Test]
        public void CopyFromFile_Test_ThrowsFileNotFoundForMissingFile() {
            string nonExistingFileName = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
            Assert.Throws<FileNotFoundException>(() => new OpenXmlHelper().CopyFromFile(nonExistingFileName, null));
        }

        [Test]
        public void CopyFromFile_Test_CopiesCorrectData() {
            const string expected = "test";
            _tmpFile = GetNewRandomFilePath(GetOrCreateTmpDirectory());
            using (DocX doc = DocX.Create(_tmpFile)) {
                doc.InsertParagraph().Append(expected);
                doc.Save();
            }

            Range targetRange = _application.Documents.Add().Range();
            targetRange.Text = "unexpected";

            new OpenXmlHelper().CopyFromFile(_tmpFile, targetRange);

            string result = targetRange.Text.Trim();

            Assert.That(result, Is.EqualTo(expected));
        }

        [Test]
        public void CopyFromFile_Test_CopiesCorrectStyle() {
            _tmpFile = GetNewRandomFilePath(GetOrCreateTmpDirectory());
            using (DocX doc = DocX.Create(_tmpFile)) {
                doc.InsertParagraph()
                    .Append("text")
                    .Bold()
                    .Italic()
                    .Color(Color.Green)
                    .FontSize(20);
                doc.Save();
            }

            Range targetRange = _application.Documents.Add().Range();
            targetRange.Text = "unexpected";

            new OpenXmlHelper().CopyFromFile(_tmpFile, targetRange);

            Assert.NotZero(targetRange.Bold);
            Assert.NotZero(targetRange.Italic);
            Assert.That(targetRange.Font.Size, Is.EqualTo(20).Within(0.001));
            Assert.That(targetRange.Font.Color, Is.EqualTo(WdColor.wdColorGreen));
        }

        [Test]
        public void CopyFromFile_Test_CopiesCorrectHyperlinksToBookmarks() {
            _tmpFile = GetNewRandomFilePath(GetOrCreateTmpDirectory());
            using (DocX doc = DocX.Create(_tmpFile)) {
                Hyperlink hyperlink = doc.AddHyperlink("hypelink text", "bookmark");
                doc.InsertParagraph()
                    .AppendBookmark("bookmark")
                    .Append("bookmark text");
                doc.InsertParagraph()
                    .AppendHyperlink(hyperlink);
                doc.Save();
            }

            Range targetRange = _application.Documents.Add().Range();
            targetRange.Text = "unexpected";

            new OpenXmlHelper().CopyFromFile(_tmpFile, targetRange);

            Bookmarks bookmarks = targetRange.Document.Bookmarks;
            Hyperlinks hyperlinks = targetRange.Document.Hyperlinks;

            Assert.NotZero(bookmarks.Count);
            Assert.NotZero(hyperlinks.Count);
            Assert.That(hyperlinks[1].SubAddress, Is.EqualTo(bookmarks[1].Name));
        }

        [Test]
        public void CopyFromFile_Test_ClipboardIsEmpty() {
            _tmpFile = GetNewRandomFilePath(GetOrCreateTmpDirectory());
            using (DocX doc = DocX.Create(_tmpFile)) {
                doc.InsertParagraph().Append("text");
                doc.Save();
            }

            Range targetRange = _application.Documents.Add().Range();
            targetRange.Text = "unexpected";

            new OpenXmlHelper().CopyFromFile(_tmpFile, targetRange);

            Assert.True(IsClipboardEmpty());
        }
    }
}