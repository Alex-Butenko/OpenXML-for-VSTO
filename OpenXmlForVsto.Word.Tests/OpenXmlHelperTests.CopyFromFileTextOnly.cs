using Microsoft.Office.Interop.Word;
using NUnit.Framework;
using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using Xceed.Words.NET;
using Hyperlink = Xceed.Document.NET.Hyperlink;

namespace OpenXmlForVsto.Word.Tests {
    [TestFixture]
    public partial class OpenXmlHelperTests {
        [Test]
        public void CopyFromFileTextOnly_Test_ThrowsNREForNullFile() {
            Range range = _application.Documents.Add().Range();
            Assert.Throws<ArgumentNullException>(() => new OpenXmlHelper().CopyFromFile(null, range));
        }

        [Test]
        public void CopyFromFileTextOnly_Test_ThrowsNREForNullRange() {
            string anyExistingFileName = Assembly.GetAssembly(typeof(OpenXmlHelperTests)).Location;
            Assert.Throws<ArgumentNullException>(() => new OpenXmlHelper().CopyFromFileTextOnly(anyExistingFileName, null));
        }

        [Test]
        public void CopyFromFileTextOnly_Test_ThrowsFileNotFoundForMissingFile() {
            string nonExistingFileName = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
            Assert.Throws<FileNotFoundException>(() => new OpenXmlHelper().CopyFromFileTextOnly(nonExistingFileName, null));
        }

        [Test]
        public void CopyFromFileTextOnly_Test_CopiesCorrectData() {
            const string expected = "test";
            _tmpFile = GetNewRandomFilePath(GetOrCreateTmpDirectory());
            using (DocX doc = DocX.Create(_tmpFile)) {
                doc.InsertParagraph().Append(expected);
                doc.Save();
            }

            Range targetRange = _application.Documents.Add().Range();
            targetRange.Text = "unexpected";

            new OpenXmlHelper().CopyFromFileTextOnly(_tmpFile, targetRange);

            string result = targetRange.Document.Range().Text.Trim();

            Assert.That(result, Is.EqualTo(expected));
        }

        [Test]
        public void CopyFromFileTextOnly_Test_TargetRangeDoesNotKeepContent() {
            _tmpFile = GetNewRandomFilePath(GetOrCreateTmpDirectory());
            using (DocX doc = DocX.Create(_tmpFile)) {
                doc.InsertParagraph().Append("test");
                doc.Save();
            }

            Range targetRange = _application.Documents.Add().Range();
            targetRange.Text = "unexpected";

            new OpenXmlHelper().CopyFromFileTextOnly(_tmpFile, targetRange);

            // Range.PasteSpecial works differently from Range.Paste.
            // Pasted content adds before target range and target range clears.
            string result = targetRange.Text;

            Assert.IsNull(result);
        }

        [Test]
        public void CopyFromFileTextOnly_Test_CopiesNoStyle() {
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

            new OpenXmlHelper().CopyFromFileTextOnly(_tmpFile, targetRange);

            Range result = targetRange.Document.Range(1, 2);

            Assert.Zero(result.Bold);
            Assert.Zero(result.Italic);
            Assert.That(result.Font.Size, Is.Not.EqualTo(20).Within(0.001));
            Assert.That(result.Font.Color, Is.Not.EqualTo(WdColor.wdColorGreen));
        }

        [Test]
        public void CopyFromFile_Test_CopiesNoHyperlinksToBookmarks() {
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

            new OpenXmlHelper().CopyFromFileTextOnly(_tmpFile, targetRange);

            Assert.Zero(targetRange.Document.Bookmarks.Count);
            Assert.Zero(targetRange.Document.Hyperlinks.Count);
        }

        [Test]
        public void CopyFromFileTextOnly_Test_ClipboardIsEmpty() {
            _tmpFile = GetNewRandomFilePath(GetOrCreateTmpDirectory());
            using (DocX doc = DocX.Create(_tmpFile)) {
                doc.InsertParagraph().Append("text");
                doc.Save();
            }

            Range targetRange = _application.Documents.Add().Range();
            targetRange.Text = "unexpected";

            new OpenXmlHelper().CopyFromFileTextOnly(_tmpFile, targetRange);

            Assert.True(IsClipboardEmpty());
        }
    }
}