using Microsoft.Office.Interop.Word;
using System;
using System.IO;
using System.Runtime.InteropServices;

namespace OpenXmlForVsto.Word {
    public class OpenXmlHelper {
        /// <summary>
        /// Copy a range to a separate .docx file.
        /// It works exactly same as manual copy from one Word file to a new file,
        /// then saving and closing new file,
        /// except it removes trailing empty paragraph that Word adds when coping ranges with new line at the end.
        /// Note: remove this temporary .docx file after using,
        /// because processes have limitation on amount of created temporary files.
        /// </summary>
        /// <param name="sourceRange">Single-area range that will be copied.</param>
        /// <returns>Full path to a new temporary .docx file with copied range.</returns>
        public string CopyToFile(Range sourceRange) {
            if (sourceRange == null) throw new ArgumentNullException(nameof(sourceRange));

            Document targetDocument = CreateTargetDocument(sourceRange);

            try {
                Copy(sourceRange, targetDocument);

                return SaveAndClose(targetDocument);
            }
            finally {
                ClearClipboard();
            }
        }

        /// <summary>
        /// Copy a range to a separate .docx file.
        /// It works exactly same as manual copy text only from a Word file to a new file,
        /// then saving and closing new file,
        /// Note: remove this temporary .docx file after using,
        /// because processes have limitation on amount of created temporary files.
        /// </summary>
        /// <param name="sourceRange">Single-area range that will be copied.</param>
        /// <returns>Full path to a new temporary .docx file with copied range.</returns>
        public string CopyToFileTextOnly(Range sourceRange) {
            if (sourceRange == null) throw new ArgumentNullException(nameof(sourceRange));

            Document targetDocument = CreateTargetDocument(sourceRange);

            try {
                CopyTextOnly(sourceRange, targetDocument);

                return SaveAndClose(targetDocument);
            }
            finally {
                ClearClipboard();
            }
        }

        /// <summary>
        /// Copy whole document range from a provided file to the range.
        /// It works exactly same as manual open of provided Word file,
        /// copy whole content from it to another document and closing provided file,
        /// except it removes trailing empty paragraph that Word adds when coping ranges with new line at the end.
        /// </summary>
        /// <param name="sourceFile">Full path to .docx file.</param>
        /// <param name="targetRange">Single-area range to copy to.</param>
        public void CopyFromFile(string sourceFile, Range targetRange) {
            if (sourceFile == null) throw new ArgumentNullException(nameof(sourceFile));
            if (!File.Exists(sourceFile)) throw new FileNotFoundException("File does not exist", sourceFile);
            if (targetRange == null) throw new ArgumentNullException(nameof(targetRange));

            Document sourceDocument = targetRange.Application.Documents.Open(sourceFile, Visible: false);

            try {
                Copy(sourceDocument.Range(), targetRange);
                sourceDocument.Close();
            }
            finally {
                ClearClipboard();
            }
        }

        /// <summary>
        /// Copy whole document text from a provided file to the range.
        /// It works exactly same as manual open of provided Word file,
        /// copy whole text from it to another document and closing provided file,
        /// </summary>
        /// <param name="sourceFile">Full path to .docx file.</param>
        /// <param name="targetRange">Single-area range to copy to.</param>
        public void CopyFromFileTextOnly(string sourceFile, Range targetRange) {
            if (sourceFile == null) throw new ArgumentNullException(nameof(sourceFile));
            if (!File.Exists(sourceFile)) throw new FileNotFoundException("File does not exist", sourceFile);
            if (targetRange == null) throw new ArgumentNullException(nameof(targetRange));

            Document sourceDocument = targetRange.Application.Documents.Open(sourceFile, Visible: false);

            try {
                CopyTextOnly(sourceDocument.Range(), targetRange);
                sourceDocument.Close();
            }
            finally {
                ClearClipboard();
            }
        }

        static void Copy(Range source, Document target) => Copy(source, target.Range());

        static void Copy(Range source, Range target) {
            source.Copy();
            target.Paste();
            if (source.Paragraphs.Count < target.Paragraphs.Count) {
                Paragraph lastParagraph = target.Paragraphs[target.Paragraphs.Count];
                lastParagraph.Range.Delete();
            }
        }

        static void CopyTextOnly(Range source, Document target) => CopyTextOnly(source, target.Range());

        static void CopyTextOnly(Range source, Range target) {
            source.Copy();
            target.PasteSpecial(DataType: WdPasteDataType.wdPasteText);
        }

        static Document CreateTargetDocument(Range sourceRange) =>
            sourceRange.Application.Documents.Add(Visible: false);

        static string GetOrCreateTmpDirectory() {
            string tmpDirPath = Path.Combine(Path.GetTempPath(), "OpenXmlForVsto");
            if (!Directory.Exists(tmpDirPath)) Directory.CreateDirectory(tmpDirPath);
            return tmpDirPath;
        }

        static string GetNewRandomFilePath(string directory) =>
            Path.Combine(directory, Path.GetRandomFileName()) + ".docx";

        static string SaveAndClose(Document document) {
            string tmpFilePath = GetNewRandomFilePath(GetOrCreateTmpDirectory());
            document.SaveAs(tmpFilePath, WdSaveFormat.wdFormatXMLDocument);
            document.Close();
            return tmpFilePath;
        }

        static void ClearClipboard() {
            OpenClipboard(IntPtr.Zero);
            EmptyClipboard();
            CloseClipboard();
        }

        [DllImport("user32.dll")]
        static extern bool OpenClipboard(IntPtr hWndNewOwner);

        [DllImport("user32.dll")]
        static extern bool EmptyClipboard();

        [DllImport("user32.dll")]
        static extern bool CloseClipboard();
    }
}