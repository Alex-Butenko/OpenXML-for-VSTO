using Microsoft.Office.Interop.Word;
using NUnit.Framework;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;

namespace OpenXmlForVsto.Word.Tests {
    [TestFixture]
    public partial class OpenXmlHelperTests {
        [SetUp]
        public void Setup() {
            _application = new Application {
                DisplayAlerts = WdAlertLevel.wdAlertsNone
            };
        }

        [TearDown]
        public void TearDown() {
            int processId = 0;
            try {
                if (_application.Windows.Count > 0) {
                    int hWnd = _application.Windows[1].Hwnd;
                    GetWindowThreadProcessId((IntPtr)hWnd, out processId);
                }

                _application.DisplayAlerts = WdAlertLevel.wdAlertsNone;
                _application.Quit(WdSaveOptions.wdDoNotSaveChanges);
            }
            catch { }
            finally {
                try {
                    Thread.Sleep(3000);
                    Process.GetProcessById(processId).Kill();
                }
                catch { }
            }
            try {
                if (_tmpFile != null && File.Exists(_tmpFile)) File.Delete(_tmpFile);
            }
            catch { }
        }

        [DllImport("user32.dll")]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out int processId);

        Application _application;
        string _tmpFile;

        static bool IsClipboardEmpty() =>
            !typeof(DataFormats)
                .GetFields(BindingFlags.Public | BindingFlags.Static)
                .Select(x => x.Name)
                .Any(Clipboard.ContainsData);

        static string GetOrCreateTmpDirectory() {
            string tmpDirPath = Path.Combine(Path.GetTempPath(), "OpenXmlForVsto");
            if (!Directory.Exists(tmpDirPath)) Directory.CreateDirectory(tmpDirPath);
            return tmpDirPath;
        }

        static string GetNewRandomFilePath(string directory) =>
            Path.Combine(directory, Path.GetRandomFileName()) + ".docx";

        [Test]
        public void CopyToFile_And_CopyFromFile_Test_DoesNotChangeWhitespaces() {
            Range sourceRange = _application.Documents.Add().Range();
            const string expected = "test";
            sourceRange.Text = expected;

            _tmpFile = new OpenXmlHelper().CopyToFile(sourceRange);

            Range targetRange = _application.Documents.Add().Range();
            targetRange.Text = "unexpected";

            new OpenXmlHelper().CopyFromFile(_tmpFile, targetRange);

            Assert.That(targetRange.Document.Range().Text, Is.EqualTo(sourceRange.Document.Range().Text));
        }

        [Test]
        public void CopyToFileTextOnly_And_CopyFromFileTextOnly_Test_DoesNotChangeWhitespaces() {
            Range sourceRange = _application.Documents.Add().Range();
            sourceRange.Text = "test";

            _tmpFile = new OpenXmlHelper().CopyToFileTextOnly(sourceRange);

            Range targetRange = _application.Documents.Add().Range();
            targetRange.Text = "unexpected";

            new OpenXmlHelper().CopyFromFileTextOnly(_tmpFile, targetRange);

            Assert.That(targetRange.Document.Range().Text, Is.EqualTo(sourceRange.Document.Range().Text));
        }
    }
}