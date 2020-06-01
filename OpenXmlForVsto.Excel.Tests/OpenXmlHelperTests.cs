using Microsoft.Office.Interop.Excel;
using NUnit.Framework;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;

namespace OpenXmlForVsto.Excel.Tests {
    [TestFixture]
    public partial class OpenXmlHelperTests {
        [SetUp]
        public void Setup() {
            _application = new Application { DisplayAlerts = false };
        }

        [TearDown]
        public void TearDown() {
            int processId = 0;
            try {
                int hWnd = _application.Application.Hwnd;
                GetWindowThreadProcessId((IntPtr)hWnd, out processId);

                _application.DisplayAlerts = false;
                _application.Quit();
            }
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
            Path.Combine(directory, Path.GetRandomFileName()) + ".xlsx";
    }
}