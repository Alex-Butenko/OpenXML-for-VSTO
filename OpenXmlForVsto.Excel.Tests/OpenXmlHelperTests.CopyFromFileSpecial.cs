using ClosedXML.Excel;
using Microsoft.Office.Interop.Excel;
using NUnit.Framework;
using System;
using System.IO;
using System.Reflection;

namespace OpenXmlForVsto.Excel.Tests {
    [TestFixture]
    public partial class OpenXmlHelperTests {
        [Test]
        public void CopyFromFileSpecial_Test_ThrowsNREForNullFile() {
            Range range = _application.Workbooks.Add().Sheets[1].Cells[1, 1];
            Assert.Throws<ArgumentNullException>(() => new OpenXmlHelper().CopyFromFileSpecial(null, range, XlPasteType.xlPasteAll));

            (range.Worksheet.Parent as Workbook)?.Close();
        }

        [Test]
        public void CopyFromFileSpecial_Test_ThrowsNREForNullRange() {
            string anyExistingFileName = Assembly.GetAssembly(typeof(OpenXmlHelperTests)).Location;
            Assert.Throws<ArgumentNullException>(() => new OpenXmlHelper().CopyFromFileSpecial(anyExistingFileName, null, XlPasteType.xlPasteAll));
        }

        [Test]
        public void CopyFromFileSpecial_Test_ThrowsFileNotFoundForMissingFile() {
            string nonExistingFileName = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
            Assert.Throws<FileNotFoundException>(() => new OpenXmlHelper().CopyFromFileSpecial(nonExistingFileName, null, XlPasteType.xlPasteAll));
        }

        [Test]
        public void CopyFromFileSpecial_Test_CopiesProvidedSheet() {
            const int expected1 = 1;
            const string sheetName = "testSheetName";
            _tmpFile = GetNewRandomFilePath(GetOrCreateTmpDirectory());
            using (XLWorkbook workbook = new XLWorkbook()) {
                workbook.Worksheets.Add();
                IXLWorksheet sourceSheet = workbook.Worksheets.Add();
                sourceSheet.Name = sheetName;

                sourceSheet.Cell(1, 1).Value = expected1;
                workbook.SaveAs(_tmpFile);
            }

            Worksheet targetSheet = _application.Workbooks.Add().Sheets[1];
            Range range = targetSheet.Cells[1, 1];

            new OpenXmlHelper().CopyFromFileSpecial(_tmpFile, range, XlPasteType.xlPasteAll, sheetName);

            double result1 = (double)targetSheet.Cells[1, 1].Value;

            Assert.That(result1, Is.EqualTo(expected1).Within(0.0001));

            (targetSheet.Parent as Workbook)?.Close();
        }

        [Test]
        public void CopyFromFileSpecial_Test_CopiesCorrectPosition() {
            const int expected1 = 1;
            const int expected2 = 2;
            _tmpFile = GetNewRandomFilePath(GetOrCreateTmpDirectory());
            using (XLWorkbook workbook = new XLWorkbook()) {
                IXLWorksheet sourceSheet = workbook.Worksheets.Add();

                sourceSheet.Cell(11, 11).Value = expected1;
                sourceSheet.Cell(12, 12).Value = expected2;
                workbook.SaveAs(_tmpFile);
            }

            Worksheet targetSheet = _application.Workbooks.Add().Sheets[1];
            Range range = targetSheet.Range[targetSheet.Cells[11, 11], targetSheet.Cells[12, 12]];

            new OpenXmlHelper().CopyFromFileSpecial(_tmpFile, range, XlPasteType.xlPasteAll);

            object result1 = targetSheet.Cells[11, 11].Value;
            object result2 = targetSheet.Cells[12, 12].Value;

            Assert.That(result1, Is.EqualTo(expected1).Within(0.0001));
            Assert.That(result2, Is.EqualTo(expected2).Within(0.0001));

            (targetSheet.Parent as Workbook)?.Close();
        }

        [Test]
        public void CopyFromFileSpecial_Test_CopiesCorrectData() {
            const int expected1 = 1;
            const string expected2 = "test";
            DateTime expected3 = DateTime.Now;
            _tmpFile = GetNewRandomFilePath(GetOrCreateTmpDirectory());
            using (XLWorkbook workbook = new XLWorkbook()) {
                IXLWorksheet sourceSheet = workbook.Worksheets.Add();

                sourceSheet.Cell(1, 1).Value = expected1;
                sourceSheet.Cell(2, 2).Value = expected2;
                sourceSheet.Cell(1, 2).Value = expected3;
                workbook.SaveAs(_tmpFile);
            }

            Worksheet targetSheet = _application.Workbooks.Add().Sheets[1];
            Range range = targetSheet.Range[targetSheet.Cells[1, 1], targetSheet.Cells[2, 2]];

            new OpenXmlHelper().CopyFromFileSpecial(_tmpFile, range, XlPasteType.xlPasteAll);

            double result1 = (double)targetSheet.Cells[1, 1].Value;
            string result2 = (string)targetSheet.Cells[2, 2].Value;
            DateTime result3 = (DateTime)targetSheet.Cells[1, 2].Value;

            Assert.That(result1, Is.EqualTo(expected1).Within(0.0001));
            Assert.That(result2, Is.EqualTo(expected2));
            Assert.That(result3, Is.EqualTo(expected3).Within(1).Seconds);

            (targetSheet.Parent as Workbook)?.Close();
        }

        [Test]
        public void CopyFromFileSpecial_Test_CopiesCorrectStyle() {
            _tmpFile = GetNewRandomFilePath(GetOrCreateTmpDirectory());
            const string numberFormat = "0%";
            using (XLWorkbook workbook = new XLWorkbook()) {
                IXLWorksheet sourceSheet = workbook.Worksheets.Add();

                IXLStyle style = sourceSheet.Cell(1, 1).Style;
                style.Font.Bold = true;
                style.Font.Italic = true;
                style.NumberFormat.Format = numberFormat;
                workbook.SaveAs(_tmpFile);
            }

            Worksheet targetSheet = _application.Workbooks.Add().Sheets[1];
            Range range = targetSheet.Cells[1, 1];

            new OpenXmlHelper().CopyFromFileSpecial(_tmpFile, range, XlPasteType.xlPasteAll);

            Assert.That(range.Font.Bold, Is.True);
            Assert.That(range.Font.Italic, Is.True);
            Assert.That(range.NumberFormat, Is.EqualTo(numberFormat));

            (targetSheet.Parent as Workbook)?.Close();
        }

        [Test]
        public void CopyFromFileSpecial_Test_CopiesStyleOnly() {
            _tmpFile = GetNewRandomFilePath(GetOrCreateTmpDirectory());
            const string numberFormat = "0%";
            const int expected1 = 1;
            const string expected2 = "test";
            DateTime expected3 = DateTime.Now;
            using (XLWorkbook workbook = new XLWorkbook()) {
                IXLWorksheet sourceSheet = workbook.Worksheets.Add();

                IXLStyle style = sourceSheet.Cell(1, 1).Style;
                style.Font.Bold = true;
                style.Font.Italic = true;
                style.NumberFormat.Format = numberFormat;

                sourceSheet.Cell(1, 1).Value = expected1;
                sourceSheet.Cell(2, 2).Value = expected2;
                sourceSheet.Cell(1, 2).Value = expected3;

                workbook.SaveAs(_tmpFile);
            }

            Worksheet targetSheet = _application.Workbooks.Add().Sheets[1];
            Range range = targetSheet.Cells[1, 1];

            new OpenXmlHelper().CopyFromFileSpecial(_tmpFile, range, XlPasteType.xlPasteFormats);

            Assert.That(range.Font.Bold, Is.True);
            Assert.That(range.Font.Italic, Is.True);
            Assert.That(range.NumberFormat, Is.EqualTo(numberFormat));

            object result1 = targetSheet.Cells[1, 1].Value;
            object result2 = targetSheet.Cells[2, 2].Value;
            object result3 = targetSheet.Cells[1, 2].Value;

            Assert.That(result1, Is.Null);
            Assert.That(result2, Is.Null);
            Assert.That(result3, Is.Null);

            (targetSheet.Parent as Workbook)?.Close();
        }

        [Test]
        public void CopyFromFileSpecial_Test_CopiesCorrectFormulas() {
            _tmpFile = GetNewRandomFilePath(GetOrCreateTmpDirectory());
            using (XLWorkbook workbook = new XLWorkbook()) {
                IXLWorksheet sourceSheet = workbook.Worksheets.Add();

                sourceSheet.Cell(1, 1).Value = 1;
                sourceSheet.Cell(1, 2).FormulaA1 = "A1";
                workbook.SaveAs(_tmpFile);
            }

            Worksheet targetSheet = _application.Workbooks.Add().Sheets[1];
            Range range = targetSheet.Range[targetSheet.Cells[1, 1], targetSheet.Cells[1, 2]];

            new OpenXmlHelper().CopyFromFileSpecial(_tmpFile, range, XlPasteType.xlPasteFormulas);

            string result1 = (string)range.Cells[1, 2].Formula;

            Assert.That(result1, Is.EqualTo("=A1"));

            (targetSheet.Parent as Workbook)?.Close();
        }

        [Test]
        public void CopyFromFileSpecial_Test_CopiesCorrectFormulas_WithMissingSource() {
            _tmpFile = GetNewRandomFilePath(GetOrCreateTmpDirectory());
            using (XLWorkbook workbook = new XLWorkbook()) {
                IXLWorksheet sourceSheet = workbook.Worksheets.Add();

                sourceSheet.Cell(1, 1).Value = 1;
                sourceSheet.Cell(1, 2).FormulaA1 = "A1";
                workbook.SaveAs(_tmpFile);
            }

            Worksheet targetSheet = _application.Workbooks.Add().Sheets[1];
            Range range = targetSheet.Cells[1, 2];

            new OpenXmlHelper().CopyFromFileSpecial(_tmpFile, range, XlPasteType.xlPasteFormulas);

            string result1 = (string)range.Formula;

            Assert.That(result1, Is.EqualTo("=A1"));

            (targetSheet.Parent as Workbook)?.Close();
        }

        [Test]
        public void CopyFromFileSpecial_Test_ClipboardIsEmpty() {
            _tmpFile = GetNewRandomFilePath(GetOrCreateTmpDirectory());
            using (XLWorkbook workbook = new XLWorkbook()) {
                IXLWorksheet sourceSheet = workbook.Worksheets.Add();

                sourceSheet.Cell(1, 1).Value = 1;
                workbook.SaveAs(_tmpFile);
            }

            Worksheet targetSheet = _application.Workbooks.Add().Sheets[1];
            Range range = targetSheet.Cells[1, 1];

            new OpenXmlHelper().CopyFromFileSpecial(_tmpFile, range, XlPasteType.xlPasteAll);

            Assert.True(IsClipboardEmpty());
        }
    }
}