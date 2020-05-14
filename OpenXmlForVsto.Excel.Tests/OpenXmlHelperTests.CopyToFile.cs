using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Microsoft.Office.Interop.Excel;
using NUnit.Framework;
using System;
using System.Collections.Generic;

namespace OpenXmlForVsto.Excel.Tests {
    [TestFixture]
    public partial class OpenXmlHelperTests {
        [Test]
        public void CopyToFile_Test_CreatesFileForValidRange() {
            Range range = _application.Workbooks.Add().Sheets[1].Cells[1, 1];

            _tmpFile = new OpenXmlHelper().CopyToFile(range);

            FileAssert.Exists(_tmpFile);
        }

        [Test]
        public void CopyToFile_Test_ThrowsNREForNullRange() {
            Assert.Throws<ArgumentNullException>(() => new OpenXmlHelper().CopyToFile(null));
        }

        [Test]
        public void CopyToFile_Test_CreatesValidFile() {
            Range range = _application.Workbooks.Add().Sheets[1].Cells[1, 1];

            _tmpFile = new OpenXmlHelper().CopyToFile(range);

            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(_tmpFile, true)) {
                IEnumerable<ValidationErrorInfo> errors = new OpenXmlValidator().Validate(doc);

                CollectionAssert.IsEmpty(errors);
            }
        }

        [Test]
        public void CopyToFile_Test_CreatesCorrectSheetName() {
            Range range = _application.Workbooks.Add().Sheets[1].Cells[1, 1];
            const string expected = "testSheetName";
            range.Worksheet.Name = expected;

            _tmpFile = new OpenXmlHelper().CopyToFile(range);

            using (XLWorkbook workbook = new XLWorkbook(_tmpFile)) {
                string result = workbook.Worksheets.Worksheet(1).Name;

                Assert.That(result, Is.EqualTo(expected));
            }
        }

        [Test]
        public void CopyToFile_Test_CopiesCorrectPosition() {
            Worksheet sheet = _application.Workbooks.Add().Sheets[1];
            Range range = sheet.Range[sheet.Cells[11, 11], sheet.Cells[12, 12]];
            const int expected1 = 1;
            const int expected2 = 2;
            sheet.Cells[11, 11].Value = expected1;
            sheet.Cells[12, 12].Value = expected2;

            _tmpFile = new OpenXmlHelper().CopyToFile(range);

            using (XLWorkbook workbook = new XLWorkbook(_tmpFile)) {
                IXLWorksheet resultSheet = workbook.Worksheets.Worksheet(1);

                double result1 = (double)resultSheet.Cell(11, 11).Value;
                double result2 = (double)resultSheet.Cell(12, 12).Value;

                Assert.That(result1, Is.EqualTo(expected1).Within(0.0001));
                Assert.That(result2, Is.EqualTo(expected2).Within(0.0001));
            }
        }

        [Test]
        public void CopyToFile_Test_CopiesCorrectData() {
            Worksheet sheet = _application.Workbooks.Add().Sheets[1];
            Range range = sheet.Range[sheet.Cells[1, 1], sheet.Cells[2, 2]];
            const int expected1 = 1;
            const string expected2 = "test";
            DateTime expected3 = DateTime.Now;
            sheet.Cells[1, 1].Value = expected1;
            sheet.Cells[2, 2].Value = expected2;
            sheet.Cells[1, 2].Value = expected3;

            _tmpFile = new OpenXmlHelper().CopyToFile(range);

            using (XLWorkbook workbook = new XLWorkbook(_tmpFile)) {
                IXLWorksheet resultSheet = workbook.Worksheets.Worksheet(1);

                double result1 = (double)resultSheet.Cell(1, 1).Value;
                string result2 = (string)resultSheet.Cell(2, 2).Value;
                DateTime result3 = (DateTime)resultSheet.Cell(1, 2).Value;

                Assert.That(result1, Is.EqualTo(expected1).Within(0.0001));
                Assert.That(result2, Is.EqualTo(expected2));
                Assert.That(result3, Is.EqualTo(expected3).Within(1).Seconds);
            }
        }

        [Test]
        public void CopyToFile_Test_CopiesCorrectStyle() {
            Worksheet sheet = _application.Workbooks.Add().Sheets[1];
            Range range = sheet.Cells[1, 1];

            range.Font.Bold = true;
            range.Font.Italic = true;
            range.Font.Color = System.Drawing.Color.Coral;
            range.Interior.Color = System.Drawing.Color.BlueViolet;
            range.Borders.LineStyle = XlLineStyle.xlDash;

            _tmpFile = new OpenXmlHelper().CopyToFile(range);

            using (XLWorkbook workbook = new XLWorkbook(_tmpFile)) {
                IXLWorksheet resultSheet = workbook.Worksheets.Worksheet(1);

                IXLCell cell = resultSheet.Cell(1, 1);

                Assert.That(cell.Style.Font.Bold, Is.True);
                Assert.That(cell.Style.Font.Italic, Is.True);
                Assert.That(cell.Style.Font.FontColor.Color, Is.EqualTo(System.Drawing.Color.Coral));
                Assert.That(cell.Style.Fill.BackgroundColor.Color, Is.EqualTo(System.Drawing.Color.BlueViolet));
                Assert.That(cell.Style.Border.BottomBorder, Is.EqualTo(XLBorderStyleValues.Dashed));
            }
        }

        [Test]
        public void CopyToFile_Test_CopiesCorrectFormulas() {
            Worksheet sheet = _application.Workbooks.Add().Sheets[1];
            Range range = sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, 2]];
            const string expected = "A1";
            sheet.Cells[1, 1].Value = 1;
            sheet.Cells[1, 2].Formula = "=" + expected;

            _tmpFile = new OpenXmlHelper().CopyToFile(range);

            using (XLWorkbook workbook = new XLWorkbook(_tmpFile)) {
                IXLWorksheet resultSheet = workbook.Worksheets.Worksheet(1);

                string result = resultSheet.Cell(1, 2).FormulaA1;

                Assert.That(result, Is.EqualTo(expected));
            }
        }

        [Test]
        public void CopyToFile_Test_CopiesCorrectFormulas_WithMissingSource() {
            Worksheet sheet = _application.Workbooks.Add().Sheets[1];
            Range range = sheet.Cells[1, 2];
            const string expected = "A1";
            sheet.Cells[1, 1].Value = 1;
            sheet.Cells[1, 2].Formula = "=" + expected;

            _tmpFile = new OpenXmlHelper().CopyToFile(range);

            using (XLWorkbook workbook = new XLWorkbook(_tmpFile)) {
                IXLWorksheet resultSheet = workbook.Worksheets.Worksheet(1);

                string result = resultSheet.Cell(1, 2).FormulaA1;
                string countedValue = (string)resultSheet.Cell(1, 2).Value;
                string sourceValue = (string)resultSheet.Cell(1, 1).Value;

                Assert.That(result, Is.EqualTo(expected));
                Assert.That(countedValue, Is.EqualTo("0"));
                Assert.That(sourceValue, Is.EqualTo(string.Empty));
            }
        }

        [Test]
        public void CopyToFile_Test_ClipboardIsEmpty() {
            Worksheet sheet = _application.Workbooks.Add().Sheets[1];
            Range range = sheet.Cells[1, 2];
            sheet.Cells[1, 1].Value = 1;

            _tmpFile = new OpenXmlHelper().CopyToFile(range);

            Assert.True(IsClipboardEmpty());
        }
    }
}