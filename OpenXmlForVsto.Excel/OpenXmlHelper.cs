﻿using Microsoft.Office.Interop.Excel;
using System;
using System.IO;
using System.Runtime.InteropServices;

namespace OpenXmlForVsto.Excel {
    public class OpenXmlHelper {
        /// <summary>
        /// Copy a range to a separate .xlsx file,
        /// to a sheet with same name as source sheet,
        /// to the same position as source range.
        /// It works exactly same as manual copy from one Excel file to a new file
        /// then saving and closing new file.
        /// Note: remove this temporary .xlsx file after using,
        /// because processes have limitation on amount of created temporary files.
        /// </summary>
        /// <param name="sourceRange">Single-area range that will be copied.</param>
        /// <returns>Full path to a new temporary .xlsx file with copied range.</returns>
        public string CopyToFile(Range sourceRange) {
            if (sourceRange == null) throw new ArgumentNullException(nameof(sourceRange));

            var target = SetupWorkbookAndSheet(sourceRange);

            Copy(sourceRange, GetTargetRange(sourceRange, target.Item2));

            return SaveAndClose(target.Item1);
        }

        /// <summary>
        /// Copy a range to a separate .xlsx file,
        /// to a sheet with same name as source sheet,
        /// to the same position as source range.
        /// It works exactly same as manual copy special from one Excel file to a new file
        /// then saving and closing new file.
        /// Note: remove this temporary .xlsx file after using,
        /// because processes have limitation on amount of created temporary files.
        /// </summary>
        /// <param name="sourceRange">Single-area range that will be copied.</param>
        /// <param name="pasteType">Mode of special copy. Can choose to copy all, only values, only styles etc.</param>
        /// <returns>Full path to a new temporary .xlsx file with copied range.</returns>
        public string CopyToFileSpecial(Range sourceRange, XlPasteType pasteType) {
            if (sourceRange == null) throw new ArgumentNullException(nameof(sourceRange));

            var target = SetupWorkbookAndSheet(sourceRange);

            CopySpecial(sourceRange, GetTargetRange(sourceRange, target.Item2), pasteType);

            return SaveAndClose(target.Item1);
        }

        /// <summary>
        /// Copy a range from matching position in a provided file to the current worksheet.
        /// It works exactly same as manual open of provided Excel file,
        /// copy same position range from it to the current worksheet
        /// closing provided file.
        /// </summary>
        /// <param name="sourceFile">Full path to .xlsx file.</param>
        /// <param name="targetRange">Single-area range in the current worksheet to copy to.</param>
        /// <param name="sheetName">Optional name of source sheet in the provided file.
        /// If missing, the method uses first sheet.</param>
        public void CopyFromFile(string sourceFile, Range targetRange, string sheetName = null) {
            if (sourceFile == null) throw new ArgumentNullException(nameof(sourceFile));
            if (!File.Exists(sourceFile)) throw new FileNotFoundException("File does not exist", sourceFile);
            if (targetRange == null) throw new ArgumentNullException(nameof(targetRange));

            Workbook sourceWorkbook = targetRange.Application.Workbooks.Open(sourceFile);
            sourceWorkbook.Windows[1].Visible = false;

            Range sourceRange = sourceWorkbook.Worksheets[sheetName ?? (object)1].Range[targetRange.Address];
            Copy(sourceRange, targetRange);
            sourceWorkbook.Close();
        }

        /// <summary>
        /// Copy a range from matching position in a provided file to the current worksheet.
        /// It works exactly same as manual open of provided Excel file,
        /// copy (special) same position range from it to the current worksheet
        /// closing provided file.
        /// </summary>
        /// <param name="sourceFile">Full path to .xlsx file.</param>
        /// <param name="targetRange">Single-area range in the current worksheet to copy to.</param>
        /// <param name="pasteType">Mode of special copy. Can choose to copy all, only values, only styles etc.</param>
        /// <param name="sheetName">Optional name of source sheet in the provided file.
        /// If missing, the method uses first sheet.</param>
        public void CopyFromFileSpecial(string sourceFile, Range targetRange, XlPasteType pasteType, string sheetName = null) {
            if (sourceFile == null) throw new ArgumentNullException(nameof(sourceFile));
            if (!File.Exists(sourceFile)) throw new FileNotFoundException("File does not exist", sourceFile);
            if (targetRange == null) throw new ArgumentNullException(nameof(targetRange));

            Workbook sourceWorkbook = targetRange.Application.Workbooks.Open(sourceFile);
            sourceWorkbook.Windows[1].Visible = false;

            Range sourceRange = sourceWorkbook.Worksheets[sheetName ?? (object)1].Range[targetRange.Address];

            CopySpecial(sourceRange, targetRange, pasteType);

            sourceWorkbook.Close();

            ClearClipboard();
        }

        static void Copy(Range source, Range target) =>
            source.Copy(target);

        static void CopySpecial(Range source, Range target, XlPasteType pasteType) {
            source.Copy();
            target.PasteSpecial(pasteType);
        }

        static Tuple<Workbook, Worksheet> SetupWorkbookAndSheet(Range sourceRange) {
            Application app = sourceRange.Application;
            Workbook workbook = app.Workbooks.Add();
            workbook.Windows[1].Visible = false;
            Worksheet sheet = workbook.Worksheets[1];
            sheet.Name = sourceRange.Worksheet.Name;
            return new Tuple<Workbook, Worksheet>(workbook, sheet);
        }

        static Range GetTargetRange(Range sourceRange, Worksheet targetSheet) =>
            targetSheet.Cells[sourceRange.Row, sourceRange.Column];

        static string GetOrCreateTmpDirectory() {
            string tmpDirPath = Path.Combine(Path.GetTempPath(), "OpenXmlForVsto");
            if (!Directory.Exists(tmpDirPath)) Directory.CreateDirectory(tmpDirPath);
            return tmpDirPath;
        }

        static string GetNewRandomFilePath(string directory) =>
            Path.Combine(directory, Path.GetRandomFileName()) + ".xlsx";

        static string SaveAndClose(Workbook workbook) {
            string tmpFilePath = GetNewRandomFilePath(GetOrCreateTmpDirectory());
            workbook.SaveAs(tmpFilePath, XlFileFormat.xlOpenXMLWorkbook);
            workbook.Close();
            return tmpFilePath;
        }

        static void ClearClipboard() {
            OpenClipboard(IntPtr.Zero);
            EmptyClipboard();
        }

        [DllImport("user32.dll")]
        static extern bool OpenClipboard(IntPtr hWndNewOwner);

        [DllImport("user32.dll")]
        static extern bool EmptyClipboard();
    }
}