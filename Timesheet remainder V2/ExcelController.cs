using OfficeOpenXml;
using System;
using System.IO;

namespace Timesheet_remainder
{
    public class ExcelController
    {
        public void NewExcelFile(string fileLoadPath)
        {
            var fileInfoPath = new FileInfo(fileLoadPath);
            using (var excelPackage = new ExcelPackage(fileInfoPath))
            {
                var worksheet = excelPackage.Workbook.Worksheets.Add("MySheet");
                SetNewHeader(worksheet);

                excelPackage.Save();
            }
        }

        public void AddNewEntryToWorkSheet(string fileLoadPath, DateTime sheetDateTime, string TaskInputText)
        {
            var fileInfoPath = new FileInfo(fileLoadPath);
            using (var excelPackage = new ExcelPackage(fileInfoPath))
            {
                var worksheet = excelPackage.Workbook.Worksheets["MySheet"];
                var lastRow = worksheet.Dimension.End.Row;

                //Set the next last cell in the row to sheetDate.Text
                worksheet.Cells[lastRow + 1, 1].Value = sheetDateTime.ToString("dd/MM/yy");
                //Set the next last cell in the row to txtTaskInput.Text;
                worksheet.Cells[lastRow + 1, 2].Value = sheetDateTime.ToString("HH:mm:ss");
                worksheet.Cells[lastRow + 1, 3].Value = TaskInputText;

                excelPackage.Save();
            }
        }

        private void SetNewHeader(ExcelWorksheet worksheet)
        {
            worksheet.Cells["A1"].Value = "Date";
            worksheet.Cells["B1"].Value = "Time";
            worksheet.Cells["C1"].Value = "Task Description";
        }
    }
}
