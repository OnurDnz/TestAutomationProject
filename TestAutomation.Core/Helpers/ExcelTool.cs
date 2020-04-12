using Microsoft.Office.Interop.Excel;
using NUnit.Framework;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace TestAutomation.Core.Helpers
{
    public static class ExcelTool
    {
        private static Application xlApp = null;
        private static Workbooks workbooks = null;
        private static Workbook workbook = null;
        private static Range range;
        private static Hashtable sheets;
        public static string ExcelFilePath { get; set; }

        private static void OpenExcel()
        {
            xlApp = new Application();
            workbooks = xlApp.Workbooks;
            workbook = workbooks.Open(ExcelFilePath);
            sheets = new Hashtable();
            int count = 1;
            foreach (Worksheet sheet in workbook.Sheets)
            {
                sheets[count] = sheet.Name;
                count++;
            }
        }
        private static void CloseExcel()
        {
            workbook.Close(false, ExcelFilePath, null);
            Marshal.FinalReleaseComObject(workbook);
            workbook = null;

            workbooks.Close();
            Marshal.FinalReleaseComObject(workbooks);
            workbooks = null;

            xlApp.Quit();
            Marshal.FinalReleaseComObject(xlApp);
            xlApp = null;
        }
        public static void PaintToGreen(object cell1, object cell2)
        {
            OpenExcel();
            range = xlApp.get_Range(cell1, cell2);
            range.Interior.Color = XlRgbColor.rgbLightGreen;
            workbook.Save();
            CloseExcel();
        }
        public static void PaintToRed(object cell1, object cell2)
        {
            OpenExcel();
            range = xlApp.get_Range(cell1, cell2);
            range.Interior.Color = XlRgbColor.rgbRed;
            workbook.Save();
            CloseExcel();
        }
        public static string GetCellData(string sheetName, int colNumber, int rowNumber)
        {
            OpenExcel();

            string value = string.Empty;
            int sheetValue = 0;

            if (sheets.ContainsValue(sheetName))
            {
                foreach (DictionaryEntry sheet in sheets)
                {
                    if (sheet.Value.Equals(sheetName))
                    {
                        sheetValue = (int)sheet.Key;
                    }
                }
                Worksheet worksheet = null;
                worksheet = workbook.Worksheets[sheetValue] as Worksheet;
                Range range = worksheet.UsedRange;

                value = Convert.ToString((range.Cells[rowNumber, colNumber] as Range).Value2);
                Marshal.FinalReleaseComObject(worksheet);
                worksheet = null;
            }
            CloseExcel();
            return value;
        }
        public static bool SetCellData(string sheetName, string colName, int rowNumber, string value)
        {
            OpenExcel();

            int sheetValue = 0;
            int colNumber = 0;

            try
            {
                if (sheets.ContainsValue(sheetName))
                {
                    foreach (DictionaryEntry sheet in sheets)
                    {
                        if (sheet.Value.Equals(sheetName))
                        {
                            sheetValue = (int)sheet.Key;
                        }
                    }

                    Worksheet worksheet = null;
                    worksheet = workbook.Worksheets[sheetValue] as Worksheet;
                    Range range = worksheet.UsedRange;

                    for (int i = 1; i <= range.Columns.Count; i++)
                    {
                        string colNameValue = Convert.ToString((range.Cells[1, i] as Range).Value2);
                        if (colNameValue.ToLower() == colName.ToLower())
                        {
                            colNumber = i;
                            break;
                        }
                    }

                    range.Cells[rowNumber, colNumber] = value;
                    workbook.Save();
                    Marshal.FinalReleaseComObject(worksheet);
                    worksheet = null;

                    CloseExcel();
                }
            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }
        public static List<string> GetAll(int sheetİndex)
        {
            int rCnt, cCnt, rw = 0, cl = 0;
            List<string> cellDataList = new List<string>();

            OpenExcel();
            Worksheet worksheet = workbook.Worksheets.get_Item(sheetİndex);
            range = worksheet.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;

            for (rCnt = 1; rCnt <= rw; rCnt++)
            {
                for (cCnt = 1; cCnt <= cl; cCnt++)
                {
                    if (range.Cells[rCnt, cCnt].Value2 != null)
                    {
                        cellDataList.Add(range.Cells[rCnt, cCnt].Value2.ToString());
                    }
                }
            }

            CloseExcel();
            return cellDataList;
        }
        public static List<string> GetMultipleCells(int sheetİndex, string startCell, string endCell)
        {
            List<string> cellDataList = new List<string>();

            OpenExcel();
            Worksheet worksheet = workbook.Worksheets.get_Item(sheetİndex);

            range = worksheet.UsedRange;
            Range multipleCells = xlApp.get_Range(startCell, endCell);
            foreach (var item in multipleCells.Value2)
            {
                if (item != null)
                {
                    cellDataList.Add(item);
                }
            }
            CloseExcel();
            return cellDataList;
        }
        public static void WriteTestStatus()
        {
            if (TestContext.CurrentContext.Result.FailCount == 0)
            {
                MoveDownUntilToFindEmptyCell(2, 5, Status.Passed);
            }
            else
            {
                MoveDownUntilToFindEmptyCell(2, 5, Status.Failed);
            }

        }

        public static void MoveDownUntilToFindEmptyCell(int colNum, int rowName, Status status)
        {
            bool flag = true;
            while (flag)
            {
                if (GetCellData("DataSet", rowName, colNum) == null)
                {
                    SetCellData("DataSet", "TestName", colNum, $"{TestContext.CurrentContext.Test.MethodName}");
                    if (status == Status.Passed)
                    {
                        SetCellData("DataSet", "Result", colNum, status.ToString());
                        PaintToGreen("E" + colNum, "E" + colNum);
                    }
                    else
                    {
                        SetCellData("DataSet", "Result", colNum, status.ToString());
                        PaintToRed("E" + colNum, "E" + colNum);
                    }
                    flag = false;
                }
                colNum++;
            }
        }
        public enum Status
        {
            Passed,
            Failed
        }
    }
}