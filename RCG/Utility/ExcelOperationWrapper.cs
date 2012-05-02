using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;

namespace RCG
{
    public class ExcelOperationWrapper
    {
        public static void ClearExcelSheetWithoutHeader(dynamic activeSheet, int actualExcelRowCountWithoutHeader)
        {
            //int rowCount = GetAvailableExcelRowCountWithoutHeader();
            if (actualExcelRowCountWithoutHeader > 0)
            {
                Excel.Range range = activeSheet.Range(string.Format("2:{0}", actualExcelRowCountWithoutHeader + 1));
                range.Clear();
            }
        }

        public static void ClearExcelSheetFormatWithoutHeader(dynamic activeSheet, int actualExcelRowCountWithoutHeader)
        {
            //int rowCount = GetAvailableExcelRowCountWithoutHeader();
            if (actualExcelRowCountWithoutHeader > 0)
            {
                Excel.Range range = activeSheet.Range(string.Format("2:{0}", actualExcelRowCountWithoutHeader + 1));
                range.FormatConditions.Delete();
            }
        }

        public static void ClearRowFormats(dynamic sheet, int rowIndex)
        {
            Excel.Range range = sheet.Range(string.Format("{0}:{0}", rowIndex));
            range.ClearFormats();
        }

        public static void SetRowBackgroundColor(dynamic sheet, int rowIndex, Color color)
        {
            Excel.Range range = sheet.Range(string.Format("{0}:{0}", rowIndex));
            
            dynamic fd;
            if (range.FormatConditions.Count > 0)
                fd = range.FormatConditions[1];
            else
                fd =
               (Excel.FormatCondition)range.FormatConditions.Add(Excel.XlFormatConditionType.xlExpression,
               Excel.XlFormatConditionOperator.xlEqual, true,
               Type.Missing);
            fd.Interior.PatternColorIndex = Excel.Constants.xlAutomatic;
            //fd.Interior.TintAndShade = 0;
            fd.Interior.Color = System.Drawing.ColorTranslator.ToWin32(color);
            //fd.StopIfTrue = false;
        }

        public static void SetRowForegroundColor(dynamic sheet, int rowIndex, Color color)
        {
            Excel.Range range = sheet.Range(string.Format("{0}:{0}", rowIndex));
            dynamic fd;
            if (range.FormatConditions.Count > 0)
                fd = range.FormatConditions[1];
            else
                fd =
               (Excel.FormatCondition)range.FormatConditions.Add(Excel.XlFormatConditionType.xlExpression,
               Excel.XlFormatConditionOperator.xlEqual, true,
               Type.Missing);
            fd.Font.Color = System.Drawing.ColorTranslator.ToWin32(color);
            //fd.StopIfTrue = false;
        }

        public static void SetRowFontBold(dynamic sheet, int rowIndex, bool flag)
        {
            Excel.Range range = sheet.Range(string.Format("{0}:{0}", rowIndex));
            dynamic fd;
            if (range.FormatConditions.Count > 0)
                fd = range.FormatConditions[1];
            else
                fd =
               (Excel.FormatCondition)range.FormatConditions.Add(Excel.XlFormatConditionType.xlExpression,
               Excel.XlFormatConditionOperator.xlEqual, true,
               Type.Missing);
            fd.Font.Bold = flag;
            //fd.StopIfTrue = false;
        }

        public static void SetRowFontItalic(dynamic sheet, int rowIndex, bool flag)
        {
            Excel.Range range = sheet.Range(string.Format("{0}:{0}", rowIndex));
            dynamic fd;
            if (range.FormatConditions.Count > 0)
                fd = range.FormatConditions[1];
            else
                fd =
               (Excel.FormatCondition)range.FormatConditions.Add(Excel.XlFormatConditionType.xlExpression,
               Excel.XlFormatConditionOperator.xlEqual, true,
               Type.Missing);
            fd.Font.Italic = flag;
            //fd.StopIfTrue = false;
        }

        public static dynamic FindExcelActiveSheet(Excel.Application excel, string sheetName)
        {
            dynamic activeSheet = FindExcelSheet(excel, sheetName);
            if (activeSheet == null)
            {
                activeSheet = excel.Application.Sheets.Add();
                activeSheet.Name = sheetName;
            }

            return activeSheet;
        }

        public static object FindExcelSheet(Excel.Application excelApp, string name)
        {
            foreach (dynamic d in excelApp.Application.Sheets)
            {
                if (d.Name.ToString().Trim() == name.Trim())
                    return d;
            }
            return null;
        }
    }
}
