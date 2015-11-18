using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Excel;

namespace NationalitetsParser
{
   class ExcelWrapper
   {
      private readonly Sheets mWorksheets;
      private List<string> mColumnHeaders;
      private Worksheet mCurrentWorksheet;

      public ExcelWrapper(string pPath)
      {
         Application app = new Application();

         Workbook wb = app.Workbooks.Open(pPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                 Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                 Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                 Type.Missing, Type.Missing);

         mWorksheets = wb.Sheets;
         mColumnHeaders = new List<string>();
      }

      public ExcelWrapper()
      {
         Application xlApp = new Application();

         if (xlApp == null)
         {
            Console.WriteLine("EXCEL could not be started. Check that your office installation and project references are correct.");
            return;
         }
         xlApp.Visible = true;
         Workbook wb = xlApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
         Worksheet ws = (Worksheet)wb.Worksheets[1];

         if (ws == null)
         {
            Console.WriteLine("Worksheet could not be created. Check that your office installation and project references are correct.");
         }

         mCurrentWorksheet = ws;
         mColumnHeaders = new List<string>();
      }

      public void SetActiveWorksheetByFirstRow(Func<Range, bool> pSelectFunc)
      {
         foreach (Worksheet worksheet in mWorksheets)
         {
            Range firstRow = worksheet.Rows[0];
            if (pSelectFunc(firstRow))
            {
               mCurrentWorksheet = worksheet;
               UpdateColumnHeaders();
               return;
            }
         }
      }

      public void SetActiveWorksheetByName(string pName)
      {
         mCurrentWorksheet = mWorksheets[pName];
         UpdateColumnHeaders();
      }

      private void UpdateColumnHeaders()
      {
         mColumnHeaders.Clear();
         var firstRow = mCurrentWorksheet.Rows[1];
         foreach (var cell in firstRow.Cells)
            if (cell.Text == "")
               break;
            else
               mColumnHeaders.Add(cell.Text);
      }

      public int GetColumnIndex(string pColumnName)
      {
         for (int i = mColumnHeaders.Count(); i > 0; i--)
         {
            if (mColumnHeaders[i - 1].Equals(pColumnName))
               return i;
         }
         return -1;
      }

      public Range GetRows()
      {
         return mCurrentWorksheet.Rows;
      }

      public static ExcelWrapper CreateNewWorksheet(string[] pHeaderStrings = null)
      {
         ExcelWrapper wrapper = new ExcelWrapper();
         if (pHeaderStrings == null)
            return wrapper;
         for (int i = pHeaderStrings.Count() - 1; i >= 0; i--)
         {
            wrapper.SetCellsByRange(NumberToExcelIndex(i + 1) + "1", pHeaderStrings[i]);
         }
         return wrapper;
      }

      public static String NumberToExcelIndex(int number)
      {
         string ret = "";
         while (--number >= 0)
         {
            ret = (Char)('A' + (number % 26)) + ret;
            number /= 26;
         }
         return ret;
      }

      public void SetCellsByRange(string pRange, string pValue)
      {
         Range aRange = mCurrentWorksheet.Range[pRange];
         aRange.Value2 = pValue;
         if (aRange.Row == 1)
            UpdateColumnHeaders();
      }

      public void SetCellsByColumnHeader(Range pRow1, string pHeader, string pCell)
      {
         int colIndex = GetColumnIndex(pHeader);
         if (colIndex < 0)
            return;
         Range cell = pRow1.Cells[1, colIndex];
         cell.NumberFormat = "@";
         if (pCell != null) cell.Value2 = pCell;
      }

      public string GetTextByColumnHeader(Range pRow1, string pHeader)
      {
         int colIndex = GetColumnIndex(pHeader);
         if (colIndex < 0)
            return null;
         object value2 = ((Range) pRow1.Cells[1, colIndex]).Value2;
         if (value2 == null)
            return null;
         return value2.ToString();
      }

      public float GetFloatByColumnHeader(Range pRow1, string pHeader)
      {
         int colIndex = GetColumnIndex(pHeader);
         if (colIndex < 0)
            return -1;
         Range cell = pRow1.Cells[1, colIndex];
         string value2 = cell.Value2;
         if (value2 == null)
            return -1;
         float ret = -1;
         try
         {
            ret = float.Parse(value2);
            return ret;
         }
         catch (Exception e)
         {
            // ignored
         }
         return ret;
      }

      public int GetIntByColumnHeader(Range pRow1, string pHeader)
      {
         int colIndex = GetColumnIndex(pHeader);
         if (colIndex < 0)
            return -1;
         Range cell = pRow1.Cells[1, colIndex];
         string value2 = cell.Value2;
         if (value2 == null)
            return -1;
         int ret = -1;
         try
         {
            ret = int.Parse(value2);
            return ret;
         }
         catch (Exception e)
         {
            // ignored
         }
         return ret;
      }

      public void SetCellsByColumnHeader(Range pRow1, string pHeader, float pCell)
      {
         int colIndex = GetColumnIndex(pHeader);
         if (colIndex < 0)
            return;
         Range cell = pRow1.Cells[1, colIndex];
         cell.NumberFormat = "0.00";
         if (pCell != null) cell.Value2 = pCell;
      }
   }
}
