using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;


namespace ConsoleAppTFS
{
    public class Excel2
    {
        private int row = 2;
        private string filename;
        private Microsoft.Office.Interop.Excel.Application excelApplication = new Microsoft.Office.Interop.Excel.Application();
        //excelApplication.Visible = true;
        //dynamic excelWorkBook = excelApplication.Workbooks.Add();
        private Microsoft.Office.Interop.Excel.Workbook excelWorkBook = null;
        public Excel2(String titolo)
        {

            excelWorkBook = excelApplication.Workbooks.Add();
            excelApplication.DisplayAlerts = false;
            filename = @"C:\report\" + titolo + "_" + DateTime.Today.Month + "-" + DateTime.Today.Year + ".xls";
            excelApplication.Cells[1, 1] = "Applicazione";
            excelApplication.Cells[1, 2] = "Data";
            excelApplication.Cells[1, 3] = "Linguaggio";
            excelApplication.Cells[1, 4] = "Files";
            excelApplication.Cells[1, 5] = "Righe";
            excelApplication.Cells[1, 6] = "Utenti";
            ((Microsoft.Office.Interop.Excel.Range)excelApplication.Cells[1, 1]).EntireRow.Font.Bold = true;
            ((Microsoft.Office.Interop.Excel.Range)excelApplication.Cells[1, 1]).EntireColumn.AutoFit();
            ((Microsoft.Office.Interop.Excel.Range)excelApplication.Cells[1, 2]).ColumnWidth = 10;
            ((Microsoft.Office.Interop.Excel.Range)excelApplication.Cells[1, 3]).EntireColumn.AutoFit();
            ((Microsoft.Office.Interop.Excel.Range)excelApplication.Cells[1, 4]).EntireColumn.AutoFit();
            ((Microsoft.Office.Interop.Excel.Range)excelApplication.Cells[1, 5]).EntireColumn.AutoFit();
            ((Microsoft.Office.Interop.Excel.Range)excelApplication.Cells[1, 6]).EntireColumn.AutoFit();
            excelWorkBook.SaveAs(filename, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal,
                System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, false,
                Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlShared, false, false, System.Reflection.Missing.Value,
                System.Reflection.Missing.Value, System.Reflection.Missing.Value);
        }
        public void inserisciRiga(List<string> l)
        {

            //dynamic excelWorkBook = excelApplication.Workbooks.Add();
            //Excel.Worksheet wkSheetData = excelWorkBook.ActiveSheet;
            int conto = 0;
            if (l.Count != 0)
            {
                foreach (string p in l)
                {
                    conto++;
                    ((Microsoft.Office.Interop.Excel.Range)excelApplication.Cells[row, conto]).Value = " " +p;
                    ((Microsoft.Office.Interop.Excel.Range)excelApplication.Cells[row, conto]).EntireColumn.AutoFit();
                }
                row++;
            }

            // funziona.

            excelWorkBook.Save();

        }
        public void close()
        {
            excelWorkBook.Close(Missing.Value, Missing.Value, Missing.Value);
        }
    }
}