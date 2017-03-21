using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace ReadXlsx
{
    public class CMultiArray<T>
    {
        private T[][] _array;

        public T this[int index, int index2 = 0]
        {
            get { return _array[index][index2]; }
            set { _array[index][index2] = value; }
        }
    }
    class Program
    {
        private static Excel.Application excelApp = null;
        private static Excel.Workbooks excelWorkBooks = null;
        private static Excel.Workbook excelWorkBook = null;
        private static Excel.Sheets excelSheets = null;
        private static Excel.Worksheet excelWorkSheet = null;
        private static Excel.Range excelRange = null;
        private static string sFile = "";

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        private static void Main(string[] args)
        {
            System.Diagnostics.Process[] PROC = Process.GetProcessesByName("EXCEL");
            foreach (System.Diagnostics.Process PK in PROC)
            {//User excel process always have window name
             //COM process do not.
                if (PK.MainWindowTitle.Length == 0)
                    PK.Kill();
            }
            //new Program().RoepAan();
            if (args != null && args.Length > 0 && args[0] != null && args[0] != "" && File.Exists(args[0]))
            {
                sFile = args[0];
            }
            else
            {
                sFile = @"C:\Users\Sytse\Desktop\Map1.xlsx";
            }
            try
            {
                
            }
            catch (Exception)
            {
                
                throw;
            }

            /*var cult = new CultureInfo("en-US");
            cult.DateTimeFormat.DateSeparator = "-";
            cult.DateTimeFormat.ShortDatePattern = "dd-MM-yyyy";
            Thread.CurrentThread.CurrentCulture = cult;*/
            /*unchecked
            {
                uint g = uint.MaxValue + 10;
            }*/
            

            OpenExcel();
            //PrintExcel();
            //CloseExcel();
            Console.ReadKey();
        }

        private static void Foo(int getal)
        {
            
        }

        private static void OpenExcel()
        {
            excelApp = new Excel.Application();
            int hwnd = excelApp.Hwnd;
            excelWorkBooks = excelApp.Workbooks;
            excelWorkBook = excelWorkBooks.Open( sFile );
            excelApp.DisplayAlerts = false;
            excelSheets = excelApp.Worksheets;
            excelWorkSheet = excelSheets["Blad1"];
            //excelWorkSheet.Unprotect( "test" );

            Func<string, string> getCharsFromString = ((inputStr) =>
            {
                return inputStr.Where(c => "ABCDEFGHIJKLMNOPQRSTUVWXYZ".Contains(c)).Aggregate("", (current, c) => current + c);
            });
            Func<string, int> getNumbersFromString = ((inputStr) =>
            {
                return int.Parse(inputStr.Where(c => "0123456789".Contains(c)).Aggregate("", (current, c) => current + c));
                //string sRet = inputStr.Where(c => "0123456789".Contains(c)).Aggregate("", (current, c) => current + c);
                //while (sRet.Length < 10) sRet = "0" + sRet;
                //return sRet;
            });

            var sortedList = new SortedDictionary<string,string>();
            for (int i = 0; i < 5; i++)
            {
                for (int j = 1; j < 1001; j++)
                {
                    switch (i)
                    {
                        case 0:
                            sortedList.Add("A" + j, "A" + j);
                            break;
                        case 1:
                            sortedList.Add("B" + j, "A" + j);
                            break;
                        case 2:
                            sortedList.Add("C" + j, "A" + j);
                            break;
                        case 3:
                            sortedList.Add("D" + j, "A" + j);
                            break;
                        case 4:
                            sortedList.Add("E" + j, "A" + j);
                            break;
                    }
                }
            }

            var sortedList2 = new SortedDictionary<int,int>();

            var res = from keyValPair in sortedList2
                select keyValPair;




            var result =    from keyValPair in sortedList
                            let key = keyValPair.Key
                            let columnName = getCharsFromString(key)
                            orderby columnName.Length,
                                    columnName,
                                    getNumbersFromString(key)
                            select keyValPair;






            var a = result.ToArray().ToDictionary(d => d.Key, d => d.Value);

            Excel.Range usedlast = excelWorkSheet.Range["A1", "AA20"];

            object[,] inputStrings = usedlast.Value2 as object[,];
            Parallel.For(1, 21, (i) =>
            {
                Parallel.For(1, 28, (j) =>
                {
                    inputStrings[i, j] = i + "," + j;
                });
            });
            usedlast.Value2 = inputStrings;
            Marshal.FinalReleaseComObject(usedlast);
            excelWorkBook.SaveAs(sFile);
            excelWorkBook.Close(false);
            excelApp.Quit();

            if (excelRange != null) Marshal.ReleaseComObject(excelRange);
            if (excelWorkSheet != null) Marshal.ReleaseComObject(excelWorkSheet);
            if (excelSheets != null) Marshal.ReleaseComObject(excelSheets);
            if (excelWorkBook != null) Marshal.ReleaseComObject(excelWorkBook);
            if (excelWorkBooks != null) Marshal.ReleaseComObject(excelWorkBooks);
            if (excelApp != null) Marshal.ReleaseComObject(excelApp);
            excelRange = null;
            excelWorkSheet = null;
            excelSheets = null;
            excelWorkBook = null;
            excelWorkBooks = null;
            excelApp = null;

            /*Thread.CurrentThread.CurrentUICulture = new CultureInfo
            (
                excelApp.LanguageSettings.LanguageID[MsoAppLanguageID.msoLanguageIDUI]
            );*/
        }

        private static void PrintExcel()
        {
            var langSet = excelApp.LanguageSettings;
            var lang = langSet.LanguageID[MsoAppLanguageID.msoLanguageIDUI];
            var cult = CultureInfo.GetCultureInfo(lang);
            var name = cult.Name;

            Excel.Range usedlast = excelWorkSheet.Cells;
            Excel.Range last = usedlast.Find("*", Missing.Value, Missing.Value, Missing.Value, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious, false, Missing.Value, Missing.Value);
            var iMaxRow = last.Row;

            // Find the last real column
            //nInLastCol = oSheet.Cells.Find("*", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;

            /*Excel.Range usedlast = excelWorkSheet.UsedRange;
            Excel.Range last = usedlast.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            var iets = last.NumberFormat;
            var iets3 = last.NumberFormatLocal;
            var lastVal = last.Value;
            
            string tmp = last?.Value ?? "";
            
            if (lastVal is DateTime )
            {
                string v = "";
            }*/
            
            //int iMaxRow = last.Row;
            //string sColumn = GetExcelColumnName(last.Column);

            Marshal.FinalReleaseComObject(usedlast);
            //Marshal.FinalReleaseComObject(last);

            int iCurrRow = 0;//12;
            while (iCurrRow < iMaxRow)
            {
                iCurrRow++;
                excelRange = excelWorkSheet.Range["A"+iCurrRow, "O"+iCurrRow];
                Array arr = excelRange.Value;

                Marshal.ReleaseComObject(excelRange);
                
                bool bFirst = true;
                foreach( var val in arr )
                {
                    if (!bFirst)
                    {
                        Console.Write(" - ");
                        Console.Write( $"{val}" );
                    }
                    else
                    {
                        Console.Write( $"{val}" );
                        bFirst = false;
                    }
                }

                Console.Write( "\n" );
            }
            Marshal.ReleaseComObject(excelRange);
        }

        private static void CloseExcel()
        {
            excelWorkSheet.Protect( "test" );
            excelWorkBook.Close( false );
            excelApp.Quit();

            if (excelRange != null) Marshal.ReleaseComObject(excelRange);
            if( excelWorkSheet != null ) Marshal.ReleaseComObject( excelWorkSheet );
            if( excelSheets != null ) Marshal.ReleaseComObject( excelSheets );
            if( excelWorkBook != null ) Marshal.ReleaseComObject( excelWorkBook );
            if( excelWorkBooks != null ) Marshal.ReleaseComObject( excelWorkBooks );
            if( excelApp != null ) Marshal.ReleaseComObject( excelApp );
            excelRange = null;
            excelWorkSheet = null;
            excelSheets = null;
            excelWorkBook = null;
            excelWorkBooks = null;
            excelApp = null;
        }

        private static string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;

            while (dividend > 0)
            {
                var modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }

        private void RoepAan()
        {
            var list = new List<string>
            {
                "1",
                "2"
            };

            DoeShit(list);
        }

        private void DoeShit(List<string> list)
        {
            list.Add("3");
        }
    }
}
