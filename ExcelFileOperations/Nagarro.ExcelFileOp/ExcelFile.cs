using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace Nagarro.ExcelFileOp
{
    public class ExcelFile
    {
        private string FileFullPathName;
        private Application xlApp;
        private Workbook xlWorkBook;
        public Worksheet xlWorkSheet;
        private Range range;

        /// <summary>
        /// Default constructor initializes excel file variables
        /// </summary>
        private ExcelFile()
        {
        }

        /// <summary>
        /// Creates new excel file 
        /// </summary>
        /// <param name="FileFullPathName">Full path of the file plus file name with extension of .xls</param>
        public ExcelFile(string FileFullPathName)
            : this()
        {
            this.FileFullPathName = FileFullPathName;

        }

        /// <summary>
        /// Creates new excel file
        /// </summary>
        /// <param name="FileFullPathName">Full path of the file plus file name with extension of .xls</param>
        public void CreateExcelFile()
        {
            try
            {
                xlApp = new Application();

                if (xlApp == null)
                {
                    throw new Exception("Excel is not properly installed!!");
                }

                object misValue = System.Reflection.Missing.Value;

                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(1);

                xlWorkBook.SaveAs(FileFullPathName, XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);


                SaveAndClose();


                //xlWorkBook.Close(true, misValue, misValue);
                //xlApp.Quit();

                //releaseObject(xlWorkSheet);
                //releaseObject(xlWorkBook);
                //releaseObject(xlApp);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        public string ReadExcelFile()
        {
            
            string str="";
            int rCnt = 0;
            int cCnt = 0;

            xlApp = new Application();
            xlWorkBook = xlApp.Workbooks.Open(FileFullPathName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;

            for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
            {
                for (cCnt = 1; cCnt <= range.Columns.Count; cCnt++)
                {
                    str = str + " " + (string)(range.Cells[rCnt, cCnt] as Range).Value2.ToString();
                }
                str = str + "\n";
            }

            SaveAndClose();
            //xlWorkBook.Close(true, null, null);
            //xlApp.Quit();

            //releaseObject(xlWorkSheet);
            //releaseObject(xlWorkBook);
            //releaseObject(xlApp);

            return str;
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                throw new Exception("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        public void SaveAndClose()
        {
            object misValue = System.Reflection.Missing.Value;
            xlWorkBook.Save();

            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);

        }

        public Range GetCellArray()
        {
           
            //object misValue = System.Reflection.Missing.Value;

            //xlApp = new Application();
            //xlWorkBook = xlApp.Workbooks.Open(FileFullPathName, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            //xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(1);

            return (Range)xlWorkSheet.Cells;


        }

        public void OpenExelFile()
        {
           
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Application();
            //xlWorkBook = xlApp.Workbooks.Open(FileFullPathName, 0, true, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkBook = xlApp.Workbooks.Open(FileFullPathName, 0, false, 5, "", "", false, XlPlatform.xlWindows, "", true, false, 0, true, false, false);

            xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(1);


            //xlWorkSheet.get_Range("A1", "A1").Value2 = "Newest content";

            //xlWorkSheet.Cells[1, 1] = "Newer content";
            // MessageBox.Show(xlWorkSheet.get_Range("A1", "A1").Value2.ToString());

            //xlWorkBook.Close(true, misValue, misValue);
            //xlApp.Quit();

            //releaseObject(xlWorkSheet);
            //releaseObject(xlWorkBook);
            //releaseObject(xlApp);
        }
    }
}
