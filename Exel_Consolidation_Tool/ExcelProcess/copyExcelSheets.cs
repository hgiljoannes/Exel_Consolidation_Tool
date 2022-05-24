using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Reflection;
using System.IO;
using System.Linq;

namespace Exel_Consolidation_Tool.ExcelProcess
{
    public class copyExcelSheets
    {
        _Application app = new Microsoft.Office.Interop.Excel.Application();
        Workbook curWorkBook = null;
        Workbook destWorkbook = null;
        Worksheet workSheet = null;
        Worksheet newWorksheet = null;
        Object defaultArg = Type.Missing;
       
        public void Copy(string pathExcel, string ExcelRoot)
        {
            try
            {

                // Copy the source sheet
                curWorkBook = app.Workbooks.Open(pathExcel, defaultArg, defaultArg, defaultArg, defaultArg, defaultArg, defaultArg, defaultArg, defaultArg, defaultArg, defaultArg, defaultArg, defaultArg, defaultArg, defaultArg);
                destWorkbook = app.Workbooks.Open(ExcelRoot + @"\ExcelRoot.xlsx", defaultArg, false, defaultArg, defaultArg, defaultArg, defaultArg, defaultArg, defaultArg, defaultArg, defaultArg, defaultArg, defaultArg, defaultArg, defaultArg);

                foreach (Worksheet worksheet in curWorkBook.Worksheets)
                {
                    workSheet = worksheet;
                    workSheet.UsedRange.Copy(defaultArg);

                    // Paste on destination sheet
                    newWorksheet = (Worksheet)destWorkbook.Worksheets.Add(defaultArg, defaultArg, defaultArg, defaultArg);
                    newWorksheet.UsedRange._PasteSpecial(XlPasteType.xlPasteValues, XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                }
                

            }
            catch (Exception exc)
            {
                MessageBox.Show("Error Coping File to Root Excel " + exc.Message);
               
                throw;
            }
            finally
            {

                if (curWorkBook != null)
                {
                    curWorkBook.Save();
                    curWorkBook.Close(defaultArg, defaultArg, defaultArg);

                }

                if (destWorkbook != null)
                {
                    destWorkbook.Save();
                    destWorkbook.Close(defaultArg, defaultArg, defaultArg);
                }

            

            }

            app.Quit();
            
        }



    }
}
