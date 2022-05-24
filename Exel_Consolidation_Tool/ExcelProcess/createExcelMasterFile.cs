using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
// Add the following namespace to work with excel files
using Excel = Microsoft.Office.Interop.Excel;

namespace Exel_Consolidation_Tool.ExcelProcess
{
    public class createExcelMasterFile
    {
        public void createFile(string path)
        {
            var ExcelRoot = Directory.GetFiles(path).Where(name => name.EndsWith("ExcelRoot.xlsx"));
            var excel = new Excel.Application();
            if (!ExcelRoot.Any())
            {
                try
                {
                    

                    var workBooks = excel.Workbooks;
                    var workBook = workBooks.Add();
                    var workSheet = (Excel.Worksheet)excel.ActiveSheet;

                    workBook.SaveAs(path + @"\ExcelRoot.xlsx");
                    workBook.Close(false, Type.Missing, Type.Missing);
                    //excel.Quit();
                }
                catch (Exception exc)
                {
                    MessageBox.Show("Error creating Root Excel " + exc.Message);
                    throw;
                }
                finally
                {
                    excel.Quit();
                }
                

            }
            
        }


    }

    

}
