using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Assistent_Modul_Online
{
    class WExcel
    {
        public Excel.Application application = null;
        public Excel.Workbook workbook = null;
        public Excel.Workbooks workbooks = null;
        public Excel.Sheets worksheets = null;
        public Excel.Worksheet worksheet = null;

        public void StartWork (int ic, int jc, string a, string b)
        {            
            try
            {
                application = new Excel.Application();
                application.SheetsInNewWorkbook = 1;
                workbooks = application.Workbooks;
                workbook = workbooks.Add();
                worksheets = application.Sheets;
                worksheet = worksheets.Item[1];
                worksheet.Name = "График сборки строк";
                application.Visible = true;
                for (int i = 1; i < ic; i++)
                {
                    for (int j = 1; j < jc; j++)
                        worksheet.Cells[i, j] = String.Format("d{0}{1}", a, b);
                }
                application.Quit();
            }
            catch
            {

            }
            finally
            {
                
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(worksheets);
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(workbooks);
                Marshal.ReleaseComObject(application);


            }
        }
        public void AddInf()
        {
            
        }
    }
}
