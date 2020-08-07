using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using ExcelApp = Microsoft.Office.Interop.Excel;
namespace Organogram
{
    class Program
    {
        static void Main(string[] args)
        {
            // Opening excel from given path, parsing all data.
            Excel excel = new Excel(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"companies_data.csv"));

            excel.GetChildren(0,0);

            // Closing excel.
            excel.excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel.excelApp);

            Console.ReadKey();
        }
    }
}
