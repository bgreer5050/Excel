using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel1
{
    class Program
    {
        static void Main(string[] args)
        {
            var ExcelObj = new Microsoft.Office.Interop.Excel.Application();
            ExcelObj.Visible = false;

            // Workbook wb = ExcelObj.Workbooks.Add()


            try
            {
                Console.WriteLine(Environment.CurrentDirectory);
                //ExcelObj.Workbooks.Open(Environment.CurrentDirectory.ToString() + "\this.xlsx");
                Workbook wb =  ExcelObj.Workbooks.Add();
                wb.Save();
                Console.WriteLine(wb.FullName);
                wb.SaveAs("test.xlsx");
                Console.WriteLine(wb.FullName);


                Console.WriteLine("Success");
            }
            catch (Exception)
            {
                Console.WriteLine("Fail");

            }
            finally
            {
                
            }
            
            while(true)
            {

            }
        }
    }
}
