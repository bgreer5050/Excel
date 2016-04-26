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
            ExcelObj.Visible = true;

            // Workbook wb = ExcelObj.Workbooks.Add()


            try
            {
                Console.WriteLine(Environment.CurrentDirectory);
                //ExcelObj.Workbooks.Open(Environment.CurrentDirectory.ToString() + "\this.xlsx");
                Workbook wb = ExcelObj.Workbooks.Open("test2.xlsx"); //  .Workbooks.Add();

                Worksheet ws = wb.Worksheets[1];
                Microsoft.Office.Interop.Excel.Range xlRange = ws.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;

                for (int i = 1; i <= rowCount; i++)
{
                    for (int j = 1; j <= colCount; j++)
  {
                        Console.WriteLine(xlRange.Cells[i, j].Value2.ToString());
                    }
                }


             
                //wb.Save();
                //Console.WriteLine(wb.FullName);
                //wb.SaveAs("test.xlsx");
                //Console.WriteLine(wb.FullName);
                //wb.Op
                //wb.Open("test.xlsx");


                Console.WriteLine("Success");
            }
            catch (Exception ex)
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
