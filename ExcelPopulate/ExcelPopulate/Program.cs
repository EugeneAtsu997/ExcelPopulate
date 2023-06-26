// See https://aka.ms/new-console-template for more information
using System;
using UpdateExcel;

namespace ExcelPopulate
{
    public class Program
    {
        public static void Main()
        {
            Console.WriteLine("Hello, World!");
            Excel excel = new Excel("C:\\Users\\EUGENE\\Documents\\SampleOne.xlsx", "Problem");
            excel.ReadCells();
            excel.WriteCells();
            excel.Department();
            excel.jobStatus();
        }
    }
}
