using Microsoft.Office.Interop.Excel;
using System;
using System.Text;
using Range = Microsoft.Office.Interop.Excel.Range;




namespace UpdateExcel
{
    class Excel
    {
        string path;
        //string SavePath;
        Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
        Workbook wb;
        Worksheet ws;





        public Excel(string path, string sheetName)
        {
            path = path;
            //SavePath = SavePath;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[sheetName];
        }

        public string Results;
        string email;
        List<string> mail = new List<string>();

        public void ReadCells()
        {

            Range cell = ws.Range["A2:A22"];

            foreach (string result in cell.Value)
            {
                Results = result;

                email = string.Concat(Results) + "@gmail.com";
                mail.Add(email);

            }

        }



        public void WriteCells()
        {
            Range CellRange = ws.Range["C2:C22"];

            string[,] data = new string[mail.Count, 1];

            for (int i = 0; i < mail.Count; i++)
            {
                data[i, 0] = mail[i];
            }
            CellRange.Value = data;
            //wb.Close();
        }


        public void Department()
        {
            Range cellRange = ws.Range["D2:D22"];
            Random random = new Random();

            string[] departmentOptions = { "Technology", "Sales", "Marketing", "Administration", "HR", "Finance" };

            foreach (Range cell in cellRange.Cells)
            {
                int randomIndex = random.Next(departmentOptions.Length);
                cell.Value = departmentOptions[randomIndex];
            }
            //wb.Close();
        }


        public void salary()
        {

            Range range = ws.Range["E2:E22"];

            int startingSalary = 1000;
            int salaryIncrement = 500;

            Range firstCell = range.Cells[1];
            firstCell.Value = startingSalary;

            for (int row = 2; row <= range.Rows.Count; row++)
            {
                Range cell = range.Cells[row];
                int previousSalary = Convert.ToInt32(range.Cells[row - 1].Value);
                int newSalary = previousSalary + salaryIncrement;
                cell.Value = newSalary;
            }

        }

        public void jobStatus()
        {
            Range cellRange = ws.Range["F2:F22"];
            Random random = new Random();

            string[] departmentOptions = { "Permanent", "Contract" };

            foreach (Range cell in cellRange.Cells)
            {
                int randomIndex = random.Next(departmentOptions.Length);
                cell.Value = departmentOptions[randomIndex];
            }
            //wb.Close();

            string fileName = $"Employee_{DateTime.Now:dd_MM_yyyy}.xlsx";
            string savePath = "C:\\Users\\EUGENE\\SavePath" + "\\" + fileName;
            wb.SaveAs(savePath);
            wb.Close();
            excel.Quit();
        }


    }


}
