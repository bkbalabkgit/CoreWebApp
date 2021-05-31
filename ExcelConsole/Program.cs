using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
namespace ExcelConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            if (excelApp != null)
            {
                Microsoft.Office.Interop.Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(@"D:\Output File\TestBatch15_Amendment.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Microsoft.Office.Interop.Excel.Worksheet excelWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)excelWorkbook.Sheets[1];

                Microsoft.Office.Interop.Excel.Range excelRange = excelWorksheet.UsedRange;
                int rowCount = excelRange.Rows.Count;
                int colCount = excelRange.Columns.Count;

                for (int i = 1; i <= rowCount; i++)
                {
                    if (i > 1)
                    {
                        if (colCount > 1)
                        {
                            Microsoft.Office.Interop.Excel.Range firstColumn = (excelWorksheet.Cells[i, 1] as Microsoft.Office.Interop.Excel.Range);
                            string RegExKey = firstColumn.Value.ToString();

                            Microsoft.Office.Interop.Excel.Range secondColumn = (excelWorksheet.Cells[i, 2] as Microsoft.Office.Interop.Excel.Range);
                            string RegExValue = secondColumn.Value.ToString();
                            Regex regex = new Regex(RegExValue);
                            Match match = regex.Match("cdcdde");
                            if (match.Success)
                            {
                                Context.Field("LaycanDeliveryPeriod").Text= match.Value;
                                break;
                            }
                        }
                    }
                }

                excelWorkbook.Close();
                excelApp.Quit();
            }
        }
    }
}
