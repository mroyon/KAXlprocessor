using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");


            string fileName = "romfull.xlsx";
            var wbookSource = new XLWorkbook(fileName);
            var ws1 = wbookSource.Worksheet(1);
            var datarows = ws1.RangeUsed().RowsUsed().Skip(1);
            getRows(ws1, datarows);

            using (var workbookDestination = new XLWorkbook())
            {
                var worksheet = workbookDestination.Worksheets.Add("NewModed");

                //start loop here
                worksheet.Cell("A1").Value = "Hello World!";
                worksheet.Cell("A2").FormulaA1 = "=MID(A1, 7, 5)";
                //end loop here
                 workbookDestination.SaveAs("HelloWorld.xlsx");
            }
        }

        private static void getRows(IXLWorksheet ws1, IEnumerable<IXLRangeRow> datarows)
        {
            int i = 1;
            foreach (var singlerow in datarows)
            {
                var rowNumber = singlerow.RowNumber();
                var excelrow = ws1.Row(rowNumber);
                if (excelrow.IsEmpty())
                    continue;
                else
                {
                    var category = excelrow.Cell(2);
                    var linesofpiem = excelrow.Cell(3);

                    string valueofcategory = category.GetValue<string>();
                    string valueofpoems = linesofpiem.GetValue<string>();
                   
                    using (System.IO.StringReader reader = new System.IO.StringReader(valueofpoems))
                    {
                        string line;
                        while ((line = reader.ReadLine()) != null)
                        {
                            Console.WriteLine(i);
                            Console.WriteLine(line);
                            i++;
                        }
                    }
                }
            }
        }

    }
}
