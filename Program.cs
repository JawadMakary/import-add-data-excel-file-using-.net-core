using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
// import excel library
using ClosedXML.Excel;
using var wbook = new XLWorkbook("products.xlsx");

var ws1 = wbook.Worksheet(1); 
// get all data using forEach
foreach (var row in ws1.Rows())
{
    foreach (var cell in row.Cells())
    {
        Console.Write(cell.Value.ToString() + " ");
    }
    Console.WriteLine();
}
// write data to excel sheet ws1 where there is no data inside cells
ws1.Cell("A4").Value = "4";
ws1.Cell("B4").Value = "UI/UX";
ws1.Cell("C4").Value = "500";
// save the workbook
wbook.SaveAs("products.xlsx");
