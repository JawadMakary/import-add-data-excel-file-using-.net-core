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