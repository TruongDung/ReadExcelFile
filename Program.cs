// See https://aka.ms/new-console-template for more information
//Console.WriteLine("Hello, World!");
using ExcelApp = Microsoft.Office.Interop.Excel;

ExcelApp.Application excelApp = new ExcelApp.Application();

ExcelApp.Workbook excelBook = excelApp.Workbooks.Open(@"D:\Copy of 52874 (004).xlsx");
ExcelApp._Worksheet excelSheet = excelBook.Sheets[1];
ExcelApp.Range excelRange = excelSheet.UsedRange;

int rows = excelRange.Rows.Count;
int cols = excelRange.Columns.Count;

var list1 = new List<string>();
for (int i = 1; i <= rows; i++)
{
    //create new line
    Console.Write("\r\n");
    for (int j = 2; j <= 2; j++)
    {
        //write the console
        if (excelRange.Cells[i, j] != null && excelRange.Cells[i, j].Value2 != null)
            list1.Add( "'" + excelRange.Cells[i, j].Value2.ToString() + "'");
            //Console.Write(excelRange.Cells[i, j].Value2.ToString());
    }
}
var a2 = string.Join(",", list1);

excelApp.Quit();
System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
Console.ReadLine();
