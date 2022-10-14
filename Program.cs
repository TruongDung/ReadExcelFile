using ExcelApp = Microsoft.Office.Interop.Excel;

ExcelApp.Application excelApp = new ExcelApp.Application();

ExcelApp.Workbook excelBook = excelApp.Workbooks.Open(@"D:\Copy of 52874 (004).xlsx");
ExcelApp._Worksheet excelSheet = excelBook.Sheets[1];
ExcelApp.Range excelRange = excelSheet.UsedRange;

int rows = excelRange.Rows.Count;
int cols = excelRange.Columns.Count;

var stringList = new List<string>();
for (int i = 1; i <= rows; i++)
{
    Console.Write("\r\n");
    for (int j = 2; j <= 2; j++)
    {
        if (excelRange.Cells[i, j] != null && excelRange.Cells[i, j].Value2 != null)
            stringList.Add( "'" + excelRange.Cells[i, j].Value2.ToString() + "'");
    }
}
var result = string.Join(",", stringList);

excelApp.Quit();
System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
Console.ReadLine();
