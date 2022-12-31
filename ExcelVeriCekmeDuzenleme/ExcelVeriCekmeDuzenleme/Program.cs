using ExcelDataReader;
using Microsoft.Office.Interop.Excel;
using Spire.Xls;
using Workbook = Spire.Xls.Workbook;

string fileSource = @"C:\Users\ZAFER\OneDrive\Masaüstü\Yeni Microsoft Office Excel Çalışma Sayfası.xlsx";

FileStream stream = File.Open(fileSource, FileMode.Open, FileAccess.Read);

System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

IExcelDataReader excelDataReader;

if (Path.GetExtension(fileSource).ToUpper() == ".XLS")
{
    //Reading from a binary Excel file ('97-2003 format; *.xls)
    excelDataReader = ExcelReaderFactory.CreateBinaryReader(stream);
}
else
{
    //Reading from a OpenXml Excel file (2007 format; *.xlsx)
    excelDataReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
}

int counter = 1;

List<string> veriler = new List<string>();
while (excelDataReader.Read())
{
    veriler.Add(excelDataReader.GetString(0));
    ++counter;
}

string[] veriler2 = new string[excelDataReader.RowCount];



//Create a Workbook object
Workbook workbook = new Workbook();
//Remove default worksheets
workbook.Worksheets.Clear();
//Add a worksheet and name it
Spire.Xls.Worksheet worksheet = workbook.Worksheets.Add("WriteToCell");



for (int i = 1, j = 1; i < excelDataReader.RowCount; i++)
{
    string[] veri = veriler[i-1].Split(',');
    j = 1;
    foreach (string item in veri)
    {
        worksheet.Range[i, j++].Value = item;
    }
    Console.WriteLine();
}

//Auto fit column width
worksheet.AllocatedRange.AutoFitColumns();
//Save to an Excel file
workbook.SaveToFile("C:\\Users\\ZAFER\\OneDrive\\Masaüstü\\yeniDosya.xlsx", ExcelVersion.Version2016);