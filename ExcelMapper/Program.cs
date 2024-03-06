using ExcelMapper;
using OfficeOpenXml;

string path = @"C:\Users\Ömer Abay\OneDrive\Masaüstü\odev1.xlsx";
ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
using (var package = new ExcelPackage(new FileInfo(path)))
{
    var worksheet = package.Workbook.Worksheets[0];

    int rowCount = worksheet.Dimension.Rows;
    int colCount = worksheet.Dimension.Columns;
    List<string> allHeader = new List<string>();
    
    ExcelModel model;

    Console.WriteLine($"Reading data from {path}:\n");

    
    for (int col = 1; col <= colCount; col++)
    {
        allHeader.Add(worksheet.Cells[1,col].Text); //get headers 
    }

    for (int row = 2; row <= rowCount; row++)
    {
        List<string> data = new List<string>();
        for (int col = 1; col <= colCount; col++)
        {
            data.Add(worksheet.Cells[row, col].Text); //get the data row by row
        }
        model = Mapper<ExcelModel>.Map(data, allHeader); //mapping
        Console.WriteLine(model.ATTRIBUTES);
    }
}