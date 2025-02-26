using AfifahExcelTest.Models;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using Newtonsoft.Json;

IConfiguration config = new ConfigurationBuilder()
    .SetBasePath(Directory.GetCurrentDirectory()) // Ensure the base path is correct
    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
    .Build();

//string fileName = "C:\\Users\\afifah\\Desktop\\C# Article\\List of File Rename.xlsx"; //Kena masuk appsettings ni...
string fileName = config["RenameFileList"];

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

ExcelPackage package = new ExcelPackage(fileName);

ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

int startLine = 3;

int currentLine = startLine;

List<RenameRule> ruleList = new List<RenameRule>();


while (true)
{
    var senderEmail = worksheet.Cells[$"B{currentLine}"].Value;
    var senderEmail2 = worksheet.Cells[currentLine, 2].Value;

    if (senderEmail == null || senderEmail.ToString().Trim() == "")
    {
        break;
    }

    RenameRule theRule = new RenameRule()
    {
        SenderEmail = worksheet.Cells[currentLine, 2].Value.ToString(),
        Subject = worksheet.Cells[currentLine, 3].Value?.ToString() ?? string.Empty,
        Folder = worksheet.Cells[currentLine, 4].Value?.ToString() ?? string.Empty,
        FileName = worksheet.Cells[currentLine, 5].Value?.ToString() ?? string.Empty,
        Remarks = worksheet.Cells[currentLine, 7].Value?.ToString() ?? string.Empty
    };

    var isRule = worksheet.Cells[currentLine, 6].Value?.ToString() ?? null;

    if (isRule == null)
    {
        theRule.IsRule = false;
    }
    else
    {
        theRule.IsRule = isRule == "Yes" ? true : false;
    }

    ruleList.Add(theRule);

    currentLine++;
}

string niNakTestSender = "nurulhanees@kenanga.com.my";

var ruleSesuai = ruleList
    .Where(x => x.SenderEmail == niNakTestSender)
    .ToList();

Console.WriteLine(JsonConvert.SerializeObject(ruleSesuai, Formatting.Indented));

Console.ReadLine();
