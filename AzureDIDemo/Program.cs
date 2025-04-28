// See https://aka.ms/new-console-template for more information
using Azure;
using Azure.AI.FormRecognizer.DocumentAnalysis;
using OfficeOpenXml;

string folderPath = @"C:\Users\rc12048\Documents\BlueStreak";
string outputFilePath = @"C:\Users\rc12048\Documents\BlueStreak\Output.xlsx";
string modelId = "DemoBlueStreakModel";


string endpoint = "https://extractpointdi.cognitiveservices.azure.com/";
string apiKey = " ";//apikey;
var credential = new AzureKeyCredential(apiKey);
var client = new DocumentAnalysisClient(new Uri(endpoint), credential);
await ProcessPDFsAsync(folderPath, outputFilePath, modelId, client);
Console.WriteLine("Processing Complete.");

static async Task ProcessPDFsAsync(string folderPath, string outputFilePath, string modelId, DocumentAnalysisClient client)
{
    foreach (var filePath in Directory.GetFiles(folderPath, "*.pdf"))
    {
        Console.WriteLine($"Processing file: {Path.GetFileName(filePath)}");

        
        IDictionary<string, string> extractedData = await ExtractKeyValuePairsFromPDF(filePath, client, modelId);

        
        WriteToExcel(outputFilePath, Path.GetFileName(filePath), extractedData);

        Console.WriteLine($"Processed {Path.GetFileName(filePath)}");
    }
}
static async Task<IDictionary<string, string>> ExtractKeyValuePairsFromPDF(string filePath, DocumentAnalysisClient client, string modelId)
{
    using var stream = new FileStream(filePath, FileMode.Open);
    AnalyzeDocumentOperation operation = await client.AnalyzeDocumentAsync(WaitUntil.Completed, modelId, stream);
    AnalyzeResult result = operation.Value;

    
    var keyValuePairs = new Dictionary<string, string>();

    foreach (var field in result.Documents[0].Fields)
    {
        string key = field.Key;
        string value = field.Value.Content;
        keyValuePairs.Add(key, value);
    }

    return keyValuePairs;
}

static void WriteToExcel(string outputFilePath, string fileName, IDictionary<string, string> keyValuePairs)
{
    
    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

    FileInfo fileInfo = new FileInfo(outputFilePath);

    using (ExcelPackage package = new ExcelPackage(fileInfo))
    {
        ExcelWorksheet worksheet;

       
        if (package.Workbook.Worksheets.Count == 0)
        {
            worksheet = package.Workbook.Worksheets.Add("Extracted Data");
            worksheet.Cells[1, 1].Value = "FileName";
            worksheet.Cells[1, 2].Value = "Key";
            worksheet.Cells[1, 3].Value = "Value";
        }
        else
        {
            worksheet = package.Workbook.Worksheets[0];
        }

        int row = worksheet.Dimension?.Rows + 1 ?? 2;

        foreach (var kvp in keyValuePairs)
        {
            worksheet.Cells[row, 1].Value = fileName;
            worksheet.Cells[row, 2].Value = kvp.Key;
            worksheet.Cells[row, 3].Value = kvp.Value;
            row++;
        }

        package.Save();
    }
}





