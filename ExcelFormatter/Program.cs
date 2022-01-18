using System.Text;
using ClosedXML.Excel;
using Kurukuru;

const string extensionExcel = ".xlsx";


// entry

Console.OutputEncoding = Encoding.UTF8;

var argPath = args.FirstOrDefault();
if (string.IsNullOrEmpty(argPath))
{
    Console.WriteLine("enter excel file or directory");
    return;
}

var excelFiles = SearchExcelFiles(argPath);
var fileCount = excelFiles.Length;

if (fileCount == 0)
{
    Console.WriteLine($"no excel file {argPath}");
    return;
}

using var spinnerTotal = new Spinner($"{argPath}");
spinnerTotal.Start();

var failed = new List<FileInfo>();

for (var count = 0; count < excelFiles.Length; count++)
{
    spinnerTotal.Text = $"{count} / {fileCount}";

    var file = excelFiles[count];
    using var spinner = new Spinner(file.Name, Patterns.Arc);
    spinner.Start();
    try
    {
        Format(file);
        spinner.Succeed(file.Name);
    }
    catch (Exception e)
    {
        Console.WriteLine(e);
        spinner.Fail(e?.Message);
        failed.Add(file);
    }
}

if (failed.Count == 0)
{
    spinnerTotal.Succeed($"{fileCount} / {fileCount}");
}
else
{
    spinnerTotal.Fail($"{argPath} {failed.Count} / {fileCount}");
    foreach (var fileInfo in failed)
    {
        Console.WriteLine(fileInfo.Name);
    }
}


// functions

FileInfo[] SearchExcelFiles(string? path)
{
    if (string.IsNullOrEmpty(path))
    {
        return Array.Empty<FileInfo>();
    }

    if (path.EndsWith(extensionExcel))
    {
        var fileInfo = new FileInfo(path);
        return fileInfo.Exists ? new[] { fileInfo, } : Array.Empty<FileInfo>();
    }

    var dirInfo = new DirectoryInfo(path);
    if (!dirInfo.Exists)
    {
        return Array.Empty<FileInfo>();
    }

    return dirInfo.EnumerateFiles($"*{extensionExcel}")
        .ToArray();
}

void Format(FileSystemInfo info)
{
    using var book = new XLWorkbook(info.FullName);
    book.Author = null;
    foreach (var worksheet in book.Worksheets)
    {
        worksheet.Author = null;
        worksheet.Cell(1, 1).SetActive();
        worksheet.SheetView.ZoomScale = 100;
    }
    book.Save();
}
