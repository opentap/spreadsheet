using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenTap;

namespace Spreadsheet;

public sealed class Spreadsheet : IDisposable
{
    private readonly SpreadsheetDocument _document;
    private readonly WorkbookPart _workbook;
    private readonly WorksheetPart _worksheet;
    private readonly Dictionary<string, SheetTab> _sheetList = new();

    public Spreadsheet(string path, string planSheetName)
    {
        FilePath = Path.GetFullPath(path);
        string? dirPath = Path.GetDirectoryName(path);
        if (dirPath is not null && !Directory.Exists(dirPath))
        {
            Directory.CreateDirectory(dirPath);
        }
        _document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        _workbook = _document.AddWorkbookPart();
        _workbook.Workbook = new Workbook();

        PlanSheet = GetSheet(planSheetName);
    }
    
    public string FilePath { get; }

    public SheetTab PlanSheet { get; }

    public bool FileEmpty
    {
        get
        {
            Sheets? sheets = _workbook.Workbook.GetFirstChild<Sheets>();
            return !(sheets?.HasChildren ?? false);
        }
    }

    public SheetTab GetSheet(string name)
    {
        if (!_sheetList.TryGetValue(name, out SheetTab? sheet))
        {
            sheet = new SheetTab(_workbook, name);
            _sheetList.Add(name, sheet);
        }

        return sheet;
    }

    public void Dispose()
    {
        _workbook.Workbook.Save();
        _document.Save();
        _document.Dispose();
    }
}