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
    private readonly Dictionary<string?, SheetTab> _sheetList = new();
    private readonly bool _isTemplate;

    public Spreadsheet(string path, string planSheetName, bool includePlanSheet, bool isTemplate)
    {
        _isTemplate = isTemplate;
        FilePath = Path.GetFullPath(path);
        string? dirPath = Path.GetDirectoryName(path);
        if (!string.IsNullOrWhiteSpace(dirPath) && !Directory.Exists(dirPath))
        {
            Directory.CreateDirectory(dirPath);
        }

        if (isTemplate)
        {
            _document = SpreadsheetDocument.Open(path, true);
        }
        else
        {
            _document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        }
        if (_document.WorkbookPart is null)
        {
            _document.AddWorkbookPart().Workbook = new Workbook();
        }
        _workbook = _document.WorkbookPart!;
        
        PlanSheet = GetSheet(planSheetName, !includePlanSheet);
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

    public SheetTab GetSheet(string name, bool neverInclude = false)
    {
        if (!_sheetList.TryGetValue(name, out SheetTab? sheet))
        {
            sheet = new SheetTab(_workbook, name, neverInclude || _isTemplate, !_isTemplate);
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