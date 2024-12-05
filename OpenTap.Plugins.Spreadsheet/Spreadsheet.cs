using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Spreadsheet;

public sealed class Spreadsheet : IDisposable
{
    private readonly SpreadsheetDocument _document;
    private readonly WorkbookPart _workbook;
    private readonly WorksheetPart _worksheet;
    private readonly Dictionary<string?, SheetTab> _sheetList = new();
    private readonly bool _isTemplate;

    public Spreadsheet(string path, bool includePlanSheet, bool isTemplate)
    {
        _isTemplate = isTemplate;
        FilePath = Path.GetFullPath(path);
        string? dirPath = Path.GetDirectoryName(path);
        if (!string.IsNullOrWhiteSpace(dirPath) && !Directory.Exists(dirPath))
        {
            Directory.CreateDirectory(dirPath);
        }

        if (File.Exists(path))
        {
            _document = SpreadsheetDocument.Open(path, true);
        }
        else
        {
            if (isTemplate)
            {
                _document = SpreadsheetDocument.Open(path, true);
            }
            else
            {
                _document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
            } 
        }
        
        if (_document.WorkbookPart is null)
        {
            _document.AddWorkbookPart().Workbook = new Workbook();
        }
        _workbook = _document.WorkbookPart!;

        { /* Read existing sheets */
            Sheets sheets = _workbook.Workbook.GetFirstChild<Sheets>() ?? new Sheets();
            var sheetList = sheets.Elements<Sheet>().ToArray();
            foreach (Sheet sht in sheetList)
            {
                _sheetList.Add(sht.Name, new SheetTab(_workbook, sht.Name, true));
            }
        }
        PlanSheet = GetSheet("Runs Summary", !includePlanSheet);
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