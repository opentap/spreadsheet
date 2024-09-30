using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Spreadsheet;

public sealed class SheetTab
{
    private readonly WorkbookPart _workbook;
    private readonly SheetData _sheetData;
    private readonly Row _headRow;
    private readonly Columns _columns;
    private readonly Dictionary<string, uint> _columnNameToIndex;
    private readonly Dictionary<string, Column> _columnNameToColumn;
    private readonly Sheet _sheet;
    private bool _isAdded = false;

    public SheetTab(WorkbookPart workbook, string name)
    {
        _workbook = workbook;
        var worksheet =  workbook.AddNewPart<WorksheetPart>();
        worksheet.Worksheet = new Worksheet();

        _columns = worksheet.Worksheet.GetFirstChild<Columns>() ?? worksheet.Worksheet.AppendChild(new Columns());
        _sheetData = worksheet.Worksheet.GetFirstChild<SheetData>() ?? worksheet.Worksheet.AppendChild(new SheetData());
        Sheets sheets = _workbook.Workbook.GetFirstChild<Sheets>() ?? _workbook.Workbook.AppendChild(new Sheets());

        // Get or create a new sheet.
        Sheet? sheet = sheets.Elements<Sheet>().FirstOrDefault(s => s.Name == name);
        if (sheet is null)
        {
            sheet = new Sheet()
            {
                Id = workbook.GetIdOfPart(worksheet),
                Name = name,
            };
        }
        _sheet = sheet;


        // Get or create a new head row.
        Row? headRow = _sheetData.Elements<Row>().FirstOrDefault(r => (r.RowIndex?.HasValue ?? false ? r.RowIndex.Value : 0) == 0);
        if (headRow is null)
        {
            uint rowIndex = _sheetData.Elements<Row>()
                .Select(r => r.RowIndex?.HasValue ?? false ? r.RowIndex.Value : 0)
                .Append(0u)
                .Max() + 1;
            headRow = new Row()
            {
                RowIndex = rowIndex,
            };
            _sheetData.AddChild(headRow);
        }
        _headRow = headRow;
        
        // Add existing cells to column list.
        uint index = 0;
        _columnNameToIndex = new ();
        _columnNameToColumn = new Dictionary<string, Column>();
        foreach (Cell? cell in _headRow.Elements<Cell>())
        {
            if (cell is null)
            {
                continue;
            }
            ParseReference(cell.CellReference, out _, out uint row);
            _columnNameToIndex.Add(cell.InnerText, row);
        }
        
        workbook.Workbook.Save();
    }

    public void EnsureInclusion()
    {
        if (!_isAdded)
        {
            Sheets sheets = _workbook.Workbook.GetFirstChild<Sheets>() ?? _workbook.Workbook.AppendChild(new Sheets());
            uint sheetId = sheets.Elements<Sheet>()
                .Select(s => s.SheetId?.HasValue ?? false ? s.SheetId.Value : 0)
                .Append(0u)
                .Max() + 1;
            _sheet.SheetId = sheetId;
            
            sheets.Append(_sheet);
            _isAdded = true;
        }
    }

    public void AddRows(Dictionary<string, object> parameters, Dictionary<string, Array> results)
    {
        if (parameters.Count == 0 && results.Count == 0)
        {
            return;
        }
        EnsureInclusion();
        
        uint rowIndex = _sheetData.Elements<Row>()
            .Select(r => r.RowIndex?.HasValue ?? false ? r.RowIndex.Value : 0)
            .Append(0u)
            .Max() + 1;

        int rowCount = Math.Max(1, results.Select(kvp => kvp.Value.Length).Append(0).Max());
        Row[] rows = new Row[rowCount];
        for (int i = 0; i < rowCount; i++)
        {
            rows[i] = new Row()
            {
                RowIndex = rowIndex,
            };
            foreach (KeyValuePair<string,object> kvp in parameters)
            {
                string name = kvp.Key;
                object value = kvp.Value;
                AddCell(name, rowIndex, rows[i], value);
            }
            rowIndex += 1;
            _sheetData.Append(rows[i]);
        }
        foreach (KeyValuePair<string, Array> keyValuePair in results)
        {
            string name = keyValuePair.Key;
            Array data = keyValuePair.Value;
            for (int i = 0; i < data.Length; i++)
            {
                Row row = rows[i];
                object? value = data.GetValue(i);
                AddCell(name, row.RowIndex!, row, value);
            }
        }
        
        _workbook.Workbook.Save();
    }

    private void AddCell(string columnName, uint rowIndex, Row row, object value)
    {
        uint columnIndex = GetColumnIndex(columnName);
        string cellReference = CreateCellReference(columnIndex, rowIndex);
        Cell cell = new Cell()
        {
            CellReference = cellReference,
        };
        int length = SetValue(cell, value);
        _columnNameToColumn[columnName].Width = Math.Max(length, _columnNameToColumn[columnName].Width ?? 0);
        
        Cell? referenceCell = row.Elements<Cell>()
            .FirstOrDefault(c =>
            {
                ParseReference(c.CellReference, out uint cCol, out _);
                return cCol > columnIndex;
            });
        
        row.InsertBefore(cell, referenceCell);
    }

    private uint GetColumnIndex(string name)
    {
        if (!_columnNameToIndex.TryGetValue(name, out uint index))
        {
            index = (uint)_columnNameToIndex.Count;
            _columnNameToIndex.Add(name, index);
            string cellReference = CreateCellReference(index, _headRow.RowIndex ?? 1);
            Cell cell = new Cell()
            {
                CellReference = cellReference
            };
            SetValue(cell, name);
            _headRow.Append(cell);

            Column column = new Column()
            {
                Min = index + 1,
                Max = index + 1,
                Width = name.Length,
                CustomWidth = true,
            };
            _columns.Append(column);
            _columnNameToColumn.Add(name, column);
        }
        return index;
    }

    private string CreateCellReference(uint columnIndex, uint rowIndex)
    {
        const uint max = 'Z' - 'A' + 1;
        string reference = (char)('A' + columnIndex % max) + "";
        if (columnIndex >= max)
        {
            reference = (char)('A' + columnIndex / max - 1) + reference;
        }
        if (columnIndex >= max * max)
        {
            reference = (char)('A' + columnIndex / (max * max) - 1) + reference;
        }
        if (columnIndex > max * max * max)
        {
            throw new Exception("Cannot handle this many columns.");
        }
        return reference + rowIndex;
    }

    private static void ParseReference(string? reference, out uint column, out uint row)
    {
        column = 0;
        row = 0;
        if (string.IsNullOrWhiteSpace(reference))
        {
            return;
        }
        
        for (int i = 0; i < reference.Length; i++)
        {
            char c = reference[i];
            if (char.IsLetter(c))
            {
                column = column * ('Z' - 'A' + 1) + c - 'A' + 1;
            }
            else if (char.IsDigit(c))
            {
                row = row * ('9' - '0' + 1) + c - '0';
            }
        }
    }

    private int SetValue(Cell cell, object o)
    {
        switch (o)
        {
            case int i:
                cell.CellValue = new CellValue(i);
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                break;
            case float f:
                cell.CellValue = new CellValue(f);
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                break;
            case double d:
                cell.CellValue = new CellValue(d);
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                break;
            case decimal d:
                cell.CellValue = new CellValue(d);
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                break;
            case bool b:
                cell.CellValue = new CellValue(b);
                cell.DataType = new EnumValue<CellValues>(CellValues.Boolean);
                break;
            case string s:
                cell.CellValue = new CellValue(s);
                cell.DataType = new EnumValue<CellValues>(CellValues.String);
                break;
            case DateTime dt:
                cell.CellValue = new CellValue(dt);
                cell.DataType = new EnumValue<CellValues>(CellValues.Date);
                break;
            case DateTimeOffset dto:
                cell.CellValue = new CellValue(dto);
                cell.DataType = new EnumValue<CellValues>(CellValues.Date);
                break;
            default:
                cell.CellValue = new CellValue(o.ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.String);
                break;
        }

        return o.ToString().Length;
    }
} 