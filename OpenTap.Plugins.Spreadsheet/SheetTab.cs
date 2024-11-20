using System;
using System.Collections.Generic;
using System.Linq;
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
    private readonly bool _allowNewColumns;

    public SheetTab(WorkbookPart workbook, string name, bool neverInclude = false, bool allowNewColumns = true)
    {
        _workbook = workbook;
        _isAdded = neverInclude;
        _allowNewColumns = allowNewColumns;
        Name = name;

        // Get or create a new sheet.
        Sheets sheets = _workbook.Workbook.GetFirstChild<Sheets>() ?? _workbook.Workbook.AppendChild(new Sheets());
        Sheet? sheet = sheets.Elements<Sheet>().FirstOrDefault(s => s.Name == name);
        if (sheet is null)
        {
            sheet = new Sheet()
            {
                Name = name,
            };
        }

        WorksheetPart? worksheet = workbook.WorksheetParts.FirstOrDefault(w => workbook.GetIdOfPart(w) == sheet.Id);
        if (worksheet is null)
        {
            worksheet = workbook.AddNewPart<WorksheetPart>();
            worksheet.Worksheet = new Worksheet();
        }

        if (sheet.Id is null)
        {
            sheet.Id = workbook.GetIdOfPart(worksheet);
        }

        _sheet = sheet;
        _columns = worksheet.Worksheet.GetFirstChild<Columns>() ?? worksheet.Worksheet.AppendChild(new Columns());
        _sheetData = worksheet.Worksheet.GetFirstChild<SheetData>() ?? worksheet.Worksheet.AppendChild(new SheetData());

        // Get or create a new head row.
        Row? headRow = _sheetData.Elements<Row>().FirstOrDefault(r => (r.RowIndex?.HasValue ?? false ? r.RowIndex.Value : 0) == 1);
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
        _columnNameToIndex = new ();
        _columnNameToColumn = new Dictionary<string, Column>();
        foreach (Cell? cell in _headRow.Elements<Cell>())
        {
            if (cell is null)
            {
                continue;
            }
            ParseReference(cell.CellReference, out uint col, out _);
            string colName = GetCellValue(cell);
            _columnNameToIndex.Add(colName, col);
            Column column = _columns.Elements<Column>().FirstOrDefault(c => (c?.Min ?? 0) <= col && col <= (c?.Max ?? 0)) ??
                            _columns.AppendChild(new Column()
                            {
                                Min = col,
                                Max = col,
                            });
            _columnNameToColumn.Add(colName, column);
        }
        
        workbook.Workbook.Save();
    }
    
    public string Name { get; }
    
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
        if (!TryGetColumnIndex(columnName, out uint columnIndex))
        {
            return;
        }
        
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

    private bool TryGetColumnIndex(string name, out uint index)
    {
        if (_columnNameToIndex.TryGetValue(name, out index))
        {
            return true;
        }
        
        if (_allowNewColumns)
        {
            index = (uint)_columnNameToIndex.Count + 1;
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
                Min = index,
                Max = index,
                Width = name.Length,
                CustomWidth = true,
            };
            _columns.Append(column);
            _columnNameToColumn.Add(name, column);
            return true;
        }

        return false;
    }

    private static string CreateCellReference(uint columnIndex, uint rowIndex)
    {
        const int max = 'Z' - 'A' + 1;
        int col = (int)columnIndex;
        if (col <= 0)
            throw new IndexOutOfRangeException($"Column index must be greater than 0.");
        if (col > 16384 /* XFD, the max column value microsoft excel allows */)
        {
            throw new IndexOutOfRangeException("Cannot handle this many columns.");
        }

        string reference = "";
        while (col > 0)
        {
            int digit = (col - 1) % max;
            char ch = (char)('A' + digit);
            reference = ch + reference;
            // base26digits.Insert(0, i);
            col = (col - 1) / max;
        } 
        
        reference += rowIndex; 
        return reference;
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

    private string GetCellValue(Cell cell)
    {
        if (cell.DataType is null)
        {
            return cell.CellValue?.Text ?? "";
        }
        if (cell.DataType == CellValues.SharedString && (cell.CellValue?.TryGetInt(out int index) ?? false))
        {
            SharedStringTable stringTable = _workbook.SharedStringTablePart?.SharedStringTable ??
                                            (_workbook.AddNewPart<SharedStringTablePart>().SharedStringTable = new SharedStringTable());

            SharedStringItem stringItem = stringTable.Elements<SharedStringItem>().Skip(index).Take(1).First();
            return stringItem.Text?.Text ?? "";
        }

        return cell.CellValue?.Text ?? "";
    }

    private static int SetValue(Cell cell, object o)
    {
        switch (o)
        {
            case int i:
                cell.CellValue = new CellValue(i);
                cell.DataType = CellValues.Number;
                break;
            case float f:
                cell.CellValue = new CellValue(f);
                cell.DataType = CellValues.Number;
                break;
            case double d:
                cell.CellValue = new CellValue(d);
                cell.DataType = CellValues.Number;
                break;
            case decimal d:
                cell.CellValue = new CellValue(d);
                cell.DataType = CellValues.Number;
                break;
            case bool b:
                cell.CellValue = new CellValue(b);
                cell.DataType = CellValues.Boolean;
                break;
            case string s:
                cell.CellValue = new CellValue(s);
                cell.DataType = CellValues.String;
                break;
            case DateTime dt:
                cell.CellValue = new CellValue(dt);
                cell.DataType = CellValues.Date;
                break;
            case DateTimeOffset dto:
                cell.CellValue = new CellValue(dto);
                cell.DataType = CellValues.Date;
                break;
            default:
                cell.CellValue = new CellValue(o.ToString());
                cell.DataType = CellValues.String;
                break;
        }

        return o.ToString().Length;
    }
} 
