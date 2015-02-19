using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenTap.Plugins.Spreadsheet.OpenXml
{
    public class OpenXmlWorksheet : SpecificWorksheet
    {
        internal SpreadsheetDocument Document { get; set; }
        internal Sheet Worksheet { get; set; }
        internal override string Name
        {
            get
            {
                return Worksheet.Name;
            }
            set
            {
                Worksheet.Name = value;
            }
        }

        internal override int ColumnsCount
        {
            get
            {
                //Worksheet.
                return 0;
            }
        }

        internal override bool Dispose()
        {
            return true;
        }

        internal override void Initialize(object actualWorksheet)
        {
            Worksheet = (Sheet)actualWorksheet;
        }

        internal override object GetCell(int x, int y)
        {
            WorksheetPart worksheetPart = (WorksheetPart)Document.WorkbookPart.GetPartById(Worksheet.Id.Value);
            Worksheet workSheet = worksheetPart.Worksheet;
            SheetData sheetData = workSheet.GetFirstChild<SheetData>();
            IEnumerable<Row> rows = sheetData.Descendants<Row>();
            return GetCell(x - 1, y - 1, Document, rows);
        }

        internal override void SetCell(int x, int y, object value)
        {

            WorksheetPart worksheetPart = (WorksheetPart)Document.WorkbookPart.GetPartById(Worksheet.Id.Value);
            SetCell(x - 1, y - 1, worksheetPart, value);
        }

        public static T[] GetRow<T>(T[,] matrix, int row)
        {
            var columns = matrix.GetLength(1);
            var array = new T[columns];
            for (int i = 0; i < columns; ++i)
                array[i] = matrix[row, i];
            return array;
        }

        internal override void SetManyCells(object[,] cellValues, int rowOffset = 0, int columnOffset = 0)
        {

            //// start at one to satisfy Excel
            //for (int x = 1; x <= cellValues.GetLength(0); ++x)
            //{
            //    for (int y = 1; y <= cellValues.GetLength(1); ++y)
            //    {
            //        SetCell(x + VirtualSpreadsheet.FirstRowOfActualData - 1, y, cellValues[x - 1, y - 1]);
            //    }
            //}

            int numRows = cellValues.GetLength(0);
            int numCols = cellValues.GetLength(1);

            WorksheetPart worksheetPart = (WorksheetPart)Document.WorkbookPart.GetPartById(Worksheet.Id.Value);
            Worksheet workSheet = worksheetPart.Worksheet;
            SheetData originalSheetData = workSheet.GetFirstChild<SheetData>();
            IEnumerable<Row> originalRows = originalSheetData.Descendants<Row>();

            var workbookPart = Document.WorkbookPart;
            string origninalSheetId = workbookPart.GetIdOfPart(worksheetPart);

            WorksheetPart replacementPart =
            workbookPart.AddNewPart<WorksheetPart>();
            string replacementPartId = workbookPart.GetIdOfPart(replacementPart);

            OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);
            OpenXmlWriter writer = OpenXmlWriter.Create(replacementPart);

            Row r = new Row();
            Cell c = new Cell();
            while (reader.Read())
            {
                if (reader.ElementType == typeof(SheetData))
                {
                    if (reader.IsEndElement)
                        continue;
                    writer.WriteStartElement(new SheetData());

                    // Write headers
                    for (int i = 0; i < originalRows.Count(); ++i)
                    {
                        writer.WriteStartElement(r);

                        var cells = originalRows.ElementAt(i).Descendants<Cell>();
                        for (int y = 0; y < cells.Count(); ++y)
                        {
                            writer.WriteElement(cells.ElementAt(y));
                        }

                        writer.WriteEndElement();
                    }

                    // Write data
                    for (int row = 0; row < numRows; row++)
                    {
                        writer.WriteStartElement(r);
                        for (int col = 0; col < numCols; col++)
                        {
                            SetCellToValue(c, cellValues[row, col]);
                            writer.WriteElement(c);
                        }
                        writer.WriteEndElement();
                    }

                    writer.WriteEndElement();
                }
                else
                {
                    if (reader.IsStartElement)
                    {
                        writer.WriteStartElement(reader);
                    }
                    else if (reader.IsEndElement)
                    {
                        writer.WriteEndElement();
                    }
                }
            }


            reader.Close();
            writer.Close();

            Sheet sheet = workbookPart.Workbook.Descendants<Sheet>()
            .Where(s => s.Id.Value.Equals(origninalSheetId)).First();
            sheet.Id.Value = replacementPartId;
            workbookPart.DeletePart(worksheetPart);


            //SheetData sheetData = workSheet.GetFirstChild<SheetData>();
            //IEnumerable<Row> rows = sheetData.Descendants<Row>();

            //for (int i = 0; i < cellValues.GetLength(0); ++i)
            //{
            //    object[] rowData = GetRow(cellValues, i);
            //    // Insert all row data
            //    var openXmlRowData = new List<Cell>(); //new List<OpenXmlElement>();
            //    foreach (var rd in rowData)
            //    {
            //        var cell = new Cell();
            //        SetCellToValue(cell, rd);
            //        openXmlRowData.Add(cell); //ConstructCell(rd.ToString(), TypeToCellValues(rd.GetType())));
            //    }

            //    var row = new Row();
            //    row.Append(openXmlRowData);
            //    sheetData.Append(row);

            //    //if (i % 10000 == 0)
            //    //{
            //    //    workSheet.Save();
            //    //}
            //}

            //workSheet.Save();




            //Row row = null;
            //do
            //{
            //    try
            //    {
            //        row = rows.ElementAt(rowIndex);
            //    }
            //    catch
            //    {
            //        // Row doesn't exist, create a row
            //        var newRow = new Row();
            //        sheetData.Append(newRow);
            //    }
            //} while (row == null);

            //Cell cell = null;
            //do
            //{
            //    try
            //    {
            //        cell = row.Descendants<Cell>().ElementAt(columnIndex);
            //    }
            //    catch
            //    {
            //        // missing a cell here, create one:
            //        row.Append(ConstructCell("", CellValues.SharedString));
            //    }
            //} while (cell == null);


            //SetCellToValue(cell, value);

        }

        internal override void AutoFitColumns()
        {

        }




        #region Not Interface Functions

        private static object GetCell(int rowIndex, int columnIndex, SpreadsheetDocument spreadSheetDocument, IEnumerable<Row> rows)
        {
            var row = rows.ElementAt(rowIndex);
            var rowCellData = GetCellValue(spreadSheetDocument, row.Descendants<Cell>().ElementAt(columnIndex));
            return rowCellData;
        }

        private static void SetCell(int rowIndex, int columnIndex, WorksheetPart worksheetPart, object value)
        {

            Worksheet workSheet = worksheetPart.Worksheet;
            SheetData sheetData = workSheet.GetFirstChild<SheetData>();
            IEnumerable<Row> rows = sheetData.Descendants<Row>();
            //var row = rows.ElementAt(rowIndex);
            //var cell = row.Descendants<Cell>().ElementAt(columnIndex);

            Row row = null;
            do
            {
                try
                {
                    row = rows.ElementAt(rowIndex);
                }
                catch
                {
                    // Row doesn't exist, create a row
                    var newRow = new Row();
                    sheetData.Append(newRow);
                }
            } while (row == null);

            Cell cell = null;
            do
            {
                try
                {
                    cell = row.Descendants<Cell>().ElementAt(columnIndex);
                }
                catch
                {
                    // missing a cell here, create one:
                    row.Append(ConstructCell("", CellValues.SharedString));
                }
            } while (cell == null);


            SetCellToValue(cell, value);
        }

        private static Cell ConstructCell(string value, CellValues dataType)
        {
            return new Cell()
            {
                CellValue = new CellValue(value),
                DataType = new EnumValue<CellValues>(dataType)
            };
        }

        private CellValues TypeToCellValues(Type type)
        {
            CellValues cellValues = CellValues.String; // TODO: Change to SharedString to greatly improve performance
            if (type == typeof(Double) || type == typeof(Single) || 
                type == typeof(Int16) || type == typeof(Int32) || type == typeof(Int64) || 
                type == typeof(UInt16) || type == typeof(UInt32) || type == typeof(UInt64))
            {
                cellValues = CellValues.Number;
            }
            else if (type == typeof(Boolean))
            {
                cellValues = CellValues.Boolean;
            }
            else if (type == typeof(DateTime))
            {
                cellValues = CellValues.Date;
            }
            else
            {
                // All other types are simply string
            }

            return cellValues;
        }

        #region Cell Helpers
        private static object GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
            string value = cell.CellValue.InnerXml;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
            }
            else if (cell.DataType != null && cell.DataType.Value == CellValues.Boolean)
            {
                bool resultBool = false;
                bool success = Boolean.TryParse(value, out resultBool);

                if (success == true)
                    return resultBool;

                // Boolean could be expressed as a number
                try
                {
                    resultBool = Convert.ToBoolean(Convert.ToInt16(value));
                    return resultBool;
                }
                catch (FormatException)
                {

                }

                return value;
            }
            else if (cell.DataType != null && cell.DataType.Value == CellValues.Number)
            {
                // Number
                double resultDouble = 0;
                bool success = Double.TryParse(value, out resultDouble);

                if (success == true)
                    return resultDouble;

                return value;

            }
            else if (cell.DataType != null && cell.DataType.Value == CellValues.Date)
            {

                double oaDate;
                if (Double.TryParse(cell.InnerText, out oaDate))
                {
                    return DateTime.FromOADate(oaDate);
                }

                return value;
            }
            else if (cell.DataType == null)
            {
                // Could possibly be either number or datetime


                uint formatId = 0xDEADBEEF;

                if (cell.StyleIndex != null)
                {
                    int styleIndex = (int)cell.StyleIndex.Value;

                    CellFormat cellFormat = (CellFormat)document.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.ElementAt(styleIndex);

                    formatId = cellFormat.NumberFormatId.Value;
                }

                // Just assume anything not 0xDEADBEEF is date
                if (formatId != 0xDEADBEEF)  //formatId == (uint)FormatIds.DateShort || formatId == (uint)FormatIds.DateLong)
                {
                    // Try number
                    double oaDate;
                    if (Double.TryParse(cell.InnerText, out oaDate))
                    {
                        return DateTime.FromOADate(oaDate);
                    }

                }
                else
                {
                    //value = cell.InnerText;

                    double resultDouble = 0;
                    bool success = Double.TryParse(value, out resultDouble);

                    if (success == true)
                        return resultDouble;
                }

                // If for some reason this isn't a double or a valid datetime, return as string
                return value;

                //double resultDouble = 0;
                //bool success = Double.TryParse(value, out resultDouble);

                //if (success == true)
                //    return resultDouble;

                //DateTime resultDateTime = DateTime.Now;
                //success = DateTime.TryParse(value, out resultDateTime);

                //if (success == true)
                //    return resultDateTime;

                //return value;
            }
            else
            {
                // Some other kind of case
                return value;
            }
        }

        private static void SetCellToValue(Cell cell, object value)
        {

            if (value == null)
            {
                cell.DataType = new EnumValue<CellValues>(CellValues.String);
                cell.CellValue = new CellValue("NULL");
                return;
            }

            var type = value.GetType();
            if (type == typeof(Double) || type == typeof(Single) ||
                type == typeof(Int16) || type == typeof(Int32) || type == typeof(Int64) ||
                type == typeof(UInt16) || type == typeof(UInt32) || type == typeof(UInt64))
            {
                //cellValuesType = CellValues.Number;
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                //cell.StyleIndex = null;
                cell.CellValue = new CellValue(value.ToString());
            }
            else if (type == typeof(Boolean))
            {
                cell.DataType = new EnumValue<CellValues>(CellValues.Boolean);
                cell.CellValue = new CellValue(value as bool? == true ? "1" : "0");
            }
            else if (type == typeof(DateTime))
            {
                cell.DataType = new EnumValue<CellValues>(CellValues.Date);
                //cell.DataType = null;
                cell.CellValue = new CellValue((value as DateTime?).Value.ToOADate().ToString(CultureInfo.InvariantCulture));
            }
            else
            {
                cell.DataType = new EnumValue<CellValues>(CellValues.String);
                cell.CellValue = new CellValue(value.ToString());
            }
        }
        #endregion

        #endregion

    }
}
