using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenTap.Plugins.Spreadsheet.OpenXml
{
    public class OpenXmlWorkbook : SpecificWorkbook
    {

        //private WorkbookPart workbookPart = null;
        //private WorksheetPart worksheetPart = null;
        private Sheets _sheets = null;
        private UInt32Value _sheetId = 1;
        public SpreadsheetDocument Document { get; set; }
        public OpenXmlWorkbook(string filename)
        {
            _sheetId = 1;

            SpreadsheetDocumentType documentType = SpreadsheetDocumentType.Workbook;
            Document = SpreadsheetDocument.Create(filename, documentType);
            Filename = filename;

            WorkbookPart workbookPart = Document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            //worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            //worksheetPart.Worksheet = new Worksheet(new SheetData()); //new Worksheet();

            _sheets = workbookPart.Workbook.AppendChild(new Sheets());
        }

        internal override string Filename { get; set; }

        internal override bool Close()
        {
            Document.WorkbookPart.Workbook.Save();
            Document.Save();
            Document.Close();
            // TODO: Maybe put a "Document Close in OpenXmlSpecifics"
            //Document.WorkbookPart.Close
            return true;
        }

        internal override bool Dispose()
        {
            return true;
        }

        internal override void Initialize(string filename)
        {
            // Open existing file
            SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(filename, false);
            Document = spreadSheetDocument;
        }

        internal override SpecificWorksheet NewSheet()
        {

            var worksheetPart = Document.WorkbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData()); //new Worksheet();
            var workbookPart = Document.WorkbookPart;
            //var spreadsheetName = spreadsheet.Item1;

            var sid = workbookPart.GetIdOfPart(worksheetPart);
            Sheet newWorksheet = new Sheet() { Id = sid, SheetId = _sheetId, Name = "Sheet1" };
            _sheetId++;
            _sheets.Append(newWorksheet);
            workbookPart.Workbook.Save();

            var specificWorksheet = new OpenXmlWorksheet();
            specificWorksheet.Initialize(newWorksheet);
            specificWorksheet.Document = Document;

            // Actually add this sheet
            //DocumentFormat.OpenXml.Spreadsheet.Sheets.Add(specificWorksheet);

            return specificWorksheet;
            
        }

        internal override void RemoveFirstSheet()
        {
            
        }

        internal override bool RemoveSheet(SpecificWorksheet specificWorksheet)
        {
            //Document.WorkbookPart.Delete
            return true;
        }

        internal override bool RemoveSheet(string sheetName)
        {
            throw new NotImplementedException();
        }

        internal override bool Save()
        {
            // TODO: May need only one or both of these lines
            //Document.WorkbookPart.Workbook.Save();
            Document.Save();

            return true;
        }

        internal override bool Save(string newFilename)
        {
            Filename = newFilename;
            //Document.SaveAs(Filename);

            Document.Save();
            // Document.WorkbookPart.Workbook.Save();

            return true;
        }
    }
}
