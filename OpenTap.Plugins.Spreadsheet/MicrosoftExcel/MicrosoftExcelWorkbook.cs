using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace OpenTap.Plugins.Spreadsheet.MicrosoftExcel
{
    public class MicrosoftExcelWorkbook : SpecificWorkbook
    {

        public Application MicrosoftExcelApplication { get; set; }
        public _Workbook Workbook { get; set; }
        public MicrosoftExcelWorkbook(Application application)
        {
            MicrosoftExcelApplication = application;
            var workbooks = MicrosoftExcelApplication.Workbooks;
            Workbook = workbooks.Add();
        }

        internal override string Filename { get; set; }

        internal override bool Close()
        {
            Workbook.Close(false);
            return true;
        }

        internal override bool Dispose()
        {
            if (Workbook != null)
            {
                Marshal.ReleaseComObject(Workbook);
                Workbook = null;
            }

            return true;
        }

        internal override void Initialize(string filename)
        {
            var workbooks = MicrosoftExcelApplication.Workbooks;
            Workbook = workbooks.Open(filename);

            foreach (var sheet in Workbook.Sheets)
            {
                SpecificWorksheet specificWorksheet = new MicrosoftExcelWorksheet();
                specificWorksheet.Initialize(sheet);

                // Actually add this sheet
                Sheets.Add(specificWorksheet);
            }
        }

        internal override SpecificWorksheet NewSheet()
        {
            var worksheets = Workbook.Sheets;
            var latestSheet = Workbook.Sheets[Workbook.Sheets.Count];
            var newWorksheet = worksheets.Add(After: latestSheet);

            SpecificWorksheet specificWorksheet = new MicrosoftExcelWorksheet();
            specificWorksheet.Initialize(newWorksheet);

            // Actually add this sheet
            Sheets.Add(specificWorksheet);

            return specificWorksheet;
        }

        internal override void RemoveFirstSheet()
        {
            Workbook.Sheets.Delete();
        }

        internal override bool RemoveSheet(SpecificWorksheet specificWorksheet)
        {
            // TODO: This is not exactly the best to do this.
            (specificWorksheet as MicrosoftExcelWorksheet).Worksheet.Delete();
            return true;
        }

        internal override bool RemoveSheet(string sheetName)
        {
            throw new NotImplementedException();
        }

        internal override bool Save()
        {
            Workbook.Save();

            return true;
        }

        internal override bool Save(string newFilename)
        {
            Filename = newFilename;
            Workbook.SaveAs(Filename);
            
            return true;
        }
    }
}
