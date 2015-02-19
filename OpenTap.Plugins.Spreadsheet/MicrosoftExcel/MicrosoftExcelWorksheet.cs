using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace OpenTap.Plugins.Spreadsheet.MicrosoftExcel
{
    public class MicrosoftExcelWorksheet : SpecificWorksheet
    {

        internal _Worksheet Worksheet { get; set; }
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
                return Worksheet.Columns.Count;
            }
        }

        internal override bool Dispose()
        {

            if (Worksheet != null)
            {
                Marshal.ReleaseComObject(Worksheet);
                Worksheet = null;
                return true;
            }

            return false;
        }

        internal override void Initialize(object actualWorksheet)
        {
            Worksheet = (_Worksheet)actualWorksheet;
        }

        internal override object GetCell(int x, int y)
        {
            return Worksheet.Cells[x, y];
        }

        internal override void SetCell(int x, int y, object value)
        {
            Worksheet.Cells[x, y] = value;
        }

        internal override void SetManyCells(object[,] cellValues, int rowOffset = 0, int columnOffset = 0)
        {
            int numberOfRows = cellValues.GetLength(0);
            int numberOfColumns = cellValues.GetLength(1);

            // Get range to quickly set many cells
            var rng = Worksheet.Range[Worksheet.Cells[1 + (VirtualSpreadsheet.FirstRowOfActualData - 1), 1], 
                Worksheet.Cells[numberOfRows + (VirtualSpreadsheet.FirstRowOfActualData - 1), numberOfColumns]];

            // Actually set the many cells
            rng.Value = cellValues;

        }

        internal override void AutoFitColumns()
        {
            Worksheet.Columns.AutoFit();
        }
    }
}