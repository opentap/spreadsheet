namespace OpenTap.Plugins.Spreadsheet
{
    public abstract class SpecificWorksheet
    {

        /// <summary>
        /// Called to provide linking between this generic Worksheet and the underlying actual worksheet.
        /// </summary>
        /// <param name="actualWorksheet"></param>
        internal abstract void Initialize(object actualWorksheet);

        /// <summary>
        /// Name of Worksheet.
        /// </summary>
        internal abstract string Name { get; set; }

        //internal abstract int ColumnsCount { get; set; }

        /// <summary>
        /// Gets array of cell objects in such a way where each cell can easily be modified of the actual cell in the chosen library.
        /// </summary>
        //internal abstract object[,] Cells { get; set; }

        /// <summary>
        /// Could be used for example to free COM objects.
        /// </summary>
        /// <returns>True if successful.</returns>
        internal abstract bool Dispose();

        /// <summary>
        /// Gets the number of columns in the underlying table.
        /// </summary>
        internal abstract int ColumnsCount { get; }
        //{
        //    get
        //    {
        //        //CurrentSheet.Columns.Count
        //        return Cells.GetLength(1);
        //    }
        //}

        //internal string[,] Range
        //{
        //    get
        //    {
        //        return Cells;
        //    }
        //}

        internal abstract void SetManyCells(object[,] cellValues, int rowOffset = 0, int columnOffset = 0);

        internal abstract object GetCell(int x, int y);

        internal abstract void SetCell(int x, int y, object value);

        internal abstract void AutoFitColumns();
    }
}
