using System.Collections.Generic;

namespace OpenTap.Plugins.Spreadsheet
{
    public abstract class SpecificWorkbook
    {

        public List<SpecificWorksheet> Sheets { get; set; }

        public SpecificWorkbook()
        {
            Sheets = new List<SpecificWorksheet>();
        }

        /// <summary>
        /// Closes resources used by underlying workbook, may or may not be even used depending on the library.
        /// </summary>
        /// <returns></returns>
        internal abstract bool Close();

        /// <summary>
        /// Could be used for example to free COM objects.
        /// </summary>
        /// <returns>True if successful.</returns>
        internal abstract bool Dispose();

        /// <summary>
        /// Saves the current state of the cells to the corresponding filename.
        /// </summary>
        /// <returns></returns>
        internal abstract bool Save();

        /// <summary>
        /// Save As Overload to save all cells to a file.
        /// </summary>
        /// <param name="newFilename"></param>
        /// <returns></returns>
        internal abstract bool Save(string newFilename); // AKA SaveAs

        /// <summary>
        /// Filename that the underlying library is talking with or creating.
        /// </summary>
        internal abstract string Filename { get; set; }

        /// <summary>
        /// Creates a new worksheet using underlying library and adds it to the end of the list of worksheets.
        /// </summary>
        /// <returns></returns>
        internal abstract SpecificWorksheet NewSheet();

        /// <summary>
        /// Removes sheet from underlying library and removes from list of worksheets.
        /// </summary>
        /// <param name="specificWorksheet"></param>
        /// <returns></returns>
        internal abstract bool RemoveSheet(SpecificWorksheet specificWorksheet);

        /// <summary>
        /// Removes sheet by name from underlying library and removes from list of worksheets
        /// </summary>
        /// <param name="sheetName">The sheet name to remove</param>
        /// <returns></returns>
        internal abstract bool RemoveSheet(string sheetName);

        internal abstract void Initialize(string filename);

        internal abstract void RemoveFirstSheet();
    }
}
