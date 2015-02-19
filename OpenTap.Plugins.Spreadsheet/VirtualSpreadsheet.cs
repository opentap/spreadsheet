using System;
using System.Collections.Generic;
using System.Linq;

namespace OpenTap.Plugins.Spreadsheet
{

    #region Structs
    public struct VirtualColumn
    {
        public string Name;
        public int Level;
        public bool Unknown;
    }
    #endregion

    public class VirtualSpreadsheet
    {

        #region SpreadsheetSpecific
        internal SpecificWorksheet RealWorksheet { get; set; }
        #endregion

        public object[,] VirtualTable
        {
            get
            {
                return _virtualTable;
            }
        }

        // Current sheet name
        public string Name { get; set; }

        public bool IsValid { get; internal set; }

        public List<VirtualColumn> _columns = new List<VirtualColumn>();

        public int _row; // Current row of data being added, starts at row _kFirstRowOfActualData
        public int _verdictColumn;

        //internal static int ResultColumnTypeRow = 1;
        internal static int ColumnNamesRow = 1;
        internal static int FirstRowOfActualData = 2;

        public static string ColumnTypeResultColumnName = "Result"; // Row 2
        public static string ColumnTypeStepParameterColumnName = "Step Parameter"; // Row 2

        public static string StepNameColumnName = "Step Name"; // Column C
        public static string StepVerdictColumnName = "Step Verdict"; // Column D

        public object[,] _virtualTable = null;
        public bool _isVirtualTableSet = false;
        public List<List<object>> virtualTableAsList = null;
        #region Constructors

        /// <summary>
        /// Used to create a Worksheet from an existing sheet like from a template.
        /// </summary>
        /// <param name="sheet"></param>
        internal VirtualSpreadsheet(Specifics specifics)
        {

            virtualTableAsList = new List<List<object>>();

            IsValid = true;

            _row = VirtualSpreadsheet.FirstRowOfActualData;
            _verdictColumn = 0;

            specifics.CreateVirtualSpreadsheet(this);

        }

        /// <summary>
        /// Only used when creating a new sheet.
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="resultTypeName"></param>
        /// <param name="result"></param>
        /// <param name="parameters"></param>
        /// <param name="findName"></param>
        internal VirtualSpreadsheet(SpecificWorksheet realWorksheet, string resultTypeName, ResultTable result, ResultParameters parameters, Func<string, string> findName)
        {

            RealWorksheet = realWorksheet;

            virtualTableAsList = new List<List<object>>();
            Name = findName(resultTypeName);

            // Adds columns to _worksheet
            foreach (var dim in result.Columns)
            {
                MaybeAddColumn(dim.Name, dim.Name);
            }

            // Add two additional columns of Step Name and Step Verdict
            MaybeAddColumn(StepNameColumnName, StepNameColumnName);
            MaybeAddColumn(StepVerdictColumnName, StepVerdictColumnName);

            // The verdict column is set to the last column here as we just added it above
            _verdictColumn = _columns.Count;

            // Set row to first row
            _row = FirstRowOfActualData;

            // Add additional columns that are known pre-runtime
            foreach (var param in parameters)
                MaybeAddColumn(param.Name, param.Name, param.ParentLevel);
        }
        #endregion

        #region Column Operations
        /// <summary>
        /// Adds a column to the current _worksheet if it's not already been added and exist in _columns.
        /// </summary>
        /// <param name="displayName"></param>
        /// <param name="name"></param>
        /// <param name="columnType"></param>
        /// <param name="level"></param>
        private void MaybeAddColumn(string displayName, string name, int level = 0)
        {

            bool columnAlreadyExists = _columns.Exists(Column => (Column.Name == name) && (Column.Level == level) && (!Column.Unknown));

            if (columnAlreadyExists == false)
            {
                // Ensure that this is not an unknown column
                int unknownColumnIndex = _columns.FindIndex(Column => (Column.Name == displayName) && Column.Unknown);
                if (unknownColumnIndex >= 0)
                {
                    _columns[unknownColumnIndex] = new VirtualColumn { Name = name, Level = level };
                    return;
                }

                // This column does not exist and is not unknown, let's add it
                _columns.Add(new VirtualColumn { Name = name, Level = level, Unknown = false });

                // Sets row 1 to either Result or Step Parameter
                //RealWorksheet.SetCell(ResultColumnTypeRow, _columns.Count, resultColumnTypeString);

                // For hiearchy, repeat > character and append to front of display name (row 3)
                string levelCharactersRepeated = String.Empty;
                if (level > 0)
                    levelCharactersRepeated = new string('>', level);

                // Append to row 3
                string cellDisplayName = levelCharactersRepeated + displayName;
                RealWorksheet.SetCell(ColumnNamesRow, _columns.Count, cellDisplayName);
            }
        }

        internal bool Matches(string resultTypeName, ResultTable result, ResultParameters parameters)
        {
            if (Name != resultTypeName)
                return false;

            for (int i = 0; i < result.Columns.Length; i++)
            {
                if (result.Columns[i].Name != _columns[i].Name)
                    return false;
            }

            parameters.ToList().ForEach(Param => MaybeAddColumn(Param.Name, Param.Name, Param.ParentLevel));

            return true;
        }
        #endregion

        #region Virtual Table Operations
        internal void SetVirtualTable()
        {
            if (_isVirtualTableSet == true)
                return;

            _isVirtualTableSet = true;

            int longestColumn = 1;
            foreach (var columns in virtualTableAsList)
            {
                if (columns.Count > longestColumn)
                    longestColumn = columns.Count;
            }

            var virtualTable = new object[virtualTableAsList.Count, longestColumn];

            for (int rowIndex = 0; rowIndex < virtualTableAsList.Count; ++rowIndex)
            {
                var row = virtualTableAsList[rowIndex];
                for (int columnIndex = 0; columnIndex < row.Count; ++columnIndex)
                {
                    var column = row[columnIndex];
                    virtualTable[rowIndex, columnIndex] = virtualTableAsList[rowIndex][columnIndex];
                }
            }

            _virtualTable = virtualTable;
        }

        internal Tuple<List<int>, int, VirtualSpreadsheet> AddValuesToVirtualTable(TestStepRun stepRun, ResultTable result, ResultParameters parameters)
        {

            // VerdictColumn is 1-index based, if the column is not already known, add it
            if (_verdictColumn <= 0)
            {
                MaybeAddColumn(StepVerdictColumnName, StepVerdictColumnName);
                _verdictColumn = _columns.Count;
            }

            List<int> verdictRows = new List<int>();

            for (int resultRowIndex = 0; resultRowIndex < result.Rows; resultRowIndex++)
            {
                List<object> rowData = new List<object>();
                verdictRows.Add(_row);

                for (int i = 0; i < _columns.Count; i++)
                {
                    object cellValue = null;

                    if (result.Columns.Count(x => x.Name == _columns[i].Name) > 0)
                    {
                        if (result.Columns.Length > i)
                            cellValue = result.Columns[i].Data.GetValue(resultRowIndex);
                    }
                    else
                    {
                        if (_columns[i].Name == StepNameColumnName)
                        {
                            cellValue = stepRun.TestStepName;
                        }
                        else if (_columns[i].Name == StepVerdictColumnName)
                        {
                            cellValue = "Pending";
                        }
                        else
                        {
                            int parameterIndex = parameters.ToList().FindIndex(p => (p.Name == _columns[i].Name) && (p.ParentLevel == _columns[i].Level));
                            if (parameterIndex >= 0)
                                cellValue = parameters[parameterIndex].Value;
                        }
                    }

                    rowData.Add(cellValue);
                }

                // Add a row to our known cells
                virtualTableAsList.Add(rowData);

                _row++;
            }

            return new Tuple<List<int>, int, VirtualSpreadsheet>(verdictRows, _verdictColumn, this);
        }
        #endregion
    }
}
