using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

namespace OpenTap.Plugins.Spreadsheet.OpenXml
{
    public class OpenXmlSpecifics : Specifics
    {

        internal override void AutofitColumns()
        {

        }

        internal override void CurrentDomain_ProcessExit(object sender, EventArgs e)
        {

        }

        internal override void Clean()
        {

        }

        internal override void Initialize()
        {

        }

        internal override void CreateWorkbookFromExistingTemplate(string destinationFile, string templateFilename, ref List<VirtualSpreadsheet> sheets)
        {
            throw new NotImplementedException();
        }

        internal override void CreateNewWorkbook(string destinationFile)
        {
            File.Delete(destinationFile);

            // TODO: Probably should make a Create method on Workbook interface
            Workbook = new OpenXmlWorkbook(destinationFile);

            Workbook.Save(destinationFile);

            _isFirstSheet = true;
        }

        internal override void SimpleSave()
        {
            AutofitColumns();

            Workbook.Save();
        }

        internal override void CloseAndMaybeOpenFile(string destinationFilename, bool shouldOpenFile)
        {
            // TODO: Update

            string currentWorkbookPath = Workbook.Filename;
            Workbook.Close();

            // Check to see if filename changed for example verdict finished:
            if (currentWorkbookPath.Equals(destinationFilename) == false)
            {
                // Move old filename to new filename location
                File.Move(currentWorkbookPath, destinationFilename);
            }

            if (shouldOpenFile)
            {
                // Just make a new excel application instance with the saved workbook file.
                // This takes extra time, however it also ensures that the excel spreadsheet opens
                var processStartInfo = new ProcessStartInfo
                {
                    FileName = destinationFilename,
                    UseShellExecute = true,
                    RedirectStandardOutput = false,
                    RedirectStandardError = false,
                    RedirectStandardInput = false,
                };
                Process.Start(processStartInfo);
            }
        }

        internal override VirtualSpreadsheet AddNewSheetAfterExistingSheets(string resultTypeName, ResultTable result, ResultParameters parameters)
        {
            CurrentSheet = Workbook.NewSheet();


            List<string> sheetNames = new List<string>();

            foreach (var sheet in Workbook.Sheets)
            {
                sheetNames.Add(sheet.Name);
            }

            VirtualSpreadsheet virtualSheet = new VirtualSpreadsheet(
                CurrentSheet, resultTypeName, result,
                parameters, new Func<string, string>(x => FindUniqueWorksheetName(sheetNames, x)));

            SetSheetName(virtualSheet.Name);

            return virtualSheet;
        }

        internal override void SetSheetName(string name)
        {
            // TODO: Update
            List<string> sheetNames = new List<string>();
            foreach (var sheet in Workbook.Sheets)
            {
                sheetNames.Add(sheet.Name);
            }
            CurrentSheet.Name = ValidateSheetName(sheetNames, name);
        }

        internal override void CreateVirtualSpreadsheet(VirtualSpreadsheet virtualSpreadsheet)
        {
            // TODO: Update

            virtualSpreadsheet.RealWorksheet = CurrentSheet;
            virtualSpreadsheet.Name = CurrentSheet.Name;

            // Verify the last row of this sheet has data, if not this is not 
            // already a valid sheet and a new sheet will need to be created
            var firstCell = CurrentSheet.GetCell(VirtualSpreadsheet.ColumnNamesRow, 1);
            if (firstCell == null)
            {
                virtualSpreadsheet.IsValid = false;
                return;
            }

            // Go through all columns in existing sheet and extract existing information (template)
            for (int columnIndex = 1; columnIndex < CurrentSheet.ColumnsCount; columnIndex++)
            {
                string columnType = VirtualSpreadsheet.ColumnTypeStepParameterColumnName;

                // first columns are assumed to be x and y
                // TODO: This is potentially a problem
                if (columnIndex == 1 || columnIndex == 2)
                    columnType = VirtualSpreadsheet.ColumnTypeResultColumnName;

                var currentNameCell = CurrentSheet.GetCell(VirtualSpreadsheet.ColumnNamesRow, columnIndex);
                var columnName = currentNameCell as string;

                if ((columnType == null) || (columnName == null))
                {
                    if (columnIndex == 1)
                        virtualSpreadsheet.IsValid = false;
                    break;
                }

                int level = columnName.Length - columnName.TrimStart('>').Length;

                if (columnType.Equals(VirtualSpreadsheet.ColumnTypeResultColumnName))
                    virtualSpreadsheet._columns.Add(new VirtualColumn { Name = columnName, Level = 0, Unknown = false });
                else if (columnType.Equals(VirtualSpreadsheet.ColumnTypeStepParameterColumnName))
                {
                    if (columnName == VirtualSpreadsheet.StepVerdictColumnName)
                        virtualSpreadsheet._verdictColumn = columnIndex;

                    var resultColumn = new VirtualColumn { Name = columnName.TrimStart('>'), Level = level, Unknown = true };

                    virtualSpreadsheet._columns.Add(resultColumn);
                }
                else
                {
                    virtualSpreadsheet.IsValid = false;
                    return;
                }
            }
        }

        internal override bool AllowOverwrite { get; set; }
    }
}
