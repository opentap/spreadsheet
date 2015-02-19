using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace OpenTap.Plugins.Spreadsheet.MicrosoftExcel
{
    internal class MicrosoftExcelSpecifics : Specifics
    {

        internal override bool AllowOverwrite { get; set; }

        private Application _microsoftExcelApplication = null;
        internal Application MicrosoftExcelApplication
        {
            get
            {
                return _microsoftExcelApplication;
            }
        }

        internal override void AutofitColumns()
        {
            // Make columns autofit
            foreach (var worksheet in Workbook.Sheets)
            {
                worksheet.AutoFitColumns();
            }
        }

        internal override void CurrentDomain_ProcessExit(object sender, EventArgs e)
        {
            var workbook = Workbook as MicrosoftExcelWorkbook;
            if (workbook == null || workbook.Workbook == null)
                return;

            AutofitColumns();

            Workbook.Save();

            Workbook.Close();
            _microsoftExcelApplication?.Quit();
        }

        internal override void Clean()
        {

            // Ensure prior workbook is cleaned up before using
            if (Workbook != null)
                Workbook.Dispose();

            // Ensure prior Microsoft Excel application is cleaned up before using
            if (_microsoftExcelApplication != null) { Marshal.ReleaseComObject(_microsoftExcelApplication); }

            _microsoftExcelApplication = null;

            // Ensure COM objects are cleaned up
            GC.Collect();
        }

        internal override void Initialize()
        {
            // Create a new Microsoft Excel application
            // TODO: Add Instantiation
            _microsoftExcelApplication = new Application();
            _microsoftExcelApplication.Visible = false;
        }

        internal override void CreateWorkbookFromExistingTemplate(string destinationFile, string templateFilename, ref List<VirtualSpreadsheet> sheets)
        {

            // Use a temporary file until the very end
            destinationFile = GetTempFileName(destinationFile);

            Workbook = new MicrosoftExcelWorkbook(_microsoftExcelApplication);
            var extension = Path.GetExtension(templateFilename);

            if (extension.Equals(".xltx", StringComparison.OrdinalIgnoreCase))
            {
                Workbook.Initialize(templateFilename);
                Workbook.Save(destinationFile);
            }
            else
            {
                File.Copy(templateFilename, destinationFile, true);
                Workbook.Initialize(destinationFile);
                Workbook.Filename = destinationFile;
            }

            foreach (MicrosoftExcelWorksheet sheet in Workbook.Sheets)
            {
                CurrentSheet = sheet;
                // Attempt to load existing worksheet from template
                VirtualSpreadsheet resultSheet = new VirtualSpreadsheet(this);

                // Requirement is simply that there is at least 3 rows
                if (resultSheet.IsValid == false)
                    continue; // this is invalid, try another sheet

                sheets.Add(resultSheet);
            }

            _isFirstSheet = false;
        }

        private string GetTempFileName(string originalFilename)
        {
            var targetFilename = String.Empty;
            do
            {
                targetFilename = Path.Combine(Path.GetDirectoryName(originalFilename), Guid.NewGuid().ToString() + ".xlsx");
            } while (File.Exists(targetFilename));
            return targetFilename;
        }

        internal override void CreateNewWorkbook(string destinationFile)
        {

            // Use a temporary file until the very end
            destinationFile = GetTempFileName(destinationFile);
            
            Workbook = new MicrosoftExcelWorkbook(_microsoftExcelApplication);

            Workbook.Save(destinationFile);

            // Ensure there is only a single sheet (Microsoft Excel by default creates 3 sheets)
            for (int i = Workbook.Sheets.Count; i > 1; i--)
                (Workbook.Sheets[i] as MicrosoftExcelWorksheet).Worksheet.Delete();

            _isFirstSheet = true;
        }

        internal override void SimpleSave()
        {

            AutofitColumns();

            // Sort by Sample Number, not quite working yet
            //foreach (_Worksheet worksheet in _workbook.Sheets)
            //{
            //    Range r = _microsoftExcelApplication.Rows[ResultSheet._kFirstRowOfActualData.ToString() + ":" + worksheet.Rows.Count.ToString()];
            //    r.Select();
            //    Range selectionRange = _microsoftExcelApplication.Selection;
            //    selectionRange.Sort(Key1: _microsoftExcelApplication.Range["A3"], Order1: XlSortOrder.xlAscending, Header: XlYesNoGuess.xlGuess, OrderCustom: 1, MatchCase: false, Orientation: XlSortOrientation.xlSortRows, DataOption1: XlSortDataOption.xlSortNormal);
            //}

            // Update workbook with the saved values.
            // TODO: Could call this save function more often in case of crash or something during a long run
            Workbook.Save(); // Save periodically
        }

        internal override void CloseAndMaybeOpenFile(string destinationFilename, bool shouldOpenFile)
        {
            string currentWorkbookPath = Workbook.Filename;
            Workbook.Close();
            _microsoftExcelApplication.Quit();

            // Check to see if filename changed for example verdict finished:
            //if (currentWorkbookPath.Equals(destinationFilename) == false) // should now always be false because now using temp file
            //{
            //    // Move old filename to new filename location
            //    // TODO: Update with a check if should overwrite

            //    if (File.Exists(destinationFilename))
            //        File.Delete(destinationFilename);

            //    //File.Copy(currentWorkbookPath, destinationFilename);
            //    //File.Move(currentWorkbookPath, destinationFilename);


            //}

            // TODO: add a overwrite option here
            if (AllowOverwrite == false && File.Exists(destinationFilename))
            {
                destinationFilename = GetNewFileName(destinationFilename);
            }
            try
            {
                if (AllowOverwrite)
                {
                    File.Delete(destinationFilename);
                }
                File.Move(currentWorkbookPath, destinationFilename);
            }
            catch (Exception)
            {

            }


            if (shouldOpenFile)
            {
                // Just make a new Microsoft Excel application instance with the saved workbook file.
                // This takes extra time, however it also ensures that the Microsoft Excel spreadsheet opens
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

        private string GetNewFileName(string oldFileName)
        {
            var counter = 2;
            var extension = Path.GetExtension(oldFileName);
            var directory = Path.GetDirectoryName(oldFileName);
            var filename = Path.GetFileNameWithoutExtension(oldFileName);

            var newFileName = Path.Combine(directory,
               String.Format("{0}{1}{2}", filename, counter, extension));

            while (File.Exists(newFileName))
            {
                counter++;
                newFileName = Path.Combine(directory,
                  String.Format("{0}{1}{2}", filename, counter, extension));
            }

            return newFileName;
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

            if (_isFirstSheet && Workbook.Sheets.Count == 1) // worksheets.Count == 2)
            {
                Workbook.RemoveSheet(Workbook.Sheets[0]);
                _isFirstSheet = false;
            }

            return virtualSheet;
        }

        internal override void SetSheetName(string name)
        {
            List<string> sheetNames = new List<string>();
            foreach (var sheet in Workbook.Sheets)
            {
                sheetNames.Add(sheet.Name);
            }
            CurrentSheet.Name = ValidateSheetName(sheetNames, name);
        }

        internal override void CreateVirtualSpreadsheet(VirtualSpreadsheet virtualSpreadsheet)
        {

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

                var currentNameCell = CurrentSheet.GetCell(VirtualSpreadsheet.ColumnNamesRow, columnIndex);
                var columnName = currentNameCell as string;

                if (columnName == null)
                {
                    if (columnIndex == 1)
                        virtualSpreadsheet.IsValid = false;
                    break;
                }

                int level = columnName.Length - columnName.TrimStart('>').Length;

                if (columnName == VirtualSpreadsheet.StepVerdictColumnName)
                    virtualSpreadsheet._verdictColumn = columnIndex;

                var resultColumn = new VirtualColumn { Name = columnName.TrimStart('>'), Level = level, Unknown = true };

                virtualSpreadsheet._columns.Add(resultColumn);
            }
        }
    }
}
