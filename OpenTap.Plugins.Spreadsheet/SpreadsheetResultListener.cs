using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using OpenTap.Plugins.Spreadsheet.MicrosoftExcel;
using OpenTap.Plugins.Spreadsheet.OpenXml;

namespace OpenTap.Plugins.Spreadsheet
{

    [Display("Spreadsheet", Group: "Spreadsheet")]
    public class SpreadsheetResultListener : ResultListener, IFileResultStore
    {
        
        #region Settings
        [Display("Filename", "The name of the spreadsheet where the results are written.", Order: 1)]
        [FilePath(FilePathAttribute.BehaviorChoice.Open, "xls?")]
        public MacroString FilePath { get; set; }

        [Display("Template", "The spreadsheet template used as the starting point for the generated spreadsheet.", Order: 2)]
        [FilePath(FilePathAttribute.BehaviorChoice.Open, "xl??")] // match *.xltx, *.xlsx and *.xlsm
        public MacroString TemplatePath
        {
            get
            {
                return _templatePath;
            }
            set
            {
                _templatePath.Text = RemoveInvalidXmlChars(value.Text);
            }
        }

        [Display("Open Spreadsheet", "Open spreadsheet in default spreadsheet application after the test plan run is completed.", Order: 3)]
        public bool ShowSpreadsheet { get; set; }

        [Display("Arrange by Step Type", "Choose whether you want to arrange results by name or combined by step type.", "Advanced", Order: 4, Collapsed: true)]
        public bool ArrangeByStepType { get; set; }

        public enum SpreadsheetLibrary
        {
            [Display("Microsoft Excel")]
            MicrosoftExcel,
            [Display("OpenXML")]
            OpenXml,
        }

#if USE_MICROSOFT_EXCEL
        [XmlIgnore]
#endif
        [Display("Spreadsheet Library", "Which spreadsheet library to utilize.", "Advanced", Order: 5)]
        public SpreadsheetLibrary ChosenLibrary { get; set; }

        // TODO: Enable in a future version
        [Browsable(false)]
        public bool AllowOverwrite { get; set; }

        #endregion

        #region Private Members
        private List<VirtualSpreadsheet> _sheets = null;
        private Dictionary<string, VirtualSpreadsheet> _resultSheets = null;
        private Dictionary<TestStepRun, List<Tuple<List<int>, int, VirtualSpreadsheet>>> _stepVerdicts = null;
        private string _defaultFilePath = "Results\\<Date>-<Verdict>.xlsx";
        private MacroString _templatePath = null;
        private static readonly object _stepVerdictsLock = new object();
        private static Regex _invalidXmlChars = new Regex(
    @"(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F\uFEFF\uFFFE\uFFFF]",
    RegexOptions.Compiled);

        private TestPlanRun _planRun = null;
        private bool _isClosed = false;
        private Exception _ioException = null;
        private static object _multipleLock = new object();
        private static Mutex mutex = new Mutex(false, "ExcelMutex");

        private Specifics ActualSpreadsheet { get; set; }
        #endregion

        public SpreadsheetResultListener()
        {
            ChosenLibrary = SpreadsheetLibrary.OpenXml;
            Name = "Spreadsheet";

            _sheets = new List<VirtualSpreadsheet>();
            _resultSheets = new Dictionary<string, VirtualSpreadsheet>();
            _stepVerdicts = new Dictionary<TestStepRun, List<Tuple<List<int>, int, VirtualSpreadsheet>>>();
            _templatePath = new MacroString();
            FilePath = new MacroString() { Text = _defaultFilePath };
            ArrangeByStepType = true;

            this.Rules.Add(() => ValidateFile(), "Invalid file or unsupported spreadsheet file format. Supported format(s): xlsx, xlsm.", nameof(FilePath));
            this.Rules.Add(() => ValidateTemplateFile(), "Invalid file or unsupported template file format. Supported format(s): xltx, xlsx, xlsm.", nameof(TemplatePath));

            ShowSpreadsheet = false;
        }

        #region Private Methods
        private bool ValidateFile()
        {
            return ValidateReportFile(FilePath, new List<string>() { "xlsx", "xlsm" });
        }

        private bool ValidateTemplateFile()
        {
            if (String.IsNullOrEmpty(_templatePath)) // Template file is optional so it can be left empty
            {
                return true;
            }
            else
            {
                return ValidateFile(_templatePath, new List<string>() { "xltx", "xlsx", "xlsm" });
            }
        }

        public static bool ValidateFile(string filePath, List<string> requiredFileExtensions)
        {
            if (File.Exists(filePath))
            {
                var currentFileExtension = Path.GetExtension(filePath).TrimStart('.');
                foreach (string requiredFileExtension in requiredFileExtensions)
                {
                    if (String.Equals(currentFileExtension, requiredFileExtension, StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        public static bool ValidateReportFile(string expandedFilePath, List<string> requiredFileExtensions)
        {
            // Note: The file itself doesn't have to necessarily exist just yet, as it will be created later.
            // We will check and ensure that the path isn't empty and that file has the correct extension
            if (String.IsNullOrWhiteSpace(expandedFilePath) == false && ContainsInvalidPathCharacters(expandedFilePath) == false)
            {
                var currentFileExtension = String.Empty;
                try
                {
                    // check whether the file extension is correct
                    currentFileExtension = Path.GetExtension(expandedFilePath).TrimStart('.');
                }
                catch
                {
                    return false;
                }

                foreach (string requiredFileExtension in requiredFileExtensions)
                {
                    if (String.Equals(currentFileExtension, requiredFileExtension, StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool ContainsInvalidPathCharacters(string filePath)
        {
            for (var i = 0; i < filePath.Length; i++)
            {
                int c = filePath[i];

                if (c == '\"' || c == '<' || c == '>' || c == '|' || c == '*' || c == '?' || c < 32)
                    return true;
            }

            return false;
        }

        public static string RemoveInvalidXmlChars(string text)
        {
            if (string.IsNullOrEmpty(text))
                return String.Empty;
            return _invalidXmlChars.Replace(text, "");
        }

        private void CollectFinalResults()
        {
            // Collect the final result in case additional OnResultPublished calls have been made:
            int totalVerdicts = 0;
            lock (_stepVerdictsLock)
            {
                foreach (var stepRun in _stepVerdicts.Keys)
                {
                    foreach (var stepVerdict in _stepVerdicts[stepRun])
                    {
                        stepVerdict.Item3.SetVirtualTable();
                        foreach (int currentRow in stepVerdict.Item1)
                        {
                            // Set cell of worksheet in corresponding stepRun to that stepRun's verdict
                            string searchCriteria = ArrangeByStepType ? stepRun.TestStepTypeName : stepRun.TestStepId.ToString();
                            var sheet = _resultSheets[searchCriteria];

                            sheet.VirtualTable[currentRow - (VirtualSpreadsheet.FirstRowOfActualData - 1) - 1, stepVerdict.Item2 - 1] = stepRun.Verdict.ToString();
                        }

                        totalVerdicts++;
                    }
                }
            }
        }

        private void SaveVirtualTableToSpreadsheet()
        {
            // Add Virtual Spreadsheet Table to Real Spreadsheet Table
            foreach (var sheet in _resultSheets.Values)
            {
                sheet.RealWorksheet.SetManyCells(sheet.VirtualTable);
            }

            ActualSpreadsheet.SimpleSave();
        }

        #endregion


        #region TAP Events
        public override void OnTestPlanRunStart(TestPlanRun planRun)
        {
            if (ChosenLibrary == SpreadsheetLibrary.MicrosoftExcel)
            {
                ActualSpreadsheet = new MicrosoftExcelSpecifics();
            }
            else if (ChosenLibrary == SpreadsheetLibrary.OpenXml)
            {
                ActualSpreadsheet = new OpenXmlSpecifics();
            }

            ActualSpreadsheet.AllowOverwrite = AllowOverwrite;

            _planRun = planRun;
            _ioException = null;
            ActualSpreadsheet.Clean();
            ActualSpreadsheet.Initialize();

            // Make current Workbook (_workbook) associated with the user-defined template (if provided), 
            // or create a new workbook and save to user's destination file
            string destinationFile = Path.GetFullPath(FilePath.Expand(planRun));
            Directory.CreateDirectory(Path.GetDirectoryName(destinationFile));

            if (File.Exists(_templatePath))
            {
                string expandedTemplateFilePath = Path.GetFullPath(_templatePath.Expand(planRun));
                ActualSpreadsheet.CreateWorkbookFromExistingTemplate(destinationFile, expandedTemplateFilePath, ref _sheets);
            }
            else
            {
                try
                {
                    ActualSpreadsheet.CreateNewWorkbook(destinationFile);
                }
                catch(IOException ioex)
                {
                    _ioException = ioex;
                }             
            }
        }

        public override void OnTestStepRunStart(TestStepRun stepRun)
        {
            base.OnTestStepRunStart(stepRun);

            if (_ioException != null)
            {
                ActualSpreadsheet.Clean();
                
                Log.Warning("The report file name for '{0}' conflicts with another '{1}'. Please specify a unique report file name.", this.Name, this.GetType().Name);
                throw new Exception(_ioException.Message);
            }

            // Create new instance in dictionary for this steprun type.
            lock (_stepVerdictsLock)
            {
                _stepVerdicts[stepRun] = new List<Tuple<List<int>, int, VirtualSpreadsheet>>();
            }
        }

        // This function can actually get called after TestStepRunCompleted!
        public override void OnResultPublished(Guid stepRunId, ResultTable result)
        {
            OnActivity();
            try
            {
                mutex.WaitOne();
                var stepRun = _stepVerdicts.Keys.Where(x => x.Id == stepRunId).FirstOrDefault();
                if (stepRun == null)
                {
                    // This stepRun does not exist for some reason?
                    // TODO: Figure out something better to do here, this may not be possible anyway
                    return;
                }
                var parameters = stepRun.Parameters;

                VirtualSpreadsheet destinationSheet = null;

                string resultTypeName = ArrangeByStepType ? result.Name : stepRun.TestStepName;
                string searchCriteria = ArrangeByStepType ? stepRun.TestStepTypeName : stepRun.TestStepId.ToString();

                // If dictionary contains a sheet by this result type
                // For example, 3 test steps of RampResults would use the same key of RampResults
                if (_resultSheets.ContainsKey(searchCriteria))
                {
                    // We found a match
                    destinationSheet = _resultSheets[searchCriteria];

                    // Call Matches to add any incidental columns
                    destinationSheet.Matches(resultTypeName, result, parameters);
                }
                else // Key already exists so let's add a sheet
                {

                    // Search through current sheets to see if we find a match
                    foreach (var sheet in _sheets)
                    {
                        if (sheet.Matches(resultTypeName, result, parameters))
                        {
                            // If the parameters are already in place just go ahead
                            _resultSheets[searchCriteria] = sheet;
                            destinationSheet = sheet;
                            break;
                        }
                    }

                    // First time, we have to create a new sheet
                    if (destinationSheet == null)
                    {

                        var sheet = ActualSpreadsheet.AddNewSheetAfterExistingSheets(resultTypeName, result, parameters);

                        _sheets.Add(sheet);
                        _resultSheets[searchCriteria] = sheet;
                        destinationSheet = sheet;
                    }
                }

                // Fixes key not present in dictionary
                if (_stepVerdicts.ContainsKey(stepRun) == false)
                    _stepVerdicts[stepRun] = new List<Tuple<List<int>, int, VirtualSpreadsheet>>();

                // We have a sheet so populate the results from TAP into Spreadsheet cells and store them in our dictionary
                _stepVerdicts[stepRun].Add(destinationSheet.AddValuesToVirtualTable(stepRun, result, parameters));

                // TODO: Should collect data periodically
                //CollectFinalResults();

                ActualSpreadsheet.SimpleSave();
                mutex.ReleaseMutex();
            }
            catch (OutOfMemoryException)
            {
                Log.Warning("Warning: Not enough memory. Some data may not be available in the generated report");
            }
            catch(Exception ex)
            {
                Log.Error(ex);
            }
        }

        public override void OnTestPlanRunCompleted(TestPlanRun planRun, Stream logStream)
        {
            try
            {
                mutex.WaitOne();
                CollectFinalResults();

                SaveVirtualTableToSpreadsheet();
                    
                ActualSpreadsheet.CloseAndMaybeOpenFile(Path.GetFullPath(FilePath.Expand(planRun)), ShowSpreadsheet);
                mutex.ReleaseMutex();
            }
            catch (OutOfMemoryException)
            {
                Log.Warning("Warning: Not enough memory. Some data may not be available in the generated report");
            }
            catch (Exception ex)
            {
                Log.Error(ex.Message);
            }
            finally
            {
                // Ensure prior sheets, if any are properly cleaned up before starting a new test plan
                _sheets.ForEach(sheet =>
                {
                //if (sheet.RealWorksheet != null) { Marshal.ReleaseComObject(sheet.RealWorksheet); }
                sheet.RealWorksheet.Dispose();
                });
                _sheets.Clear();

                ActualSpreadsheet.Clean();

                _resultSheets.Clear();

                // Clear prior testrun associated worksheets (and verdict rows, columns)
                lock (_stepVerdictsLock)
                {
                    _stepVerdicts.Clear();
                }
            }
        }

        public override void Open()
        {
            _isClosed = false;

            base.Open();
        }

        public override void Close()
        {
            base.Close();

            _isClosed = true;
        }
#endregion

        public override string ToString()
        {
            string filepath = _planRun == null ? FilePath : FilePath.Expand(_planRun);
            if (_isClosed)
                _planRun = null;
            return string.Format("{0} ({1})", Name, filepath);
        }

#region IFileResultStore
        public string DefaultExtension
        {
            get { return "xlsx"; }
        }

        string IFileResultStore.FilePath
        {
            get { return FilePath; }
            set { this.FilePath.Text = value; }
        }
#endregion
    }
}
