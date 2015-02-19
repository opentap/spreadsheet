using System;
using System.Collections.Generic;

namespace OpenTap.Plugins.Spreadsheet
{
    public abstract class Specifics
    {

        internal SpecificWorksheet CurrentSheet { get; set; }

        internal SpecificWorkbook Workbook { get; set; }

        protected const int _kMaxSheets = 1000;
        protected bool _isFirstSheet = true;
        protected static Dictionary<string, string> _invalidMicrosoftExcelSheetCharactersDictionary = new Dictionary<string, string>()
        {
            { @"\", "" },
            { ":", "" },
            { "/", "" },
            { "?", "" },
            { "*", "" },
            { "[", "" },
            { "]", "" },
        };

        internal Specifics()
        {
            AppDomain.CurrentDomain.ProcessExit += new EventHandler(CurrentDomain_ProcessExit);
        }

        internal abstract void CurrentDomain_ProcessExit(object sender, EventArgs e);

        internal abstract void AutofitColumns();

        internal abstract void Clean();

        internal abstract void Initialize();

        internal abstract void CreateWorkbookFromExistingTemplate(string destinationFile, string templateFilename, ref List<VirtualSpreadsheet> sheets);

        internal abstract void CreateNewWorkbook(string destinationFile);

        internal abstract void SimpleSave();

        internal abstract void CloseAndMaybeOpenFile(string destinationFilename, bool shouldOpenFile);

        internal abstract VirtualSpreadsheet AddNewSheetAfterExistingSheets(string resultTypeName, ResultTable result, ResultParameters parameters);

        internal abstract void SetSheetName(string name);

        internal abstract void CreateVirtualSpreadsheet(VirtualSpreadsheet virtualSpreadsheet);

        internal abstract bool AllowOverwrite { get; set; }

        internal string ValidateSheetName(IList<string> sheetNames, string potentialSheetName)
        {
            if (String.IsNullOrEmpty(potentialSheetName))
            {
                return FindUniqueWorksheetName(sheetNames, "DEFAULT");
            }

            // Remove unsupported characters
            foreach (var character in _invalidMicrosoftExcelSheetCharactersDictionary)
                potentialSheetName = potentialSheetName.Replace(character.Key, character.Value);

            // Trim characters off until a unique name can be found
            string currentString = FindUniqueWorksheetName(sheetNames, potentialSheetName);
            while (true)
            {
                if (potentialSheetName.Length == 1)
                    break;

                if (currentString.Length > 31)
                {

                    potentialSheetName = potentialSheetName.Substring(0, potentialSheetName.Length - 1);
                    currentString = FindUniqueWorksheetName(sheetNames, potentialSheetName);
                    continue;
                }
                else
                    break;
            }

            return FindUniqueWorksheetName(sheetNames, potentialSheetName);
        }

        internal string FindUniqueWorksheetName(IList<string> sheetNames, string arg)
        {

            int count = 0;
            foreach (var sheetName in sheetNames)
            {
                if (sheetName == arg)
                    count++;
            }

            if (count == 0)
                return arg;

            for (int i = 0; i < _kMaxSheets; i++)
            {
                string newName = String.Format("{0} {1}", arg, i);

                int newNameCount = 0;
                foreach (var sheetName in sheetNames)
                {
                    if (sheetName == arg)
                        newNameCount++;
                }

                if (count == 0)
                    return newName;
            }
            throw new Exception("Failed to find unique name for worksheet.");
        }

    }
}
