using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using OpenTap;

namespace Spreadsheet;

[Flags]
public enum Include
{
    [Display("Step Parameters", "Include the parameters of steps.", Order: 2)]
    StepParameters = 1 << 0,
    [Display("Plan Parameters", "Include the parameters of the test plan.", Order: 3)]
    PlanParameters = 1 << 1,
    [Display("Results", "Include all results.", Order: 4)]
    Results = 1 << 2,
    [Display("Run Id", "Include the run id of all steps and the test plan.", Order: 5)]
    RunId = 1 << 3,
}

[Display("Spreadsheet", "Save results in a spreadsheet.", "Text")]
public sealed class SpreadsheetResultListener : ResultListener
{
    private static readonly Dictionary<string, Array> EmptyResults = new();
    
    private readonly Dictionary<Guid, TestStepRun> _stepRuns = new();
    private readonly Dictionary<Guid, TestPlanRun> _planRuns = new();
    
    [Display("Filename", "The name of the spreadsheet file where the results are written.", Order: 1)]
    [FilePath(FilePathAttribute.BehaviorChoice.Open, "xls?")]
    public MacroString Path { get; set; } = new MacroString()
    {
        Text = "Results/<Date>-<Verdict>.xlsx"
    };
    
    [Display("Template Path", "Base the result sheet on an existing document.", Order: 1.2)]
    [FilePath(FilePathAttribute.BehaviorChoice.Open, "xls?")]
    public MacroString TemplatePath { get; set; } = new MacroString()
    {
        Text = ""
    };
    
    public enum FileExistsBehavior
    {
        [Display("Abort Testplan", Description:"Stop the test plan with an Error verdict if the file exists.")]
        FailWithError,
        [Display("Overwrite", Description:"Overwrite the existing file.")]
        OverwriteExisting,
        [Display("Append", Description:"Append new data to the end of the existing file.")]
        Append,
    }

    [Display("Overwrite Behavior", "What should happen if the file already exists.", Order: 1.1)]
    public FileExistsBehavior OverwriteBehavior { get; set; } = FileExistsBehavior.FailWithError; 

    [Display("Runs Summary Sheet", "Include a sheet with all runs and their parameters.", Order: 10)]
    public bool RunSummarySheetEnabled { get; set; } = true;

    [Display("Use Full Column Names",
        "Column headers will be fully qualified. E.g. the Verdict parameter from a step will become 'Step/Verdict'.",
        Order: 11)]
    public bool FullColumnHeaderName { get; set; } = true;

    [Display("Sheet Name", "The name of the sheet where results are written.", Order: 4)]
    public MacroString SheetName { get; set; } = new MacroString()
    {
        Text = "<ResultName>",
    };
    
    [Display("Results to Include", "The results that will be saved in the sheet.", Order: 4.1)]
    public Include Include { get; set; } = Include.PlanParameters | Include.StepParameters | Include.Results | Include.RunId;

    [Display("Open File After Run", "Opens the file in your default spreadsheet program after plan run.", Order: 99)]
    public bool OpenFile { get; set; } = true;
    
    private Spreadsheet? _spreadSheet;

    public SpreadsheetResultListener()
    {
        Name = "Spreadsheet";
    }
    
    public override void Open()
    {
        base.Open();
    }

    public override void OnTestPlanRunStart(TestPlanRun planRun)
    {
        _planRuns.Clear();
        _stepRuns.Clear();
        base.OnTestPlanRunStart(planRun);
        string templatePath = TemplatePath.Expand(planRun);
        string filePath = Path.Expand(planRun);
        if (File.Exists(filePath))
        {
            if (OverwriteBehavior == FileExistsBehavior.FailWithError)
            {
                throw new Exception($"The result sheet file '{filePath}' already exists.");
            } 
            if (OverwriteBehavior == FileExistsBehavior.OverwriteExisting)
            {
                File.Delete(filePath);
            }
        }

        bool isTemplate = !string.IsNullOrWhiteSpace(templatePath);
        if (isTemplate)
        {
            File.Copy(templatePath, filePath);
        }
        _spreadSheet = new Spreadsheet(filePath, RunSummarySheetEnabled, isTemplate);
        
        if (RunSummarySheetEnabled)
            _spreadSheet.PlanSheet.AddRows(CreateParameters("Plan", planRun), EmptyResults);
        _planRuns[planRun.Id] = planRun;
    }

    public override void OnTestPlanRunCompleted(TestPlanRun planRun, Stream logStream)
    {
        base.OnTestPlanRunCompleted(planRun, logStream);
        if (_spreadSheet is null)
        {
            return;
        }

        string path = _spreadSheet.FilePath;
        if (_spreadSheet.FileEmpty)
        {
            _spreadSheet.Dispose();
            Log.Warning("Result spreadsheet is empty. Deleting file.");
            File.Delete(path);
            return;
        }
        
        _spreadSheet.Dispose();
        planRun.PublishArtifact(path);
        Log.Info($"Created spreadsheet at '{path}'");

        if (OpenFile)
        {
            try
            {
                path = '"' + path + '"';
                Process? process = null;
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                {
                    Log.Info($"Windows detected, opening file \"{path}\"");
                    process = Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
                }
                else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
                {
                    Log.Info($"Linux detected, running \"xdg-open '{path}'\"");
                    process = Process.Start("xdg-open", path);
                }
                else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                {
                    Log.Info($"OSX detected, running \"open '{path}'\"");
                    process = Process.Start("open", path);
                }
                else
                {
                    Log.Warning("This platform does not support automatically opening the spreadsheet.");
                }

                if ((process?.HasExited ?? false) && process.ExitCode == 0)
                {
                    Log.Warning($"Process exited with code {process.ExitCode}");
                }
                
            }
            catch (Exception ex)
            {
                Log.Warning($"Something went wrong trying to open your file. {ex.Message}");
            }
        }
    }

    public override void OnTestStepRunStart(TestStepRun stepRun)
    {
        if (_spreadSheet is null)
        {
            throw new NullReferenceException();
        }
        
        _stepRuns.Add(stepRun.Id, stepRun);
        _planRuns[stepRun.Id] = _planRuns[stepRun.Parent];
        base.OnTestStepRunStart(stepRun);
    }

    public override void OnTestStepRunCompleted(TestStepRun stepRun)
    {
        if (_spreadSheet is null)
        {
            throw new NullReferenceException();
        }

        if (RunSummarySheetEnabled)
            _spreadSheet.PlanSheet.AddRows(CreateParameters("Step", stepRun), EmptyResults);
        
        base.OnTestStepRunCompleted(stepRun);
    }

    public override void OnResultPublished(Guid stepRunId, ResultTable result)
    {
        TestStepRun stepRun = _stepRuns[stepRunId];
        TestPlanRun planRun = _planRuns[stepRun.Id];
        SheetTab sheet = GetSheet(planRun, stepRun, result);
        sheet.AddRows(
            Include.HasFlag(Include.StepParameters) ? CreateParameters("Step", stepRun, result) : CreateIdParameters(stepRun, result),
            Include.HasFlag(Include.Results) ? CreateResults(result) : EmptyResults);
        base.OnResultPublished(stepRunId, result);
    }
    
    private SheetTab GetSheet(TestPlanRun planRun, TestStepRun? stepRun = null, ResultTable? table = null)
    {
        if (_spreadSheet is null)
        {
            throw new NullReferenceException();
        }

        return _spreadSheet.GetSheet(GetSheetName(planRun, stepRun, table));
    }
    
    private string GetSheetName(TestPlanRun planRun, TestStepRun? stepRun = null, ResultTable? table = null)
    {
        Dictionary<string, object> parameters = new Dictionary<string, object>();
        if (stepRun is not null)
        {
            parameters.Add("RunId", stepRun.Id);
            parameters.Add("StepId", stepRun.TestStepId);
            parameters.Add("StepName", stepRun.TestStepName);
            string stepTypeFull = stepRun.TestStepTypeName.Substring(0, stepRun.TestStepTypeName.IndexOf(','));
            parameters.Add("StepType", stepTypeFull.Substring(stepTypeFull.LastIndexOf('.') + 1));
            parameters.Add("StepTypeFull", stepTypeFull);
        }
        else
        {
            parameters.Add("RunId", planRun.Id);
            parameters.Add("StepId", planRun.Id);
            parameters.Add("StepName", planRun.TestPlanName);
            string stepTypeFull = typeof(TestPlan).FullName ?? nameof(TestPlan);
            parameters.Add("StepType", stepTypeFull.Substring(stepTypeFull.LastIndexOf('.') + 1));
            parameters.Add("StepTypeFull", stepTypeFull);
        }

        if (table is not null)
        {
            parameters.Add("ResultName", table.Name);
        }
        else
        {
            parameters.Add("ResultName", planRun.TestPlanName);
        }

        string sheetName = SheetName.Expand(planRun, stepRun?.StartTime, null, parameters);

        /*  From microsoft docs.
         * https://support.microsoft.com/en-gb/office/rename-a-worksheet-3f1f7148-ee83-404d-8ef0-9ff99fbad1f9#:~:text=Important%3A%20Worksheet%20names%20cannot%3A,Contain%20more%20than%2031%20characters.
         * Important:  Worksheet names cannot:
         * Be blank .
         * Contain more than 31 characters.
         * Contain any of the following characters: / \ ? * : [ ]
         *      For example, 02/17/2016 would not be a valid worksheet name, but 02-17-2016 would work fine.
         * Begin or end with an apostrophe ('), but they can be used in between text or numbers in a name.
         * Be named "History". This is a reserved word Excel uses internally.
         */
        sheetName = sheetName
            .Replace('/', '.')
            .Replace('\\', '.')
            .Replace('?', '.')
            .Replace(':', '.')
            .Replace('[', '.')
            .Replace(']', '.');

        while (sheetName.StartsWith("'"))
        {
            sheetName = sheetName.Substring(1);
        }

        while (sheetName.EndsWith("'"))
        {
            sheetName = sheetName.Substring(0, sheetName.Length - 1);
        }

        if (string.IsNullOrWhiteSpace(sheetName))
        {
            sheetName = ".";
        }

        if (sheetName.Length > 31)
        {
            sheetName = sheetName.Substring(0, 28) + "...";
        }

        if (sheetName == "History")
        {
            sheetName = ".History";
        }

        return sheetName;
    }

    private Dictionary<string, Array> CreateResults(ResultTable result)
    {
        if (FullColumnHeaderName)
        {
            return result.Columns.ToDictionary(c => "Results/" + c.Name, c => c.Data);
        }
        
        return result.Columns.ToDictionary(c => c.Name, c => c.Data);
    }

    private Dictionary<string, object> CreateIdParameters(TestRun run, ResultTable? table = null)
    {
        Dictionary<string, object> parameters = new Dictionary<string, object>();
        if (!Include.HasFlag(Include.RunId))
        {
            return parameters;
        }
        
        parameters.Add("RunId", run.Id);
        if (run is TestStepRun stepRun)
        {
            parameters.Add("ParentRunId", stepRun.Parent);
        }
        else
        {
            parameters.Add("ParentRunId", "");
        }
        
        if (table is not null)
        {
            parameters.Add("ResultName", table.Name);
        }

        return parameters;
    }

    private Dictionary<string, object> CreateParameters(string prefix, TestRun run, ResultTable? table = null)
    {
        Dictionary<string, object> parameters = CreateIdParameters(run, table);

        foreach (ResultParameter parameter in run.Parameters)
        {
            string name = parameter.Name;
            if (!string.IsNullOrWhiteSpace(parameter.Group))
                name = $"{parameter.Group}/{name}";
            if (FullColumnHeaderName)
                name = $"{prefix}/{name}";
            if (!parameters.ContainsKey(name))
                parameters.Add(name, parameter.Value);
        }
        }
        return parameters;
    }

    public override void Close()
    {
        base.Close();
    }
}