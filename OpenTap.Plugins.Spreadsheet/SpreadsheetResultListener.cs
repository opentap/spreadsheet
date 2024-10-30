using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using DocumentFormat.OpenXml.Presentation;
using OpenTap;

namespace Spreadsheet;

[Flags]
public enum Include
{
    [Display("All", "Include everything.", Order: 0)]
    All = 0b111111,
    [Display("None", "Don't include anything.", Order: 1)]
    None = 0,
    [Display("Step parameters", "Include the parameters of steps.", Order: 2)]
    StepParameters = 1 << 0,
    [Display("Plan parameters", "Include the parameters of the test plan.", Order: 3)]
    PlanParameters = 1 << 1,
    [Display("Results", "Include all results.", Order: 4)]
    Results = 1 << 2,
    [Display("Run id", "Include the run id of all steps and the test plan.", Order: 5)]
    RunId = 1 << 3,
    [Display("Column type prefix", "Include the prefix of columns. Example: 'Step/Verdict' => 'Verdict'.", Order: 6)]
    ColumnTypePrefix = 1 << 4,
    [Display("Test plan sheet", "Include a first sheet with plan parameters and step parameters for steps without results.", Order: 7)]
    TestPlanSheet = 1 << 5,
}

[Display("Spreadsheet", "Save results in a spreadsheet.", "Database")]
public sealed class SpreadsheetResultListener : ResultListener
{
    private static readonly Dictionary<string, Array> EmptyResults = new();
    
    private readonly Dictionary<Guid, TestStepRun> _stepRuns = new();
    private readonly Dictionary<Guid, TestPlanRun> _planRuns = new();
    private readonly HashSet<Guid> _parametersWritten = new();
    
    [Display("Filename", "The name of the spreadsheet where the results are written.", Order: 1)]
    [FilePath(FilePathAttribute.BehaviorChoice.Open, "xls?")]
    public MacroString Path { get; set; } = new MacroString()
    {
        Text = "Results/<Date>-<Verdict>.xlsx"
    };
    
    [Display("Template path", "The path to a template of how to write the results.", Order: 1)]
    [FilePath(FilePathAttribute.BehaviorChoice.Open, "xls?")]
    public MacroString TemplatePath { get; set; } = new MacroString()
    {
        Text = ""
    };

    [Display("Open file", "Opens the file in your default spreadsheet program after plan run.", Order: 2)]
    public bool OpenFile { get; set; } = true;

    [Display("Include", "Include parts of the data in the resulting file.", Order: 3)]
    public Include Include { get; set; } = Include.All;

    [Display("Sheet name", "Decides how data will be divided into separate sheets.", Order: 4)]
    public MacroString SheetName { get; set; } = new MacroString()
    {
        Text = "<ResultName>",
    };
    
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
        base.OnTestPlanRunStart(planRun);
        string templatePath = TemplatePath.Expand(planRun);
        string filePath = Path.Expand(planRun);
        if (File.Exists(filePath))
        {
            File.Delete(filePath);
        }

        bool isTemplate = !string.IsNullOrWhiteSpace(templatePath);
        if (isTemplate)
        {
            File.Copy(templatePath, filePath);
        }
        _spreadSheet = new Spreadsheet(filePath, GetSheetName(planRun), Include.HasFlag(Include.TestPlanSheet), isTemplate);
        
        GetSheet(planRun).AddRows(
            Include.HasFlag(Include.PlanParameters) ? CreateParameters("Plan", planRun) : CreateIdParameters(planRun),
            EmptyResults);
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
            Log.Warning("Created spreadsheet is empty, deleting file. Maybe check if you excluded too much.");
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

        TestPlanRun planRun = _planRuns[stepRun.Id];
        
        if (!_parametersWritten.Contains(stepRun.Id))
        {
            GetSheet(planRun, stepRun).AddRows(
                Include.HasFlag(Include.StepParameters) ? CreateParameters("Step", stepRun) : CreateIdParameters(stepRun),
                EmptyResults);
        }
        
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
        _parametersWritten.Add(stepRunId);
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
        if (Include.HasFlag(Include.ColumnTypePrefix))
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
        if (!Include.HasFlag(Include.ColumnTypePrefix))
        {
            prefix = "";
        }

        Dictionary<string, object> parameters = CreateIdParameters(run, table);

        foreach (ResultParameter parameter in run.Parameters)
        {
            string name = (!string.IsNullOrWhiteSpace(prefix) ? prefix + "/" : "") +
                          (!string.IsNullOrWhiteSpace(parameter.Group) ? parameter.Group + "/" : "") +
                          parameter.Name;
            
            parameters.Add(name, parameter.Value);
        }
        return parameters;
    }

    public override void Close()
    {
        base.Close();
    }
}