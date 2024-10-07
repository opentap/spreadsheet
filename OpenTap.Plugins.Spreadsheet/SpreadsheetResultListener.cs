using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using OpenTap;

namespace Spreadsheet;

[Flags]
public enum Include
{
    [Display("All", "Include everything.")]
    All = 0b111111,
    [Display("Step parameters", "Include the parameters of steps.")]
    StepParameters = 1 << 0,
    [Display("Plan parameters", "Include the parameters of the test plan.")]
    PlanParameters = 1 << 1,
    [Display("Results", "Include all results.")]
    Results = 1 << 2,
    [Display("Run id", "Include the run id of all steps and the test plan.")]
    RunId = 1 << 3,
    [Display("Column type prefix", "Include the prefix of columns. Example: 'Step/Verdict' => 'Verdict'.")]
    ColumnTypePrefix = 1 << 4,
    [Display("Test plan sheet", "Include a first sheet with plan parameters and step parameters for steps without results.")]
    TestPlanSheet = 1 << 5,
    [Display("None", "Include nothing in the spreadsheet (due to limitations with .xls files this wont generate a file at all).")]
    None = 0,
}

public enum SplitBy
{
    [Display("Result name", "Split data into sheets by the name of result tables.")]
    ResultName,
    [Display("Step name", "Split data into sheets by the name of steps.")]
    StepName,
    [Display("Step run", "Split data into sheets by the id step runs.")]
    StepRun,
    [Display("Step type", "Split data into sheets by the type of steps.")]
    StepType,
    [Display("No split", "Put all data into the plan sheet.")]
    NoSplit,
}

[Display("Spreadsheet", "Save results in a spreadsheet.", "Database")]
public sealed class SpreadsheetResultListener : ResultListener
{
    private static readonly Dictionary<string, Array> EmptyResults = new();
    
    private readonly Dictionary<Guid, TestRun> _testRuns = new();
    private readonly HashSet<Guid> _parametersWritten = new();
    
    [Display("Filename", "The name of the spreadsheet where the results are written.", Order: 1)]
    [FilePath(FilePathAttribute.BehaviorChoice.Open, "xls?")]
    public MacroString Path { get; set; } = new MacroString()
    {
        Text = "Results/<Date>-<Verdict>.xlsx"
    };

    [Display("Open file", "Opens the file in your default spreadsheet program after plan run.", Order: 2)]
    public bool OpenFile { get; set; } = true;

    [Display("Include", "Include parts of the data in the resulting file.", Order: 3)]
    public Include Include { get; set; } = Include.All;

    [Display("Split by", "Decides how data will be divided into separate sheets.", Order: 4)]
    public SplitBy SplitBy { get; set; } = SplitBy.ResultName;
    
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
        _spreadSheet = new Spreadsheet(Path.Expand(), GetSheetName(planRun), Include.HasFlag(Include.TestPlanSheet));
        GetSheet(planRun).AddRows(
            Include.HasFlag(Include.PlanParameters) ? CreateParameters("Plan", planRun) : CreateIdParameters(planRun),
            EmptyResults);
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
            Process.Start(path);
        }
    }

    public override void OnTestStepRunStart(TestStepRun stepRun)
    {
        if (_spreadSheet is null)
        {
            throw new NullReferenceException();
        }
        
        _testRuns.Add(stepRun.Id, stepRun);
        base.OnTestStepRunStart(stepRun);
    }

    public override void OnTestStepRunCompleted(TestStepRun stepRun)
    {
        if (_spreadSheet is null)
        {
            throw new NullReferenceException();
        }
        
        if (!_parametersWritten.Contains(stepRun.Id))
        {
            GetSheet(stepRun).AddRows(
                Include.HasFlag(Include.StepParameters) ? CreateParameters("Step", stepRun) : CreateIdParameters(stepRun),
                EmptyResults);
        }
        
        base.OnTestStepRunCompleted(stepRun);
    }

    public override void OnResultPublished(Guid stepRunId, ResultTable result)
    {
        TestRun run = _testRuns[stepRunId];
        SheetTab sheet = GetSheet(run, result);
        sheet.AddRows(
            Include.HasFlag(Include.StepParameters) ? CreateParameters("Step", run, result) : CreateIdParameters(run, result),
            Include.HasFlag(Include.Results) ? CreateResults(result) : EmptyResults);
        _parametersWritten.Add(stepRunId);
        base.OnResultPublished(stepRunId, result);
    }
    
    private SheetTab GetSheet(TestRun run, ResultTable? table = null)
    {
        if (_spreadSheet is null)
        {
            throw new NullReferenceException();
        }

        return _spreadSheet.GetSheet(GetSheetName(run, table));
    }
    
    private string GetSheetName(TestRun run, ResultTable? table = null)
    {
        if (run is TestPlanRun planRun)
        {
            return Include.HasFlag(Include.RunId) ? planRun.Id.ToString().Substring(0, 8) : planRun.TestPlanName;
        }
        
        if (_spreadSheet is null)
        {
            throw new NullReferenceException();
        }

        TestStepRun? stepRun = run as TestStepRun;
        
        return SplitBy switch
        {
            SplitBy.ResultName when table is not null => table.Name,
            SplitBy.StepName when stepRun is not null => stepRun.TestStepName,
            // Tabs cannot be more than 31 characters long. So we just get the first 8 for the tab names.
            SplitBy.StepRun => run.Id.ToString().Substring(0, 8),
            SplitBy.StepType when stepRun is not null => stepRun.TestStepTypeName,
            SplitBy.NoSplit => _spreadSheet.PlanSheet.Name,
            _ => _spreadSheet.PlanSheet.Name,
        };
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