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
    [Display("Step parameters", "Include the parameters of steps.")]
    StepParameters = 1 << 0,
    [Display("Plan parameters", "Include the parameters of the test plan.")]
    PlanParameters = 1 << 1,
    [Display("Results", "Include all results.")]
    Results = 1 << 2,
    [Display("Run id", "Include the run id of all steps and the plan.")]
    RunId = 1 << 3,
    [Display("Column type prefix", "Include the prefix of columns. Example: 'Step/Verdict' => 'Verdict'.")]
    ColumnTypePrefix = 1 << 4,
    [Display("All", "Include everything.")]
    All = 0b11111,
}

[Display("Spreadsheet", "Save results in an excel file.", "Database")]
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
        _spreadSheet = new Spreadsheet(Path.Expand(), planRun.TestPlanName);
        _spreadSheet.PlanSheet.AddRows(
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
            _spreadSheet.PlanSheet.AddRows(
                Include.HasFlag(Include.StepParameters) ? CreateParameters("Step", stepRun) : CreateIdParameters(stepRun),
                EmptyResults);
        }
        
        base.OnTestStepRunCompleted(stepRun);
    }

    public override void OnResultPublished(Guid stepRunId, ResultTable result)
    {
        if (_spreadSheet is null)
        {
            throw new NullReferenceException();
        }
        
        TestRun stepRun = _testRuns[stepRunId];
        SheetTab sheet = _spreadSheet.GetSheet(result.Name);
        sheet.AddRows(
            Include.HasFlag(Include.StepParameters) ? CreateParameters("Step", stepRun, result) : CreateIdParameters(stepRun, result),
            Include.HasFlag(Include.Results) ? CreateResults(result) : EmptyResults);
        _parametersWritten.Add(stepRunId);
        base.OnResultPublished(stepRunId, result);
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
        
        parameters.Add("StepRunId", run.Id);
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