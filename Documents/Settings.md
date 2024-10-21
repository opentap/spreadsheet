# Settings

## Filename
> Default value: `Results/<Date>-<Verdict>.xlsx`

This is a path to where the spreadsheet will be created.

The filename is a macro string, which means you can use macros like `<Verdict>` `<TestPlanName>` etc.
More info [here](https://doc.opentap.io/Developer%20Guide/Appendix/Readme.html#result-listeners).

## Open file
> Default value: `true`

If true the resulting spreadsheet will be opened after the plan is done running. The program used to open the file will be your system default for the filetype.

## Include
> Default value: `All`

Specify what data to include in the spreadsheet.
Here is a list of what types of data can be included along with examples:

| Name               | Description                                                                               | Example                                                                                                   |
|--------------------|-------------------------------------------------------------------------------------------|-----------------------------------------------------------------------------------------------------------|
| All                | Includes everything in the generated spreadsheet.                                         |                                                                                                           |
| Step parameters    | Include the parameters of steps.                                                          | Include columns by names like 'Step/Verdict'                                                              |
| Plan parameters    | Include the parameters of the test plan.                                                  | Include columns by names like 'Plan/StartTime'                                                            |
| Results            | Include all results                                                                       | Include columns by names like 'Results/Value' (for a 'Sine Result' step from the 'Demonstration' package) |
| Run id             | Include the run id of all steps and the test plan.                                        | Include the 3 default columns: 'RunId', 'ParentRunId' and 'ResultName'                                    |
| Column type prefix | Include the prefix of columns.                                                            | Columns wont have their prefixes ('Step', 'Plan' and 'Result'). So 'Step/Verdict' becomes 'Verdict' etc.  |
| Test plan sheet    | Include a first sheet with plan parameters and step parameters for steps without results. | Include a sheet with parameters for the testplan                                                          |
| None               | Don't include anything in the generated spreadsheet.                                      | Wont generate a TestPlan since no sheets are created.                                                     |

## Sheet Name
> Default value: `<ResultName>`

Will split the spreadsheet into multiple different sheets if two steps and/or results generate different sheet names.

The sheet name is a macro string, which means you can use macros like `<Verdict>` and `<TestPlanName>` etc.
More info [here](https://doc.opentap.io/Developer%20Guide/Appendix/Readme.html#result-listeners).

This option is similar to Filename, however there it does have more macros available than the ones documented on [doc.opentap.io](https://doc.opentap.io/Developer%20Guide/Appendix/Readme.html#result-listeners)
These macros will have different names for the plan than it will for the individual result and/or step.

| Macro            | Step/Result value                                | Step/Result example.              | Plan value            | Plan example       |
|------------------|--------------------------------------------------|-----------------------------------|-----------------------|--------------------|
| `<RunId>`        | The GUID of the step run.                        | `bd4e6...`                        | The GUID of the plan. | `d62b7...`         |
| `<StepId>`       | The GUID of the step.                            | `7a3e4...`                        | The GUID of the plan. | `a862f...`         |
| `<ResultName>`   | The name of the result table.                    | `RampResults`                     | The name of the plan. | `MyTestPlan`       |
| `<StepName>`     | The name of the step.                            | `MyTestStep`                      | The name of the plan. | `MyTestPlan`       |
| `<StepType>`     | The name of the step type.                       | `RampResultsStep`                 | `TestPlan`            | `TestPlan`         |
| `<StepTypeFull>` | The name (including namespace) of the step type. | `OpenTap.Plugins.Demo.ResultsAnd` | `OpenTap.TestPlan`    | `OpenTap.TestPlan` |

All sheets names will be updated to comply with sheet name restrictions according to the [microsoft documentation](https://support.microsoft.com/en-gb/office/rename-a-worksheet-3f1f7148-ee83-404d-8ef0-9ff99fbad1f9#:~:text=Important%3A%20Worksheet%20names%20cannot%3A,Contain%20more%20than%2031%20characters.)

1. All illegal characters get replaced with `.`
2. All beginning and trailing `'` characters are removed.
3. Name is truncated to 31 characters.
4. If the resulting name is equal to `History` it will be renamed to `.History`