# Settings

## Filename
> Default value: `Results/<Date>-<Verdict>.xlsx`

This is a to where the spreadsheet will be created.

The filename is a macro string, which means you can use macros like `<Verdict>` `<TestPlanName>` etc.
More info [here](https://doc.opentap.io/Developer%20Guide/Appendix/Readme.html#result-listeners).

## Open file
> Default value: `true`

If true the resulting spreadsheet will be opened after the plan is done running. The program used to open the file will be your system default for the filetype.

## Include
> Default value: `All`

Specify what data to include in the spreadsheet.
Here is a list of what types of data can be included along with examples:

|Name|Description|Example|
|-|-|-|
|All|Includes everything in the generated spreadsheet.||
|Step parameters|Include the parameters of steps.|Include columns by names like 'Step/Verdict'|
|Plan parameters|Include the parameters of the test plan.|Include columns by names like 'Plan/StartTime'|
|Results|Include all results|Include columns by names like 'Results/Value' (for a 'Sine Result' step from the 'Demonstration' package)|
|Run id|Include the run id of all steps and the test plan.|Include the 3 default columns: 'RunId', 'ParentRunId' and 'ResultName'|
|Column type prefix|Include the prefix of columns.|Columns wont have their prefixes ('Step', 'Plan' and 'Result'). So 'Step/Verdict' becomes 'Verdict' etc.|
|Test plan sheet|Include a first sheet with plan parameters and step parameters for steps without results.|Include a sheet with name `<PlanName>` or `<PlanId>`
|None|Include nothing in the spreadsheet. (due to limitations with .xls files this wont generate a file at all).||

## Split by
> Default value: `ResultName`

Specify how to split the data into sheets. (A sheet is a tab in the spreadsheet.)
This option decides how to name sheets and how many sheets there are. However it does not decide whether the plan sheet is generated or not. That is specified by the [include parameter](#Include). This means if you select no split and dont include the test plan you wont get a spreadsheet at all.

|Name|Description|Example|
|-|-|-|
|Result name| Split data into sheets by the name of result tables.|One sheet per result name.|
|Step name| Split data into sheets by the name of steps.|One sheet per step name.|
|Step run| Split data into sheets by the id step runs.|One sheet per step run.|
|Step type| Split data into sheets by the type of steps.|One sheet per step type.|
|No split| Put all data into the plan sheet.|Only one sheet.|
