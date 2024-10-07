# Content

The OpenTap spreadsheet plugin can write any test plan results to spreadsheets.

## Settings

### Filename
> Default value: `Results/<Date>-<Verdict>.xlsx`

This is a to where the spreadsheet will be created.

The filename is a macro string, which means you can use macros like `<Verdict>` `<TestPlanName>` etc.
More info [here](https://doc.opentap.io/Developer%20Guide/Appendix/Readme.html#result-listeners).

### Open file
> Default value: `true`

If true the resulting spreadsheet will be opened after the plan is done running. The program used to open the file will be your system default for the filetype.

### Include
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