# Getting started

## Using the Spreadsheet Result Listener
This section goes through how to add a spreadsheet result listener to your test plan to use it later using a GUI.

- Start by installing some GUI for OpenTAP. This example uses the free and opensource plugin [TUI](https://github.com/StefanHolst/opentap-tui)
- Next open up your test plan using the TUI. `tap tui MyPlan.TapPlan`
  
  ![The opened TUI](Pictures/Getting-started(1).png)

- Now go to result settings.

  ![Path to result settings in TUI](Pictures/Getting-started(2).png)

  ![Empty results in TUI](Pictures/Getting-started(3).png)

- Add the spreadsheet result listener.

  ![Add the result listener](Pictures/Getting-started(4).png)

- (Optional) Now you can edit all the settings according to your needs. More information on all the settings are in the [settings](Settings.md) tab.

  ![Modifying the result listener](Pictures/Getting-started(5).png)

- Now the result listener has been added and is ready to be used.

For more information on how to use OpenTAP you can take a look at the [OpenTAP docs](https://doc.opentap.io/User%20Guide/Introduction/Readme.html)

## Generating results
> prerequisite: [Using the spreadsheet Result Listener](#using-the-spreadsheet-result-listener)

This section goes through how to run a test plan and include some example data.

- Install the Demonstration package `tap package install Demonstration`
- Ensure you have added the result listener and added the settings you wanted (see [Using the spreadsheet Result Listener](#using-the-spreadsheet-result-listener)).
- Add a sine result test step.

  ![Add a sine result test step](Pictures/Getting-started(6).png)

 - (Optional) Change the settings of the sine result step.
 - (Optional) Save the test plan.
 - Run the test plan

  If the setting to open file is true the generated spreadsheet should now open.

  If you used default settings you should now have an open spreadsheet with two tabs that look similar to the pictures below.

  ![Plan sheet](Pictures/Getting-started(7).png)

  ![Sine result step sheet](Pictures/Getting-started(8).png)