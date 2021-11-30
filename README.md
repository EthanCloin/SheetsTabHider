
# SheetsTabHider
SheetsTabHider is a Google Apps Script project that allows users to control visibility of tabs in a Spreadsheet

## Why Should I Use This?
Google Sheets can support a significant number of Sheet tabs within a Spreadsheet window. Sometimes it can be a struggle to quickly locate and enter desired Sheets. By hiding all tabs except the currently useful ones, you can more easily work with particularly dense Spreadsheets.

## How Do I Use This?
SheetsTabHider is written as a "container-bound" Google Apps Script project. This means all you need to do is copy the .gs files from this repository into the Apps Script Editor in your Spreadsheet. See below for step-by-step instructions!

### Open the Apps Script Editor
In any Spreadsheet, navigate to the Apps Script Editor inside the Extensions menu. 

<img src="https://drive.google.com/uc?export=view&id=10P0UfuGniM3f_x6tnH4gu4arn6qXenGZ" width="600" height="400"/>


NOTE: Google Sheets before Nov. 2021 may have a different menu arrangement from the screenshot below. In such cases, you can find the Apps Script Editor by selecting Tools --> Script Editor. 

### Paste Content of Both Files
For your Spreadsheet to utilize the functions available in this repository, simply paste the content of the two .gs files into your project like the screenshot below. To keep things neat, be sure to name your project appropriately in the top left of the editor interface.

<img src="https://drive.google.com/uc?export=view&id=1nSAizIuDEHMDDlUF8Tam0rgr4rjq3B1u" width="600" height="400"/>

### Install the Trigger
Navigate to the *InterfaceScript* file in your editor. Ensure the function "installTrigger" appears in the dropdown on the top of the editor and press "Run". Installing this trigger will enable the script to execute an onOpen() function each time you open or refresh your Spreadsheet. 

### Grant Permissions
Upon running "installTrigger" Google will prompt you to grant permissions for this project. Grant all requested permissions to ensure the project runs as advertised. You may see this warning during the permissions dialogue:

<img src="https://drive.google.com/uc?export=view&id=1H2L31XtvHhK2GO9EsjWVlzUQjBLYZdtH" width="600" height="400"/>

Click "Advanced" and continue to grant permissions for the script. After this, "installTrigger" should have run. Check the Project Triggers by clicking the clock icon on the left side of the editor. If there are no existing triggers, run "installTrigger" again. If your Project Triggers matches the below screenshot, you're  done with the installation!

<img src="https://drive.google.com/uc?export=view&id=1bdfwdx5eltCQ_4RWaVMxcUDv5CKhkT9O" width="600" height="400"/>

Save and close your project and check your Spreadsheet. You should see a new Sheet and a new Menu called "TabHider". If you don't, refresh the spreadsheet to trigger the onOpen function.

### Hide/Reveal Your Sheets!
This script allows you to perform bulk and individual operations. The bulk operation will be executed first, so you could select Hide All and then Reveal the individual sheets that you are interested in. Then use the TabHider dropdown to execute the commands you selected!

<img src="https://drive.google.com/uc?export=view&id=1y3UAAWEuHg1MiJgaMDF3CCZB4Zdm47ba" width="600" height="400"/>
