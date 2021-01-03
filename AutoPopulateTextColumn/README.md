# Auto Populate Text Column
AutoPopulateTextColumn PCF (PowerApps Component framework) Control that auto populate a PowerApps text column with Regarding Lookup details. The text format is configurable for each table type and can be provided as input property while setting up the control.

Auto populate the activity subject with Regarding details. The format is configurable and can be setup for each table type.

### Features
---
* Auto populate activity subject on create of activities from timeline.
* User can add additional text in the subject before save.
* On change of Regarding value the control re-populates with new formatted details.

### Configure the control
---
Control Property | Description | Required
------------ | ------------- | -------------
Bound Column for Auto Populate | Text column to attach the control | Yes |
Regarding Value | Bind to a value on a Lookup.Regarging column | Yes |
Configurations Value | Provide the configurations related to regarding table type, select columns and text format(tableType1;column1,column2;column1 - column2^tableType2;column1,column2;column1 - column2) | Yes |


### Screenshots
---
```bash
incident;ticketnumber,title;title - Case Number:ticketnumber^account;name;Account Name:name
```

# Get required tools
---

Be sure to update your Microsoft PowerApps CLI to the latest version: 
```bash
pac install latest
```
# Build the control
---

* Clone the repo/ download the zip file.
* Navigate to ./AutoPopulateTextColumn/ folder.
* Copy the folder path and open it in visual studio code.
* Open the terminal, and run the following command to install the project dependencies:
```bash
npm install
```
* After npm installation, navigate to  node_modules -> pcf-scripts -> Open the 'featureflags.json' file. Update 'pcfAllowLookup' attributes to 'on' ("pcfAllowLookup": "on").

Then run the command:
```bash
npm run start
```
### Build the solution
---

```bash
msbuild /t:restore
``` 

```bash
msbuild
``` 
* You will have the solution file in SolutionFolder/bin/debug folder!

If you want to change the solution type you have to edit the .cdsproj file:
```bash
Solution Packager overrides, un-comment to use: SolutionPackagerType (Managed, Unmanaged, Both)
  <PropertyGroup>
    <SolutionPackageType>Managed</SolutionPackageType>
  </PropertyGroup>

  ```
