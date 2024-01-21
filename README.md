# Schedule Generator - Google Sheets App Project

This project is a Google Sheets App designed to generate a timetable according to class data.

## How to Use

Follow these steps to integrate the Schedule Generator into your Google Sheets project:

1. Open your Google Sheets project.
2. Open the Script Editor:
   - Click on "Extensions" in the top menu.
   - Select "Apps Script."
3. In the Script Editor, paste the content of `app_script.gs` into the `code.gs` file.
4. Modify the manifest file (`appsscript.json`) to grant necessary permissions for the script to execute.

## Manifest Modification

Ensure that the `appsscript.json` file includes the required permissions. You may need to add the following sections:

```json
{
  "oauthScopes": ["..."],
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8"
}
