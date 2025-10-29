# mstodoexporter
A CLI Tool to export all data from a local Microsoft To Do  database.

## Requirements

Microsoft To Do must be installed, user authenticated, and account synchronized.

Update the configuration values in appsettings.json:

- dbPath - full file path to the sqlite database (e.g. %Localappdata%\\Packages\\Microsoft.Todos_8wekyb3d8bbwe\\LocalState\\AccountsRoot\\0892acdb3df940acb21ad7e3cbe37926\\todosqlite.db)
- outputDir - full path to the export location (e.g. %userprofile%\\mstodoexporter_output)
- clearOutputDirBeforeExport (true/false) - clear existing outputDir, if it exists, before exporting data
- archiveOutput (true/false) - create a zip archive of the exported data
- removeOutputDirAfterArchive (true/false) - remove outputDir contents after exporting data
- archiveOutputDirIfExistsBeforeExport (true/false) - create a zip archive of existing outputDir, if it exists, before exporting data
- nonInteractive (true/false) - execute without requiring any user input, such as asking to continue upon encountering an error

## Release History

### Version 1 (2025-10-29)

This is the first version of the tool.

The goal of this release was to just make it work.

The code needs to be cleaned up and restructured.
