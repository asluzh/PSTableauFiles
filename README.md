## PSTableauFiles
This is a PowerShell module that facilitates working with local Tableau content files (workbooks, datasources).
It will allow to perform document structure analysis and updates, such as:
- get the embedded data sources, columns and definitions
- get the structure of report (dashboards, sheets)
- replace embedded data files

## Install and Importing Module

tbd

## Usage Examples

tbd

## Help Files
The help files for each cmdlet are located in the *docs* folder.

# Testing
This repository also contains a suite of module tests for PSTableauFiles:
- Module integrity tests (PSTableauFiles.Module.Tests.ps1)
- Basic unit tests for module functions (PSTableauFiles.Unit.Tests.ps1)

The tests can be executed using Pester, e.g.

    Invoke-Pester -Tag Module
    Invoke-Pester -Tag Unit -Output Diagnostic
