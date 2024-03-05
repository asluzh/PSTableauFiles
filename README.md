## PSTableauFiles
This is a PowerShell module that facilitates working with local Tableau content files (workbooks, datasources).
It will allow to perform document structure analysis and updates, such as:
- get the embedded data sources, columns and definitions
- get the structure of report (dashboards, sheets)
- replace embedded data files

### Install PSTableauFiles from the PowerShell Gallery

    Find-Module PSTableauFiles | Install-Module -Scope CurrentUser

### Import Module

    Import-Module PSTableauFiles

## Usage Examples

tbd

## Help Files
The help files for each cmdlet are located in the *help* folder.

# Testing
This repository also contains a suite of Pester tests:
- Module validation tests (tests/PSTableauFiles.Module.Tests.ps1)
- Unit tests for module functions (tests/PSTableauFiles.Unit.Tests.ps1)
- Integration tests, with functionality testing on real Tableau environments (tests/PSTableauFiles.Integration.Tests.ps1)

The tests can be executed using Pester, e.g.

    Invoke-Pester -Tag Module
    Invoke-Pester -Tag Unit -Output Diagnostic
