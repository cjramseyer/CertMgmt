
<#
Script: CertMgmt.psm1
Author: CJ Ramseyer
Date: 9/25/2018
Modified: 11/19/2018
Version: v1.0.2

BREAKING CHANGES:
OTHER CHANGES:
- Created Certificate Management module

DISCLAIMER: 2023 CJ Ramseyer, All rights reserved.
This sample script is not supported under any support program or service. The sample script is provided AS-IS without warranty of any kind.
CJ Ramseyer disclaims all implied warranties including, without limitation, any implied warranties of merchantability or of fitness for a particular purpose.
The entire risk arising out of the use or performance of the sample script remains with you.
In no event shall CJ Ramseyer, its authors, or anyone else involved in the creation, production, or delivery of the scripts be liable for any damages whatsoever
(including, without limitation, damages for loss of business profits, business interruption, loss of business information, or other pecuniary loss) arising out
of the use of or inability to use the sample script, even if CJ Ramseyer has been advised of the possibility of such damages. 
#>

<#
.Synopsis
    Module of Certificate creation and installation functions
.Description
    Module of Certificate creation and installation functions
#>

Set-PSDebug -Strict

$ModuleBase = $PSScriptRoot

Get-ChildItem $ModuleBase\functions *.ps1 | ForEach-Object{
    . $_.FullName
}
