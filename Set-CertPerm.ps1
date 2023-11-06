
<#
Script: Set-CertPerm.psm1
Author: CJ Ramseyer
Date: 11/20/2018
Modified: 
Version: v1.0.-

BREAKING CHANGES:
OTHER CHANGES:
- Created Cerrt Permission script

DISCLAIMER: 2023 CJ Ramseyer, All rights reserved.
This sample script is not supported under any support program or service. The sample script is provided AS-IS without warranty of any kind.
CJ Ramseyer disclaims all implied warranties including, without limitation, any implied warranties of merchantability or of fitness for a particular purpose.
The entire risk arising out of the use or performance of the sample script remains with you.
In no event shall CJ Ramseyer, its authors, or anyone else involved in the creation, production, or delivery of the scripts be liable for any damages whatsoever
(including, without limitation, damages for loss of business profits, business interruption, loss of business information, or other pecuniary loss) arising out
of the use of or inability to use the sample script, even if CJ Ramseyer has been advised of the possibility of such damages. 
#>

#Requires -Version 4.0

<#
.Synopsis
    Script to set ACL on Certificate directory for NetWork Service account
.Description
    Script to set ACL on Certificate directory for NetWork Service account
.Example
    .\Set-CertPerm.ps1
.INPUTS
    Script has no inputs
.OUTPUTS
    Script has no outputs
.NOTES
    Script does not require any option.  Used only to give network service account permission to read Certificate directory
#>

 $fullpath = "C:\Documents and Settings\All Users\Application Data\Microsoft\Crypto\RSA\MachineKeys"
 $acl=Get-Acl -Path $fullPath
 $permission="NT AUTHORITY\NETWORK SERVICE","Read","Allow"
 $accessRule=new-object System.Security.AccessControl.FileSystemAccessRule $permission
 $acl.AddAccessRule($accessRule)
 Set-Acl $fullPath $acl
 