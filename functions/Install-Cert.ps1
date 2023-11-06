
<#
Script: Install-SSLCert.ps1
Author: CJ Ramseyer
Date: 9/24/2018
Modified: 11/20/2018
Version: v1.1.0

BREAKING CHANGES:
OTHER CHANGES:
- Simplified name for easier usage Install-SSLCert to Install-Cert

DISCLAIMER: 2023 CJ Ramseyer, All rights reserved.
This sample script is not supported under any support program or service. The sample script is provided AS-IS without warranty of any kind.
CJ Ramseyer disclaims all implied warranties including, without limitation, any implied warranties of merchantability or of fitness for a particular purpose.
The entire risk arising out of the use or performance of the sample script remains with you.
In no event shall CJ Ramseyer, its authors, or anyone else involved in the creation, production, or delivery of the scripts be liable for any damages whatsoever
(including, without limitation, damages for loss of business profits, business interruption, loss of business information, or other pecuniary loss) arising out
of the use of or inability to use the sample script, even if CJ Ramseyer has been advised of the possibility of such damages. 
#>

#Requires -Version 4.0

Function Install-Cert
{
    <#
    .Synopsis
        Function to install a generated cert to a target system under machine or a service
    .Description
        Function to install a generated cert to a target system under machine or a service
    .Parameter CertName
        Required parameter to specify the name of the certificate to be installed
    .Parameter ComputerName
        Optional parameter to install the certificate on a remote machine
    .Parameter InstallMode
        Optional parameter to specify the install mode for the certificate
        Options are Machine or Service.
        Currently ONLY Machine is supported in this script
    .Parameter ServiceName
        Optional Parameter to specify the name of the service where the certificate needs to be installed
    .NOTES

    #>

    [CmdletBinding()]
    param
    (
        [parameter(Mandatory=$true)][string]$CertName,
        [Parameter(Mandatory=$false,ValueFromPipeline=$true)][string]$ComputerName,
        [parameter(Mandatory=$false)][ValidateSet("Machine","Service")][string]$InstallMode = "Machine",
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true)][string]$ServiceName
    )

    switch ($InstallMode)
    {
        "Machine"
        {
            if(!([string]::IsNullOrEmpty($ComputerName)))
            {
                $Credential = Get-Credential
                New-PSDrive -Name "K" -PSProvider FileSystem -Root "\\$ComputerName\c`$\temp" -Credential $Credential
                Copy-Item C:\temp\$CertName.p7b -destination "K:\$CertName.p7b" -force
                while (!(test-Path "K:\$CertName.p7b"))
                {
                    Wait-Event -timeout 2
                    Write-Output "Waiting for $ComputerName"
                }

                if($Credential -ne [System.Management.Automation.PSCredential]::Empty)
                {
                    invoke-command -computername $ComputerName -Credential $Credential -scriptblock {param($s) certreq -q -accept c:\temp\$s.p7b 2>1} -ArgumentList $CertName
                }
                else
                {
                    invoke-command -computername $ComputerName -scriptblock {param($s) certreq -q -accept c:\temp\$s.p7b 2>1} -ArgumentList $CertName
                }
            }
            else
            {
                invoke-command -ScriptBlock {param($s) certreq -q -accept c:\temp\$s.p7b 2>1} -ArgumentList $CertName
            }
        }
        "Service"
        {
            Write-Output "Use Install-ServiceCert.psm1 locally on target system to install service based certificate;exiting"
        }
    }
}
