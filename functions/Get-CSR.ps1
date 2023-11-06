
<#
Script: Get-CSR.ps1
Author: CJ Ramseyer
Date: 11/19/2018
Modified: 
Version: v1.0.0

BREAKING CHANGES:
OTHER CHANGES:
- Created CSR cmdlet

DISCLAIMER: 2023 CJ Ramseyer, All rights reserved.
This sample script is not supported under any support program or service. The sample script is provided AS-IS without warranty of any kind.
CJ Ramseyer disclaims all implied warranties including, without limitation, any implied warranties of merchantability or of fitness for a particular purpose.
The entire risk arising out of the use or performance of the sample script remains with you.
In no event shall CJ Ramseyer, its authors, or anyone else involved in the creation, production, or delivery of the scripts be liable for any damages whatsoever
(including, without limitation, damages for loss of business profits, business interruption, loss of business information, or other pecuniary loss) arising out
of the use of or inability to use the sample script, even if CJ Ramseyer has been advised of the possibility of such damages. 
#>

#Requires -Version 4.0

Function Get-CSR
{
    <#
    .Synopsis
        Function to generate CSR (Certificate Signing Request) from the provided cert.inf file
    .Description
        Function to generate CSR (Certificate Signing Request) from the provided cert.inf file
    .Parameter CertName
        Required parameter to specify the name of the cert to be used
    .Parameter CertINF
        Optional parameter to specify the INF file to be used
    .Parameter enhanced
        Optional parameter to add enhanced features to cert such as smartcardlogin
    .Parameter ComputerName
        Optional parameter to specify a remote target machine
    .Parameter PSExec
        Optional parameter to use PSExec instead of RemotePS
    .NOTES
        If RemotePS does not work, copy INF file to target machine and generate CSR manually via: certreq -q -f -new "c:\temp\<SourceINF>.inf" "c:\temp\<TargetCSR>.csr"
        When done, copy .csr file back to local machine and rerun command using -CertONLY Parameter
    #>

    [CmdletBinding()]
    param
    (
        [parameter(Mandatory=$true)][string]$CertName,
        [parameter(Mandatory=$false)][string]$CertInf = "$CertName.inf",
        [parameter(Mandatory=$false)][string]$CertPath = "C:\temp",
        [Parameter(Mandatory=$false,ValueFromPipeline=$true)][string[]]$ComputerName,
        [parameter(Mandatory=$false)][switch]$PSExec
    )

    Set-PSDebug -Strict

    if(!([string]::IsNullOrEmpty($ComputerName)))
    {
        if(Test-Path "$Certinf")
        {
            $Credential = Get-Credential -Message 'Enter Remote System Admin Credentials (i.e. devad\$$cramse22)'
            New-PSDrive -Name "K" -PSProvider FileSystem -Root "\\$ComputerName\c`$\temp" -Credential $Credential
            Copy-Item "$CertPath\$CertName.inf" -Destination "K:\$CertName.inf"
        }
        else
        {
            Write-Output "Certificate Information file not found to generate cert"
            Break
        }    
    }

    if(!([string]::IsNullOrEmpty($ComputerName)) -and !($PSExec))
    {
        if($Credential -ne [System.Management.Automation.PSCredential]::Empty)
        {
            invoke-command -computername $ComputerName -Credential $Credential -scriptblock {param($s) certreq -q -f -new "c:\temp\$s.inf" "c:\temp\$s.csr" 2>1} -ArgumentList $CertName
        }
        else
        {
            invoke-command -computername $ComputerName -scriptblock {param($s) certreq -q -f -new "c:\temp\$s.inf" "c:\temp\$s.csr" 2>1} -ArgumentList $CertName
        }

        while (!(test-Path "K:\$CertName.csr"))
        {
            wait-event -timeout 2
            Write-Output "Waiting for $ComputerName"
        }
        if(Test-Path "K:\$CertName.csr")
        {
            Copy-Item "K:\$CertName.csr" -Destination "C:\temp\$CertName.csr"
        }
        else
        {
            Write-Output "Could not find K:\$CertName.csr"
            Write-Output "If csr only needs signing, use -CertONLY switch;ensure that csr is in c:\temp first"
            Break
        }
        Remove-PSDrive -Name "K"
    }
    elseif(!([string]::IsNullOrEmpty($ComputerName)) -and ($PSExec))
    {
        #Use PSExec for the remote command instead of remote powershell
        $PSCMD = "certreq -q -f -new `"c:\temp\$CertName.inf`" `"c:\temp\$CertName.csr`""
        psexec \\$ComputerName -u "$($Credential.UserName)" -p "$($Credential.Password)" CMD /c "$PSCMD"
    }
    else
    {
        Invoke-Command -ScriptBlock {param($s) certreq -q -f -new "c:\temp\$s.inf" "c:\temp\$s.csr"} -ArgumentList $CertName
    }

    while (!(test-Path C:\temp\$CertName.csr))
    {
        wait-event -timeout 2
        Write-Output "Waiting for $ComputerName"
    }
}
