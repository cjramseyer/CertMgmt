
<#
Script: Set-CertINF.ps1
Author: CJ Ramseyer
Date: 9/24/2018
Modified: 
Version: v1.0.0

BREAKING CHANGES:
OTHER CHANGES:
- Created function to generate cert.inf file

DISCLAIMER: 2023 CJ Ramseyer, All rights reserved.
This sample script is not supported under any support program or service. The sample script is provided AS-IS without warranty of any kind.
CJ Ramseyer disclaims all implied warranties including, without limitation, any implied warranties of merchantability or of fitness for a particular purpose.
The entire risk arising out of the use or performance of the sample script remains with you.
In no event shall CJ Ramseyer, its authors, or anyone else involved in the creation, production, or delivery of the scripts be liable for any damages whatsoever
(including, without limitation, damages for loss of business profits, business interruption, loss of business information, or other pecuniary loss) arising out
of the use of or inability to use the sample script, even if CJ Ramseyer has been advised of the possibility of such damages. 
#>

#Requires -Version 4.0

Function Set-CERTINF
{
    <#
    .Synopsis
        Function to generate the Certtificate information file used to generate the certificate request file
    .Description
        Function to generate the Certtificate information file used to generate the certificate request file
    .Parameter CertName
        Required parameter to specify the name of the certificate
    .Parameter KeyLength
        Optional parameter to specify a differnet key length
        NOTE: ONLY Values greater than 2048 whole integers will be accepted
    .Parameter enhanced
        Optional parameter ro enable extended cert attriobutes such as smartcardlogin
    .Parameter sans
        Optional parameter to specify a comma delimited list of subject alternative names
        i.e. "fdsdev.ford.com","fdsdev.ford.com"
    .Example
        Set-CertINF -Name fdsdev.ford.com -sans "fdsdev.ford.com","fdsdevs.ford.com","fmc128025.fmcc.ford.com","fmc128028.fmcc.ford.com"
    #>

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)][string[]]$CertName,
        [parameter(Mandatory=$false)][ValidateRange(2048,[int]::MaxValue)][int16]$KeyLength = 2048,
        [parameter(Mandatory=$false)][switch]$enhanced,
        [parameter(Mandatory=$false)][string[]]$sans
    )

    Set-PSDebug -Strict

    Set-Variable -Name ServerAuthentication -Value 1.3.6.1.5.5.7.3.1 -Option Constant
    Set-Variable -Name ClientAuthentication -Value 1.3.6.1.5.5.7.3.2 -Option Constant
    Set-Variable -Name EmailProtection -Value 1.3.6.1.5.5.7.3.4 -Option Constant
    Set-Variable -Name IPSecEndSystem -Value 1.3.6.1.5.5.7.3.5 -Option Constant
    Set-Variable -Name IPSecTunnel -Value 1.3.6.1.5.5.7.3.6 -Option Constant
    Set-Variable -Name IPSecUser -Value 1.3.6.1.5.5.7.3.7 -Option Constant
    Set-Variable -Name KDCAuthentication -Value 1.3.6.1.5.2.3.5 -Option Constant
    Set-Variable -Name SmartCardLogon -Value 1.3.6.1.4.1.311.20.2.2 -Option Constant
    Set-Variable -Name IPSecIKE -Value 1.3.6.1.5.5.7.3.17 -Option Constant
    Set-Variable -Name IKEIntermediate -Value 1.3.6.1.5.5.8.2.2 -Option Constant
    Set-Variable -Name SANSupport -value 2.5.29.17 -Option Constant
    Set-Variable -Name ServerFacingUnkown -Value 0 -Option Constant
    Set-Variable -Name ServerFacingInternal -Value 1 -Option Constant
    Set-Variable -Name ServerFacingExternal -Value 2 -Option Constant
    Set-Variable -Name ServerFacingMixed -Value 3 -Option Constant
    Set-Variable -Name CTEnableUnkown -Value 0 -Option Constant
    Set-Variable -Name CTEnablYes -Value 1 -Option Constant
    Set-Variable -Name CTEnablNo -Value 2 -Option Constant

    $Signature = '"$Windows NT$"'

    Write-Output "Processing $ComputerName"

    $key="$CertName"

    $INF = @()
    $INF += '[Version]'
    $INF += 'Signature='+"$Signature"
    $INF += '[NewRequest]'
    $INF += ''
    $INF += 'Subject='+"`"CN=$key, O=Big LDAP Company, OU=IT, L=Somecity, ST=Somestate, C=US`""
    $INF += ''
    $INF += 'KeySpec = 1'
    $INF += 'KeyLength = '+"$KeyLength"
    $INF += 'Exportable = TRUE'
    $INF += 'MachineKeySet = TRUE'
    $INF += ''
    $INF += 'SMIME = False'
    $INF += 'PrivateKeyArchive = FALSE'
    $INF += 'UserProtected = FALSE'
    $INF += 'UseExistingKeySet = FALSE'
    $INF += 'ProviderName = "Microsoft RSA SChannel Cryptographic Provider"'
    $INF += 'ProviderType = 12'
    $INF += 'RequestType = PKCS10'
    $INF += 'KeyUsage = 0xa0'

    if($enhanced)
    {
        $INF += ''
        $INF += '[EnhancedKeyUsageExtension]'
        $INF += 'OID='+"$ServerAuthentication"
        $INF += 'OID='+"$ClientAuthentication"
        $INF += 'OID='+"$SmartCardLogon"
        $INF += 'OID='+"$KDCAuthentication"
    }
    elseif((!$enhanced) -and (!$sans))
    {
        $INF += ''
        $INF += '[EnhancedKeyUsageExtension]'
        $INF += 'OID='+"$ServerAuthentication"
        $INF += 'OID='+"$ClientAuthentication"
    }

    if($sans)
    {
        $INF += ''
        $INF += '[Extensions]'
        $INF += ''
        $INF += "$SANSupport"+' = "{text}"'
        foreach($san in $sans)
        {
            $INF += '_continue_ = '+"`"dns=$san&`""
        }
    }

    $sans = '"'+$key.Substring($key.IndexOf(".") +1)+'"'

    Write-Output "Certificate Information file is being generated `n "

    $INF | out-file -Filepath "c:\temp\$CertName.inf" -Encoding ascii
}
