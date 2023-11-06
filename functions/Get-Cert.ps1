
<#
Script: Get-Cert.ps1
Author: CJ Ramseyer
Date: 9/24/2018
Modified: 8/16/2019
Version: v1.1.2

BREAKING CHANGES:
OTHER CHANGES:
- Added parameter to enable Certificate Transparency options

DISCLAIMER: 2023 CJ Ramseyer, All rights reserved.
This sample script is not supported under any support program or service. The sample script is provided AS-IS without warranty of any kind.
CJ Ramseyer disclaims all implied warranties including, without limitation, any implied warranties of merchantability or of fitness for a particular purpose.
The entire risk arising out of the use or performance of the sample script remains with you.
In no event shall CJ Ramseyer, its authors, or anyone else involved in the creation, production, or delivery of the scripts be liable for any damages whatsoever
(including, without limitation, damages for loss of business profits, business interruption, loss of business information, or other pecuniary loss) arising out
of the use of or inability to use the sample script, even if CJ Ramseyer has been advised of the possibility of such damages. 
#>

#Requires -Version 4.0

Function Get-Cert
{
    <#
    .Synopsis
        Function to generate SSL Certificate based on a provided cert.inf file
    .Description
        Function to generate SSL Certificate based on a provided cert.inf file
    .Parameter CertName
        Required parameter to specify the name of the cert to be used
    .Parameter ComputerName
        Optional parameter to specify a remote target machine
    .Parameter Type
        Optional parameter to change the cert type from the P7B default
    .Parameter enhanced
        Optional parameter to add enhanced features to cert such as smartcardlogin
    .Parameter CertMan
        Optional parameter to specify a differrent CertMan to use instead of the default (Prod)
    .Parameter BusinessOwner
        Optional parameter to change the default business owner
    .Parameter email
        Optional parameter to change the default businessowner email
    .Parameter RenewContact
        Optional parameter to specify the CDSID of the renewal contact
    .Parameter acr
        Optional parameter to specify the ACR number associated with the request
    .Parameter itms
        Optional parameter to specify the itms number associated with the request
    .Parameter CTEnable
        Optional parameter used to enable/disable Certificate Transparency (Option is YES or NO only)
    .Example
        Get-Cert -CertName ADLDSEngSand.company.com -Certman www.certman.company.com
        This would get a signed cert from production CertMan named adldsengsand.company.com and assign CJ Ramseyer as the cert owner without Certificate Transparency
    .NOTES
        CDSID must be a member of ADCS-Groups-SSL-API-Functions to access REST API
        Remote PowerShell MUST BE Configured AND Enabled
        If RemotePS does not work, copy INF file to target machine and generate CSR manually via: certreq -q -f -new "c:\temp\<SourceINF>.inf" "c:\temp\<TargetCSR>.csr"
        When done, copy .csr file back to local machine and rerun command using -CertONLY Parameter
    #>

    [CmdletBinding()]
    param
    (
        [parameter(Mandatory=$true)][string]$CertName,
        [Parameter(Mandatory=$false,ValueFromPipeline=$true)][string[]]$ComputerName,
        [parameter(Mandatory=$false)][ValidateSet("P7B","DER","PEM")][string]$Type = "P7B",
        [parameter(Mandatory=$false)][switch]$enhanced,
        [parameter(Mandatory=$false)][ValidateSet("www.certman.company.com","wwwqa.certman.company.com","wwwdev.certman.company.com")][string]$CertMan = "www.certman.company.com",
        [parameter(Mandatory=$false)][string]$BusinessOwner = "cjramseyer",
        [parameter(Mandatory=$false)][string]$email = "cjr351@gmail.com",
        [parameter(Mandatory=$false)][string]$RenewContact = "cjramseyer",
        [parameter(Mandatory=$false)][string]$acr = "2023-0001",
        [parameter(Mandatory=$false)][string]$itms = "00001",
        [parameter(Mandatory=$false)][switch]$PSExec,
        [parameter(Mandatory=$false)][ValidateSet("YES","NO")][string]$CTEnable = "NO"
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

    $ServerFacing = $ServerFacingExternal
    if($CTEnable -eq "YES")
    {
        [int]$CTEnable = "$CTEnablYes"
    }
    else
    {
        [int]$CTEnable = "$CTEnablNo"
    }

    if(!$enhanced)
    {
        $enhancedusage = "$ServerAuthentication,$ClientAuthentication"
    }
    else
    {
        $enhancedusage = "$ServerAuthentication,$ClientAuthentication,$KDCAuthentication,$SmartCardLogon"
    }

    if(Test-Path "$env:homedrive\$env:homepath\certdefaults.ini")
    {
        Invoke-Expression $(Get-Content "$env:homedrive\$env:homepath\certdefaults.ini" | Out-String)
    }
    else
    {
        if((!$BusinessOwner) -or (!$acr) -or (!$itms) -or (!$RenewContact))
        {
            Write-Output "No defaults file found;if defaults needed create certdefaults.ini here: $env:homedrive\$env:homepath" | Out-File $ErrorLog -Append
            Write-Output "CAUTION: Default file overrides any value specified in both the defaults and the command line" | Out-File $ErrorLog -Append
            $Ask4File = Read-Host "Do you want to set default values?"
            if($Ask4File -eq "YES")
            {
                Out-File -FilePath "$env:homedrive\$env:homepath\certdefaults.ini"
                '$BusinessOwner = ' | Out-File -FilePath "$env:homedrive\$env:homepath\certdefaults.ini" -Append
                '$email = ' | Out-File -FilePath "$env:homedrive\$env:homepath\certdefaults.ini" -Append
                '$RenewContact = ' | Out-File -FilePath "$env:homedrive\$env:homepath\certdefaults.ini" -Append
                '$acr = ' | Out-File -FilePath "$env:homedrive\$env:homepath\certdefaults.ini" -Append
                '$itms = ' | Out-File -FilePath "$env:homedrive\$env:homepath\certdefaults.ini" -Append
                Write-Output "Please update certdefaults.ini and rerun script" | Out-File $ErrorLog -Append
                Break
            }
        }
    }

    Switch ($CertMan)
    {
        "www.certman.company.com"
        {[string]$Client_rp = "urn:certman:prod"}
        "wwwqa.certman.company.com"
        {[string]$Client_rp = "urn:certman:qa"}
        "wwwdev.certman.company.com"
        {[string]$Client_rp = "urn:certman:dev"}
        Default
        {[string]$Client_rp = "urn:certman:prod"}
    }

    # Get credential for REST API
    #Force Get-Credential to use command-line to capture proper UserName for
    $key = 'HKLM:\SOFTWARE\Microsoft\PowerShell\1\ShellIds'
    Set-ItemProperty $key ConsolePrompting True
    $Creds = Get-Credential -Message "Enter CertMan credentials (i.e. xna1\\<cdsid>)"
    $reqcdsid = $Creds.username
    $pwd = $Creds.GetNetworkCredential().password
    $Bytes = [System.Text.Encoding]::ASCII.GetBytes($pwd)
    $Cred64= [Convert]::ToBase64String($bytes)
    $jsonat ='{"client_id": "' + $reqcdsid + '","client_pwdb64": "'+ $Cred64 + '","client_rp": "' + $Client_rp + '"}'
    $url = 'https://' + $Certman + '/REST/api/v1/oauth/jwttoken'
    $responseat = Invoke-RestMethod $url -ContentType 'application/json' -Method post -Body $jsonat
    if($responseat.Status -ne "Success")
    {
        Write-Output "Certman Authentication failed!"
        Remove-Variable -Name * -ErrorAction SilentlyContinue
        Break
    }
    else
    {
        Write-Output "CertMan Authentication successful!"
    }
    $access_token = $responseat.access_token
    #Remove registry entry force Get-Credential to command line
    Remove-ItemProperty $key ConsolePrompting

    $key = "$CertName"
    $sans = '"'+$key.Substring($key.IndexOf(".") +1)+'"'

    if ([string]$csr = Get-Content -Path "C:\temp\$CertName.csr")
    {
        # Query certman to determine if this is a new cert or renewall

        $headers = @{"Auth"="Bearer $access_token"}
        #$query_string = "AnyOwner=THARTSAW&MaxRecordCount=10"
        $query_string = "CN=$CertName&MaxRecordCount=1"
        $url = "https://" + $Certman + "/REST/api/v1/SSLCertificateRequests?$query_string"
        $responseq = Invoke-RestMethod $url -Method get -Headers $headers 
        #$responseq | ConvertTo-Json 

        if ($responseq.TotalRecordCount -gt 0)
        {
            #RENEW 
            $reqnum = $responseq.Certificates.Item(0).RequestNumber  
            $json2 = 
            @"
            {  
                "AcrNumber": "$acr",  
                "BusinessOwner": "$BusinessOwner",  
                "CertificateTypeRequested" : "$Type", 
                "CertificateSigningRequest": "$csr",
                "Email": "$email", 
                "EnhancedKeyUsage": "$enhancedusage", 
                "ExpireNotificationFlag": true,  
                "ITMS": "$itms",  
                "PAHAttestation": true, 
                "RenewContact" : "$RenewContact", 
                "SAN": [$sans],
                "ServerFacing": $ServerFacing, 
                "CTEnable": $CTEnable, 
                "CTEnableAttestation": true
            }
"@

            $headers = @{"Auth"="Bearer $access_token"}
            $url = "https://" + $Certman + "/REST/api/v1/SSLCertificateRequests/$reqnum"
            $response = Invoke-RestMethod $url -ContentType 'application/json' -Method post -Headers $headers -Body $json2
            #$response | ConvertTo-Json 
            $response.status
            $response.message
            $response.ssl_certificate.certStream |Out-File -FilePath C:\temp\$key.p7b -Encoding ascii
        }
        else
        {
            #NEW
            $json2 = 
            @"
            {
                "AcrNumber": "$acr",  
                "Attestor": "$reqcdsid",  
                "BusinessOwner": "$BusinessOwner",  
                "CertificateTypeRequested" : "$Type", 
                "CN": "$key", 
                "CertificateSigningRequest": "$csr",
                "Email": "$email", 
                "EnhancedKeyUsage": "$enhancedusage",  
                "ExpireNotificationFlag": true,  
                "ITMS": "$itms",  
                "PAHAttestation": true, 
                "RenewContact" : "$RenewContact", 
                "SAN": [$sans],
                "ServerFacing": $ServerFacing, 
                "CTEnable": $CTEnable, 
                "CTEnableAttestation": true
            }
"@

            $headers = @{"Auth"="Bearer $access_token"}
            $url = "https://" + $certman + "/REST/api/v1/SSLCertificateRequests"
            $response = Invoke-RestMethod $url -ContentType 'application/json' -Method post -Headers $headers -Body $json2
            #$response | ConvertTo-Json 
            $response.status
            $response.message
            if($response.status -eq "Success")
            {
                $response.ssl_certificate.certStream |Out-File -FilePath C:\temp\$CertName.p7b -Encoding ascii
            }
        }
    }
    else 
    {
        Write-Output "$CertName.csr not found"
    }
}
