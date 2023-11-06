# CertMgmt

Script to dynamically register digital certificates

**Request/Incident Description:**

**Background:**

New AD and ADLDS Servers must have a SSL certs installed to support functionality in the environment

**The Problem:**

Current environment requires SSL Certs to be requested by LL6 resource or above

**The Solution:**

1. Migrate to new scripted solution
2. Use CertMgmt module to automate request, creation, and installation of certificates
3. Also provide real-time use of functions to administer environment

**Business Project Benefits:**

Update automated tasks

**Requirements:**

- Assumes use by Directory Services operations team with proper credentials (CertMan Prod and $$ID)

**Execution Steps:**

**Note:** When creating the CertINF file for domain controllers, it is generally not necessary to use the -sans switch
**Note:** When using the -enhanced switch, it must be specified both when using Set-CertINf AND Get-Cert

1. Login to target system using appropriate credentials -OR- Run commands from remote system
2. Import-module .\CertMgmt.psd1
3. Execute CmdLet: Set-CertINF -CertName {CertName}
4. Execute Cmdlet: Get-CSR -CertName {CertName}
5. Execute Cmdlet: Get-Cert.ps1 -CertName {CertName}
6. Execute Install-Cert -CertName {CertName} - **NOTE: Machine based certs ONLY**

**Execution Steps-Add additional systems:**

1. While still logged into first system export the new certificate INCLUDING the private key
2. Copy exported cert to UTIL in same environment or directly to other servers needing the SAME cert

**NOTE:** Install-ServiceCert.psm1 (instructions below) must be used to import certificates to services)

**Execution Steps for SERVICE-BASED certificates (Install-ServiceCert.psm1) (Script must be run ONLY on systems with PS4.0 and above):**

**NOTE:** The script above MUST be executed locally (remote is not currently supported)
**NOTE:** Remember to subsitute the actual environment name for the environment being worked on

1. Export (w/ private key) Cert and Private Key from primary FDS server (pfx format) ExportedCert-FDSDev.pfx (same machine where csr was generated)
2. Login to the target system using appropriate credentials
3. Copy the Exported certificate to C:\INSTAPPS\CertMgmt
4. Execute Import-Module C:\INSTAPPS\CertMgmt\Install-ServiceCert.psm1
5. Execute Add-Certificate -Path "C:\INSTAPPS\CertMgmt\ExportedCert-FDSDev.pfx" -Store "My" -Location CERT_SYSTEM_STORE_SERVICES -ServiceName FDSSand

**NOTE:** When prompted for a password, if no password was entered during the export step, simply press ENTER

**Validation steps:**

1. Open PowerShell Prompt
2. CD "C:\Documents and Settings\All Users\Application Data\Microsoft\Crypto\RSA"
3. Execute (Get-Acl .\MachineKeys\).access
4. Validate the following entry (Network Service) is included:

FileSystemRights    :   ReadAndExecute, Syncronize
AccessControlType   :   Allow
IdentityReference   :   NT AUTHORITY\NETWORK SERVICE
IsInherited         :   False
InheritanceFlags    :   ContainerInherit, ObjectInherit
PropogationFlags    :   none

**NOTE:** If the above entry DOES NOT EXIST, the service will NOT have permission to read the Cert and SSL via port 636 will not work.  If this is tru, execute: C:\INSTAPPS\CertMgmt\Set-CertPerm.ps1

**Backout Plan (only if required and approved by AD Engineering)**: NONE

1. Remove certificates

Reference Ticket (If Any):
Proprietary
Record Series:  17.01
Record Type:  Official
Retention Period: C+3,T
