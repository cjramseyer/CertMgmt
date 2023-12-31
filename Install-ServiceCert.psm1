
<#
Script: Install-ServiceCert.psm1
Author: CJ Ramseyer
Date: 9/25/2018
Modified: 11/19/2018
Version: v1.0.2

BREAKING CHANGES:
OTHER CHANGES:
- Fixed Write-Output command
- Added cert permissions logic
- Fixed typo

DISCLAIMER: 2023 CJ Ramseyer, All rights reserved.
This sample script is not supported under any support program or service. The sample script is provided AS-IS without warranty of any kind.
CJ Ramseyer disclaims all implied warranties including, without limitation, any implied warranties of merchantability or of fitness for a particular purpose.
The entire risk arising out of the use or performance of the sample script remains with you.
In no event shall CJ Ramseyer, its authors, or anyone else involved in the creation, production, or delivery of the scripts be liable for any damages whatsoever
(including, without limitation, damages for loss of business profits, business interruption, loss of business information, or other pecuniary loss) arising out
of the use of or inability to use the sample script, even if CJ Ramseyer has been advised of the possibility of such damages. 
#>

#region Internals
#region .net Types
$certStoreTypes = @'
using System;
using System.Runtime.InteropServices;

namespace System.Security.Cryptography.X509Certificates
{
    public class Win32
    {
        [DllImport("crypt32.dll", EntryPoint="CertOpenStore", CharSet=CharSet.Auto, SetLastError=true)]
        public static extern IntPtr CertOpenStore(
            int storeProvider,
            int encodingType,
            IntPtr hcryptProv,
            int flags,
            String pvPara);
                                    
        [DllImport("crypt32.dll", EntryPoint="CertCloseStore", CharSet=CharSet.Auto, SetLastError=true)]
        [return : MarshalAs(UnmanagedType.Bool)]
        public static extern bool CertCloseStore(
            IntPtr storeProvider,
            int flags);
    }

    public enum CertStoreLocation
    {
        CERT_SYSTEM_STORE_CURRENT_USER = 0x00010000,
        CERT_SYSTEM_STORE_LOCAL_MACHINE = 0x00020000,
        CERT_SYSTEM_STORE_SERVICES = 0x00050000,
        CERT_SYSTEM_STORE_USERS = 0x00060000
    }

    [Flags]
    public enum CertStoreFlags
    {
        CERT_STORE_NO_CRYPT_RELEASE_FLAG = 0x00000001,
        CERT_STORE_SET_LOCALIZED_NAME_FLAG = 0x00000002,
        CERT_STORE_DEFER_CLOSE_UNTIL_LAST_FREE_FLAG = 0x00000004,
        CERT_STORE_DELETE_FLAG = 0x00000010,
        CERT_STORE_SHARE_STORE_FLAG = 0x00000040,
        CERT_STORE_SHARE_CONTEXT_FLAG = 0x00000080,
        CERT_STORE_MANIFOLD_FLAG = 0x00000100,
        CERT_STORE_ENUM_ARCHIVED_FLAG = 0x00000200,
        CERT_STORE_UPDATE_KEYID_FLAG = 0x00000400,
        CERT_STORE_BACKUP_RESTORE_FLAG = 0x00000800,
        CERT_STORE_READONLY_FLAG = 0x00008000,
        CERT_STORE_OPEN_EXISTING_FLAG = 0x00004000,
        CERT_STORE_CREATE_NEW_FLAG = 0x00002000,
        CERT_STORE_MAXIMUM_ALLOWED_FLAG = 0x00001000
    }

    public enum CertStoreProvider
    {
        CERT_STORE_PROV_MSG                = 1,
        CERT_STORE_PROV_MEMORY             = 2,
        CERT_STORE_PROV_FILE               = 3,
        CERT_STORE_PROV_REG                = 4,
        CERT_STORE_PROV_PKCS7              = 5,
        CERT_STORE_PROV_SERIALIZED         = 6,
        CERT_STORE_PROV_FILENAME_A         = 7,
        CERT_STORE_PROV_FILENAME_W         = 8,
        CERT_STORE_PROV_FILENAME           = CERT_STORE_PROV_FILENAME_W,
        CERT_STORE_PROV_SYSTEM_A           = 9,
        CERT_STORE_PROV_SYSTEM_W           = 10,
        CERT_STORE_PROV_SYSTEM             = CERT_STORE_PROV_SYSTEM_W,
        CERT_STORE_PROV_COLLECTION         = 11,
        CERT_STORE_PROV_SYSTEM_REGISTRY_A  = 12,
        CERT_STORE_PROV_SYSTEM_REGISTRY_W  = 13,
        CERT_STORE_PROV_SYSTEM_REGISTRY    = CERT_STORE_PROV_SYSTEM_REGISTRY_W,
        CERT_STORE_PROV_PHYSICAL_W         = 14,
        CERT_STORE_PROV_PHYSICAL           = CERT_STORE_PROV_PHYSICAL_W,
        CERT_STORE_PROV_SMART_CARD_W       = 15,
        CERT_STORE_PROV_SMART_CARD         = CERT_STORE_PROV_SMART_CARD_W,
        CERT_STORE_PROV_LDAP_W             = 16,
        CERT_STORE_PROV_LDAP               = CERT_STORE_PROV_LDAP_W
    }
}
'@

$pkiInternalsTypes = @'
using System;

namespace Pki
{
    public static class Period
    {
        public static TimeSpan ToTimeSpan(byte[] value)
        {
            var period = BitConverter.ToInt64(value, 0); period /= -10000000;
            return TimeSpan.FromSeconds(period);
        }

        public static byte[] ToByteArray(TimeSpan value)
        {
            var period = value.TotalSeconds;
            period *= -10000000;
            return BitConverter.GetBytes((long)period);
        }
    }
}

namespace Pki.CATemplate
{
    /// <summary>
    /// 2.27 msPKI-Private-Key-Flag Attribute
    /// https://msdn.microsoft.com/en-us/library/cc226547.aspx
    /// </summary>
    [Flags]
    public enum PrivateKeyFlags
    {
        None = 0, //This flag indicates that attestation data is not required when creating the certificate request. It also instructs the server to not add any attestation OIDs to the issued certificate. For more details, see [MS-WCCE] section 3.2.2.6.2.1.4.5.7.
        RequireKeyArchival = 1, //This flag instructs the client to create a key archival certificate request, as specified in [MS-WCCE] sections 3.1.2.4.2.2.2.8 and 3.2.2.6.2.1.4.5.7.
        AllowKeyExport = 16, //This flag instructs the client to allow other applications to copy the private key to a .pfx file, as specified in [PKCS12], at a later time.
        RequireStrongProtection = 32, //This flag instructs the client to use additional protection for the private key.
        RequireAlternateSignatureAlgorithm = 64, //This flag instructs the client to use an alternate signature format. For more details, see [MS-WCCE] section 3.1.2.4.2.2.2.8.
        ReuseKeysRenewal = 128, //This flag instructs the client to use the same key when renewing the certificate.<35>
        UseLegacyProvider = 256, //This flag instructs the client to process the msPKI-RA-Application-Policies attribute as specified in section 2.23.1.<36>
        TrustOnUse = 512, //This flag indicates that attestation based on the user's credentials is to be performed. For more details, see [MS-WCCE] section 3.2.2.6.2.1.4.5.7.
        ValidateCert = 1024, //This flag indicates that attestation based on the hardware certificate of the Trusted Platform Module (TPM) is to be performed. For more details, see [MS-WCCE] section 3.2.2.6.2.1.4.5.7.
        ValidateKey = 2048, //This flag indicates that attestation based on the hardware key of the TPM is to be performed. For more details, see [MS-WCCE] section 3.2.2.6.2.1.4.5.7.
        Preferred = 4096, //This flag informs the client that it SHOULD include attestation data if it is capable of doing so when creating the certificate request. It also instructs the server that attestation may or may not be completed before any certificates can be issued. For more details, see [MS-WCCE] sections 3.1.2.4.2.2.2.8 and 3.2.2.6.2.1.4.5.7.
        Required = 8192, //This flag informs the client that attestation data is required when creating the certificate request. It also instructs the server that attestation must be completed before any certificates can be issued. For more details, see [MS-WCCE] sections 3.1.2.4.2.2.2.8 and 3.2.2.6.2.1.4.5.7.
        WithoutPolicy = 16384, //This flag instructs the server to not add any certificate policy OIDs to the issued certificate even though attestation SHOULD be performed. For more details, see [MS-WCCE] section 3.2.2.6.2.1.4.5.7.
        xxx =  0x000F0000 
    }

    [Flags]
    public enum KeyUsage
    {
        DIGITAL_SIGNATURE = 0x80,
        NON_REPUDIATION = 0x40,
        KEY_ENCIPHERMENT = 0x20,
        DATA_ENCIPHERMENT = 0x10,
        KEY_AGREEMENT = 0x8,
        KEY_CERT_SIGN = 0x4,
        CRL_SIGN = 0x2,
        ENCIPHER_ONLY_KEY_USAGE = 0x1,
        DECIPHER_ONLY_KEY_USAGE = (0x80 << 8),
        NO_KEY_USAGE = 0x0
    }

    public enum KeySpec
    {
        KeyExchange = 1, //Keys used to encrypt/decrypt session keys
        Signature = 2 //Keys used to create and verify digital signatures.
    }

    /// <summary>
    /// 2.26 msPKI-Enrollment-Flag Attribute
    /// https://msdn.microsoft.com/en-us/library/cc226546.aspx
    /// </summary>
    [Flags]
    public enum EnrollmentFlags
    {
        IncludeSymmetricAlgorithms = 1,//This flag instructs the client and server to include a Secure/Multipurpose Internet Mail Extensions (S/MIME) certificate extension, as specified in RFC4262, in the request and in the issued certificate.  
        CAManagerApproval = 2,// This flag instructs the CA to put all requests in a pending state.  
        KraPublish = 4,// This flag instructs the CA to publish the issued certificate to the key recovery agent (KRA) container in Active Directory.  
        DsPublish = 8,// This flag instructs clients and CA servers to append the issued certificate to the userCertificate attribute, as specified in RFC4523, on the user object in Active Directory.  
        AutoenrollmentCheckDsCert = 16,// This flag instructs clients not to do autoenrollment for a certificate based on this template if the user's userCertificate attribute (specified in RFC4523) in Active Directory has a valid certificate based on the same template.  
        Autoenrollment = 32,//This flag instructs clients to perform autoenrollment for the specified template.  
        ReenrollExistingCert = 64,//This flag instructs clients to sign the renewal request using the private key of the existing certificate.
        RequireUserInteraction = 256,// This flag instructs the client to obtain user consent before attempting to enroll for a certificate that is based on the specified template.
        RemoveInvalidFromStore = 1024,// This flag instructs the autoenrollment client to delete any certificates that are no longer needed based on the specific template from the local certificate storage.
        AllowEnrollOnBehalfOf = 2048,//This flag instructs the server to allow enroll on behalf of(EOBO) functionality.
        IncludeOcspRevNoCheck = 4096,// This flag instructs the server to not include revocation information and add the id-pkix-ocsp-nocheck extension, as specified in RFC2560 section 4.2.2.2.1, to the certificate that is issued.    Windows Server 2003 - this flag is not supported.
        ReuseKeyTokenFull = 8192,//This flag instructs the client to reuse the private key for a smart card-based certificate renewal if it is unable to create a new private key on the card.Windows XP, Windows Server 2003 - this flag is not supported. NoRevocationInformation 16384 This flag instructs the server to not include revocation information in the issued certificate. Windows Server 2003, Windows Server 2008 - this flag is not supported.
        BasicConstraintsInEndEntityCerts = 32768,//This flag instructs the server to include Basic Constraints extension in the end entity certificates. Windows Server 2003, Windows Server 2008 - this flag is not supported.
        IgnoreEnrollOnReenrollment = 65536,//This flag instructs the CA to ignore the requirement for Enroll permissions on the template when processing renewal requests. Windows Server 2003, Windows Server 2008, Windows Server 2008 R2 - this flag is not supported.
        IssuancePoliciesFromRequest = 131072,//This flag indicates that the certificate issuance policies to be included in the issued certificate come from the request rather than from the template. The template contains a list of all of the issuance policies that the request is allowed to specify; if the request contains policies that are not listed in the template, then the request is rejected. Windows Server 2003, Windows Server 2008, Windows Server 2008 R2 - this flag is not supported.
    }

    /// <summary>
    /// 2.28 msPKI-Certificate-Name-Flag Attribute
    /// https://msdn.microsoft.com/en-us/library/cc226548.aspx
    /// </summary>
    [Flags]
    public enum NameFlags
    {
        EnrolleeSuppliesSubject = 1, //This flag instructs the client to supply subject information in the certificate request  
        OldCertSuppliesSubjectAndAltName = 8, //This flag instructs the client to reuse values of subject name and alternative subject name extensions from an existing valid certificate when creating a certificate renewal request. Windows Server 2003, Windows Server 2008 - this flag is not supported.
        EnrolleeSuppluiesAltSubject = 65536, //This flag instructs the client to supply subject alternate name information in the certificate request.  
        AltSubjectRequireDomainDNS = 4194304, //This flag instructs the CA to add the value of the requester's FQDN and NetBIOS name to the Subject Alternative Name extension of the issued certificate.  
        AltSubjectRequireDirectoryGUID = 16777216, //This flag instructs the CA to add the value of the objectGUID attribute from the requestor's user object in Active Directory to the Subject Alternative Name extension of the issued certificate.  
        AltSubjectRequireUPN = 33554432, //This flag instructs the CA to add the value of the UPN attribute from the requestor's user object in Active Directory to the Subject Alternative Name extension of the issued certificate.  
        AltSubjectRequireEmail = 67108864, //This flag instructs the CA to add the value of the e-mail attribute from the requestor's user object in Active Directory to the Subject Alternative Name extension of the issued certificate.  
        AltSubjectRequireDNS = 134217728, //This flag instructs the CA to add the value obtained from the DNS attribute of the requestor's user object in Active Directory to the Subject Alternative Name extension of the issued certificate.  
        SubjectRequireDNSasCN = 268435456, //This flag instructs the CA to add the value obtained from the DNS attribute of the requestor's user object in Active Directory as the CN in the subject of the issued certificate.  
        SubjectRequireEmail = 536870912, //This flag instructs the CA to add the value of the e-mail attribute from the requestor's user object in Active Directory as the subject of the issued certificate.  
        SubjectRequireCommonName = 1073741824, //This flag instructs the CA to set the subject name to the requestor's CN from Active Directory.  
        SubjectrequireDirectoryPath = -2147483648 //This flag instructs the CA to set the subject name to the requestor's distinguished name (DN) from Active Directory.
    }

    /// <summary>
    /// 2.4 flags Attribute
    /// https://msdn.microsoft.com/en-us/library/cc226550.aspx
    /// </summary>
    [Flags]
    public enum Flags
    {
        Undefined = 1, //Undefined.
        AddEmail = 2, //Reserved. All protocols MUST ignore this flag.
        Undefined2 = 4, //Undefined.
        DsPublish = 8, //Reserved. All protocols MUST ignore this flag.
        AllowKeyExport = 16, //Reserved. All protocols MUST ignore this flag.
        Autoenrollment = 32, //This flag indicates whether clients can perform autoenrollment for the specified template.
        MachineType = 64, //This flag indicates that this certificate template is for an end entity that represents a machine.
        IsCA = 128, //This flag indicates a certificate request for a CA certificate.
        AddTemplateName = 512, //This flag indicates that a certificate based on this section needs to include a template name certificate extension.
        DoNotPersistInDB = 1024, //This flag indicates that the record of a certificate request for a certificate that is issued need not be persisted by the CA. Windows Server 2003, Windows Server 2008 - this flag is not supported.
        IsCrossCA = 2048, //This flag indicates a certificate request for cross-certifying a certificate.
        IsDefault = 65536, //This flag indicates that the template SHOULD not be modified in any way.
        IsModified = 131072 //This flag indicates that the template MAY be modified if required.
    }
}
'@

$gpoType = @'
    using System;
    using System.Collections.Generic;
    using System.Runtime.CompilerServices;
    using System.Runtime.InteropServices;
    using System.Text;
    using System.Threading;
    using Microsoft.Win32;

    namespace GPO
    {
        /// <summary>
        /// Represent the result of group policy operations.
        /// </summary>
        public enum ResultCode
        {
            Succeed = 0,
            CreateOrOpenFailed = -1,
            SetFailed = -2,
            SaveFailed = -3
        }

        /// <summary>
        /// The WinAPI handler for GroupPlicy operations.
        /// </summary>
        public class WinAPIForGroupPolicy
        {
            // Group Policy Object open / creation flags
            const UInt32 GPO_OPEN_LOAD_REGISTRY = 0x00000001;    // Load the registry files
            const UInt32 GPO_OPEN_READ_ONLY = 0x00000002;    // Open the GPO as read only

            // Group Policy Object option flags
            const UInt32 GPO_OPTION_DISABLE_USER = 0x00000001;   // The user portion of this GPO is disabled
            const UInt32 GPO_OPTION_DISABLE_MACHINE = 0x00000002;   // The machine portion of this GPO is disabled

            const UInt32 REG_OPTION_NON_VOLATILE = 0x00000000;

            const UInt32 ERROR_MORE_DATA = 234;

            // You can find the Guid in <Gpedit.h>
            static readonly Guid REGISTRY_EXTENSION_GUID = new Guid("35378EAC-683F-11D2-A89A-00C04FBBCFA2");
            static readonly Guid CLSID_GPESnapIn = new Guid("8FC0B734-A0E1-11d1-A7D3-0000F87571E3");

            /// <summary>
            /// Group Policy Object type.
            /// </summary>
            enum GROUP_POLICY_OBJECT_TYPE
            {
                GPOTypeLocal = 0,                       // Default GPO on the local machine
                GPOTypeRemote,                          // GPO on a remote machine
                GPOTypeDS,                              // GPO in the Active Directory
                GPOTypeLocalUser,                       // User-specific GPO on the local machine
                GPOTypeLocalGroup                       // Group-specific GPO on the local machine
            }

            #region COM

            /// <summary>
            /// Group Policy Interface definition from COM.
            /// You can find the Guid in <Gpedit.h>
            /// </summary>
            [Guid("EA502723-A23D-11d1-A7D3-0000F87571E3"),
            InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
            interface IGroupPolicyObject
            {
                void New(
                [MarshalAs(UnmanagedType.LPWStr)] String pszDomainName,
                [MarshalAs(UnmanagedType.LPWStr)] String pszDisplayName,
                UInt32 dwFlags);

                void OpenDSGPO(
                    [MarshalAs(UnmanagedType.LPWStr)] String pszPath,
                    UInt32 dwFlags);

                void OpenLocalMachineGPO(UInt32 dwFlags);

                void OpenRemoteMachineGPO(
                    [MarshalAs(UnmanagedType.LPWStr)] String pszComputerName,
                    UInt32 dwFlags);

                void Save(
                    [MarshalAs(UnmanagedType.Bool)] bool bMachine,
                    [MarshalAs(UnmanagedType.Bool)] bool bAdd,
                    [MarshalAs(UnmanagedType.LPStruct)] Guid pGuidExtension,
                    [MarshalAs(UnmanagedType.LPStruct)] Guid pGuid);

                void Delete();

                void GetName(
                    [MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszName,
                    Int32 cchMaxLength);

                void GetDisplayName(
                    [MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszName,
                    Int32 cchMaxLength);

                void SetDisplayName([MarshalAs(UnmanagedType.LPWStr)] String pszName);

                void GetPath(
                    [MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszPath,
                    Int32 cchMaxPath);

                void GetDSPath(
                    UInt32 dwSection,
                    [MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszPath,
                    Int32 cchMaxPath);

                void GetFileSysPath(
                    UInt32 dwSection,
                    [MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszPath,
                    Int32 cchMaxPath);

                UInt32 GetRegistryKey(UInt32 dwSection);

                Int32 GetOptions();

                void SetOptions(UInt32 dwOptions, UInt32 dwMask);

                void GetType(out GROUP_POLICY_OBJECT_TYPE gpoType);

                void GetMachineName(
                    [MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszName,
                    Int32 cchMaxLength);

                UInt32 GetPropertySheetPages(out IntPtr hPages);
            }

            /// <summary>
            /// Group Policy Class definition from COM.
            /// You can find the Guid in <Gpedit.h>
            /// </summary>
            [ComImport, Guid("EA502722-A23D-11d1-A7D3-0000F87571E3")]
            class GroupPolicyObject { }

            #endregion

            #region WinAPI You can find definition of API for C# on: http://pinvoke.net/

            /// <summary>
            /// Opens the specified registry key. Note that key names are not case sensitive.
            /// </summary>
            /// See http://msdn.microsoft.com/en-us/library/ms724897(VS.85).aspx for more info about the parameters.<br/>
            [DllImport("advapi32.dll", CharSet = CharSet.Auto)]
            public static extern Int32 RegOpenKeyEx(
            UIntPtr hKey,
            String subKey,
            Int32 ulOptions,
            RegSAM samDesired,
            out UIntPtr hkResult);

            /// <summary>
            /// Retrieves the type and data for the specified value name associated with an open registry key.
            /// </summary>
            /// See http://msdn.microsoft.com/en-us/library/ms724911(VS.85).aspx for more info about the parameters and return value.<br/>
            [DllImport("advapi32.dll", CharSet = CharSet.Unicode, EntryPoint = "RegQueryValueExW", SetLastError = true)]
            static extern Int32 RegQueryValueEx(
            UIntPtr hKey,
            String lpValueName,
            Int32 lpReserved,
            out UInt32 lpType,
            [Out] byte[] lpData,
            ref UInt32 lpcbData);

            /// <summary>
            /// Sets the data and type of a specified value under a registry key.
            /// </summary>
            /// See http://msdn.microsoft.com/en-us/library/ms724923(VS.85).aspx for more info about the parameters and return value.<br/>
            [DllImport("advapi32.dll", SetLastError = true)]
            static extern Int32 RegSetValueEx(
            UInt32 hKey,
            [MarshalAs(UnmanagedType.LPStr)] String lpValueName,
            Int32 Reserved,
            Microsoft.Win32.RegistryValueKind dwType,
            IntPtr lpData,
            Int32 cbData);

            /// <summary>
            /// Creates the specified registry key. If the key already exists, the function opens it. Note that key names are not case sensitive.
            /// </summary>
            /// See http://msdn.microsoft.com/en-us/library/ms724844(v=VS.85).aspx for more info about the parameters and return value.<br/>
            [DllImport("advapi32.dll", SetLastError = true)]
            static extern Int32 RegCreateKeyEx(
            UInt32 hKey,
            String lpSubKey,
            UInt32 Reserved,
            String lpClass,
            RegOption dwOptions,
            RegSAM samDesired,
            IntPtr lpSecurityAttributes,
            out UInt32 phkResult,
            out RegResult lpdwDisposition);

            /// <summary>
            /// Closes a handle to the specified registry key.
            /// </summary>
            /// See http://msdn.microsoft.com/en-us/library/ms724837(VS.85).aspx for more info about the parameters and return value.<br/>
            [DllImport("advapi32.dll", SetLastError = true)]
            static extern Int32 RegCloseKey(
            UInt32 hKey);

            /// <summary>
            /// Deletes a subkey and its values from the specified platform-specific view of the registry. Note that key names are not case sensitive.
            /// </summary>
            /// See http://msdn.microsoft.com/en-us/library/ms724847(VS.85).aspx for more info about the parameters and return value.<br/>
            [DllImport("advapi32.dll", EntryPoint = "RegDeleteKeyEx", SetLastError = true)]
            public static extern Int32 RegDeleteKeyEx(
            UInt32 hKey,
            String lpSubKey,
            RegSAM samDesired,
            UInt32 Reserved);

            #endregion

            /// <summary>
            /// Registry creating volatile check.
            /// </summary>
            [Flags]
            public enum RegOption
            {
                NonVolatile = 0x0,
                Volatile = 0x1,
                CreateLink = 0x2,
                BackupRestore = 0x4,
                OpenLink = 0x8
            }

            /// <summary>
            /// Access mask the specifies the platform-specific view of the registry.
            /// </summary>
            [Flags]
            public enum RegSAM
            {
                QueryValue = 0x00000001,
                SetValue = 0x00000002,
                CreateSubKey = 0x00000004,
                EnumerateSubKeys = 0x00000008,
                Notify = 0x00000010,
                CreateLink = 0x00000020,
                WOW64_32Key = 0x00000200,
                WOW64_64Key = 0x00000100,
                WOW64_Res = 0x00000300,
                Read = 0x00020019,
                Write = 0x00020006,
                Execute = 0x00020019,
                AllAccess = 0x000f003f
            }

            /// <summary>
            /// Structure for security attributes.
            /// </summary>
            [StructLayout(LayoutKind.Sequential)]
            public struct SECURITY_ATTRIBUTES
            {
                public Int32 nLength;
                public IntPtr lpSecurityDescriptor;
                public Int32 bInheritHandle;
            }

            /// <summary>
            /// Flag returned by calling RegCreateKeyEx.
            /// </summary>
            public enum RegResult
            {
                CreatedNewKey = 0x00000001,
                OpenedExistingKey = 0x00000002
            }

            /// <summary>
            /// Class to create an object to handle the group policy operation.
            /// </summary>
            public class GroupPolicyObjectHandler
            {
                public const Int32 REG_NONE = 0;
                public const Int32 REG_SZ = 1;
                public const Int32 REG_EXPAND_SZ = 2;
                public const Int32 REG_BINARY = 3;
                public const Int32 REG_DWORD = 4;
                public const Int32 REG_DWORD_BIG_ENDIAN = 5;
                public const Int32 REG_MULTI_SZ = 7;
                public const Int32 REG_QWORD = 11;

                // Group Policy interface handler
                IGroupPolicyObject iGroupPolicyObject;
                // Group Policy object handler.
                GroupPolicyObject groupPolicyObject;

                #region constructor

                /// <summary>
                /// Constructor.
                /// </summary>
                /// <param name="remoteMachineName">Target machine name to operate group policy</param>
                /// <exception cref="System.Runtime.InteropServices.COMException">Throw when com execution throws exceptions</exception>
                public GroupPolicyObjectHandler(String remoteMachineName)
                {
                    groupPolicyObject = new GroupPolicyObject();
                    iGroupPolicyObject = (IGroupPolicyObject)groupPolicyObject;
                    try
                    {
                        if (String.IsNullOrEmpty(remoteMachineName))
                        {
                            iGroupPolicyObject.OpenLocalMachineGPO(GPO_OPEN_LOAD_REGISTRY);
                        }
                        else
                        {
                            iGroupPolicyObject.OpenRemoteMachineGPO(remoteMachineName, GPO_OPEN_LOAD_REGISTRY);
                        }
                    }
                    catch (COMException e)
                    {
                        throw e;
                    }
                }

                #endregion

                #region interface related methods

                /// <summary>
                /// Retrieves the display name for the GPO.
                /// </summary>
                /// <returns>Display name</returns>
                /// <exception cref="System.Runtime.InteropServices.COMException">Throw when com execution throws exceptions</exception>
                public String GetDisplayName()
                {
                    StringBuilder pszName = new StringBuilder(Byte.MaxValue);
                    try
                    {
                        iGroupPolicyObject.GetDisplayName(pszName, Byte.MaxValue);
                    }
                    catch (COMException e)
                    {
                        throw e;
                    }
                    return pszName.ToString();
                }

                /// <summary>
                /// Retrieves the computer name of the remote GPO.
                /// </summary>
                /// <returns>Machine name</returns>
                /// <exception cref="System.Runtime.InteropServices.COMException">Throw when com execution throws exceptions</exception>
                public String GetMachineName()
                {
                    StringBuilder pszName = new StringBuilder(Byte.MaxValue);
                    try
                    {
                        iGroupPolicyObject.GetMachineName(pszName, Byte.MaxValue);
                    }
                    catch (COMException e)
                    {
                        throw e;
                    }
                    return pszName.ToString();
                }

                /// <summary>
                /// Retrieves the options for the GPO.
                /// </summary>
                /// <returns>Options flag</returns>
                /// <exception cref="System.Runtime.InteropServices.COMException">Throw when com execution throws exceptions</exception>
                public Int32 GetOptions()
                {
                    try
                    {
                        return iGroupPolicyObject.GetOptions();
                    }
                    catch (COMException e)
                    {
                        throw e;
                    }
                }

                /// <summary>
                /// Retrieves the path to the GPO.
                /// </summary>
                /// <returns>The path to the GPO</returns>
                /// <exception cref="System.Runtime.InteropServices.COMException">Throw when com execution throws exceptions</exception>
                public String GetPath()
                {
                    StringBuilder pszName = new StringBuilder(Byte.MaxValue);
                    try
                    {
                        iGroupPolicyObject.GetPath(pszName, Byte.MaxValue);
                    }
                    catch (COMException e)
                    {
                        throw e;
                    }
                    return pszName.ToString();
                }

                /// <summary>
                /// Retrieves a handle to the root of the registry key for the machine section.
                /// </summary>
                /// <returns>A handle to the root of the registry key for the specified GPO computer section</returns>
                /// <exception cref="System.Runtime.InteropServices.COMException">Throw when com execution throws exceptions</exception>
                public UInt32 GetMachineRegistryKey()
                {
                    UInt32 handle;
                    try
                    {
                        handle = iGroupPolicyObject.GetRegistryKey(GPO_OPTION_DISABLE_MACHINE);
                    }
                    catch (COMException e)
                    {
                        throw e;
                    }
                    return handle;
                }

                /// <summary>
                /// Retrieves a handle to the root of the registry key for the user section.
                /// </summary>
                /// <returns>A handle to the root of the registry key for the specified GPO user section</returns>
                /// <exception cref="System.Runtime.InteropServices.COMException">Throw when com execution throws exceptions</exception>
                public UInt32 GetUserRegistryKey()
                {
                    UInt32 handle;
                    try
                    {
                        handle = iGroupPolicyObject.GetRegistryKey(GPO_OPTION_DISABLE_USER);
                    }
                    catch (COMException e)
                    {
                        throw e;
                    }
                    return handle;
                }

                /// <summary>
                /// Saves the specified registry policy settings to disk and updates the revision number of the GPO.
                /// </summary>
                /// <param name="isMachine">Specifies the registry policy settings to be saved. If this parameter is TRUE, the computer policy settings are saved. Otherwise, the user policy settings are saved.</param>
                /// <param name="isAdd">Specifies whether this is an add or delete operation. If this parameter is FALSE, the last policy setting for the specified extension pGuidExtension is removed. In all other cases, this parameter is TRUE.</param>
                /// <exception cref="System.Runtime.InteropServices.COMException">Throw when com execution throws exceptions</exception>
                public void Save(bool isMachine, bool isAdd)
                {
                    try
                    {
                        iGroupPolicyObject.Save(isMachine, isAdd, REGISTRY_EXTENSION_GUID, CLSID_GPESnapIn);
                    }
                    catch (COMException e)
                    {
                        throw e;
                    }
                }

                #endregion

                #region customized methods

                /// <summary>
                /// Set the group policy value.
                /// </summary>
                /// <param name="isMachine">Specifies the registry policy settings to be saved. If this parameter is TRUE, the computer policy settings are saved. Otherwise, the user policy settings are saved.</param>
                /// <param name="subKey">Group policy config full path</param>
                /// <param name="valueName">Group policy config key name</param>
                /// <param name="value">If value is null, it will envoke the delete method</param>
                /// <returns>Whether the config is successfully set</returns>
                public ResultCode SetGroupPolicy(bool isMachine, String subKey, String valueName, object value)
                {
                    UInt32 gphKey = (isMachine) ? GetMachineRegistryKey() : GetUserRegistryKey();
                    UInt32 gphSubKey;
                    UIntPtr hKey;
                    RegResult flag;

                    if (null == value)
                    {
                        // check the key's existance
                        if (RegOpenKeyEx((UIntPtr)gphKey, subKey, 0, RegSAM.QueryValue, out hKey) == 0)
                        {
                            RegCloseKey((UInt32)hKey);
                            // delete the GPO
                            Int32 hr = RegDeleteKeyEx(
                            gphKey,
                            subKey,
                            RegSAM.Write,
                            0);
                            if (0 != hr)
                            {
                                RegCloseKey(gphKey);
                                return ResultCode.CreateOrOpenFailed;
                            }
                            Save(isMachine, false);
                        }
                        else
                        {
                            // not exist
                        }

                    }
                    else
                    {
                        // set the GPO
                        Int32 hr = RegCreateKeyEx(
                        gphKey,
                        subKey,
                        0,
                        null,
                        RegOption.NonVolatile,
                        RegSAM.Write,
                        IntPtr.Zero,
                        out gphSubKey,
                        out flag);
                        if (0 != hr)
                        {
                            RegCloseKey(gphSubKey);
                            RegCloseKey(gphKey);
                            return ResultCode.CreateOrOpenFailed;
                        }

                        Int32 cbData = 4;
                        IntPtr keyValue = IntPtr.Zero;

                        if (value.GetType() == typeof(Int32))
                        {
                            keyValue = Marshal.AllocHGlobal(cbData);
                            Marshal.WriteInt32(keyValue, (Int32)value);
                            hr = RegSetValueEx(gphSubKey, valueName, 0, RegistryValueKind.DWord, keyValue, cbData);
                        }
                        else if (value.GetType() == typeof(String))
                        {
                            keyValue = Marshal.StringToHGlobalAnsi(value.ToString());
                            cbData = System.Text.Encoding.UTF8.GetByteCount(value.ToString()) + 1;
                            hr = RegSetValueEx(gphSubKey, valueName, 0, RegistryValueKind.String, keyValue, cbData);
                        }
                        else
                        {
                            RegCloseKey(gphSubKey);
                            RegCloseKey(gphKey);
                            return ResultCode.SetFailed;
                        }

                        if (0 != hr)
                        {
                            RegCloseKey(gphSubKey);
                            RegCloseKey(gphKey);
                            return ResultCode.SetFailed;
                        }
                        try
                        {
                            Save(isMachine, true);
                        }
                        catch (COMException e)
                        {
                            RegCloseKey(gphSubKey);
                            RegCloseKey(gphKey);
                            return ResultCode.SaveFailed;
                        }
                        RegCloseKey(gphSubKey);
                        RegCloseKey(gphKey);
                    }

                    return ResultCode.Succeed;
                }

                /// <summary>
                /// Get the config of the group policy.
                /// </summary>
                /// <param name="isMachine">Specifies the registry policy settings to be saved. If this parameter is TRUE, get from the computer policy settings. Otherwise, get from the user policy settings.</param>
                /// <param name="subKey">Group policy config full path</param>
                /// <param name="valueName">Group policy config key name</param>
                /// <returns>The setting of the specified config</returns>
                public object GetGroupPolicy(bool isMachine, String subKey, String valueName)
                {
                    UIntPtr gphKey = (UIntPtr)((isMachine) ? GetMachineRegistryKey() : GetUserRegistryKey());
                    UIntPtr hKey;
                    object keyValue = null;
                    UInt32 size = 1;

                    if (RegOpenKeyEx(gphKey, subKey, 0, RegSAM.QueryValue, out hKey) == 0)
                    {
                        UInt32 type;
                        byte[] data = new byte[size];  // to store retrieved the value's data

                        if (RegQueryValueEx(hKey, valueName, 0, out type, data, ref size) == 234)
                        {
                            //size retreived
                            data = new byte[size]; //redefine data
                        }

                        if (RegQueryValueEx(hKey, valueName, 0, out type, data, ref size) != 0)
                        {
                            return null;
                        }

                        switch (type)
                        {
                            case REG_NONE:
                            case REG_BINARY:
                                keyValue = data;
                                break;
                            case REG_DWORD:
                                keyValue = (((data[0] | (data[1] << 8)) | (data[2] << 16)) | (data[3] << 24));
                                break;
                            case REG_DWORD_BIG_ENDIAN:
                                keyValue = (((data[3] | (data[2] << 8)) | (data[1] << 16)) | (data[0] << 24));
                                break;
                            case REG_QWORD:
                                {
                                    UInt32 numLow = (UInt32)(((data[0] | (data[1] << 8)) | (data[2] << 16)) | (data[3] << 24));
                                    UInt32 numHigh = (UInt32)(((data[4] | (data[5] << 8)) | (data[6] << 16)) | (data[7] << 24));
                                    keyValue = (long)(((ulong)numHigh << 32) | (ulong)numLow);
                                    break;
                                }
                            case REG_SZ:
                                var s = Encoding.Unicode.GetString(data, 0, (Int32)size);
                                keyValue = s.Substring(0, s.Length - 1);
                                break;
                            case REG_EXPAND_SZ:
                                keyValue = Environment.ExpandEnvironmentVariables(Encoding.Unicode.GetString(data, 0, (Int32)size));
                                break;
                            case REG_MULTI_SZ:
                                {
                                    List<string> strings = new List<String>();
                                    String packed = Encoding.Unicode.GetString(data, 0, (Int32)size);
                                    Int32 start = 0;
                                    Int32 end = packed.IndexOf("", start);
                                    while (end > start)
                                    {
                                        strings.Add(packed.Substring(start, end - start));
                                        start = end + 1;
                                        end = packed.IndexOf("", start);
                                    }
                                    keyValue = strings.ToArray();
                                    break;
                                }
                            default:
                                throw new NotSupportedException();
                        }

                        RegCloseKey((UInt32)hKey);
                    }

                    return keyValue;
                }

                #endregion

            }
        }

        public class Helper
        {
            private static object _returnValueFromSet, _returnValueFromGet;

            /// <summary>
            /// Set policy config
            /// It will start a single thread to set group policy.
            /// </summary>
            /// <param name="isMachine">Whether is machine config</param>
            /// <param name="configFullPath">The full path configuration</param>
            /// <param name="configKey">The configureation key name</param>
            /// <param name="value">The value to set, boxed with proper type [ String, Int32 ]</param>
            /// <returns>Whether the config is successfully set</returns>
            [MethodImplAttribute(MethodImplOptions.Synchronized)]
            public static ResultCode SetGroupPolicy(bool isMachine, String configFullPath, String configKey, object value)
            {
                Thread worker = new Thread(SetGroupPolicy);
                worker.SetApartmentState(ApartmentState.STA);
                worker.Start(new object[] { isMachine, configFullPath, configKey, value });
                worker.Join();
                return (ResultCode)_returnValueFromSet;
            }

            /// <summary>
            /// Thread start for seting group policy.
            /// Called by public static ResultCode SetGroupPolicy(bool isMachine, WinRMGPConfigName configName, object value)
            /// </summary>
            /// <param name="values">
            /// values[0] - isMachine<br/>
            /// values[1] - configFullPath<br/>
            /// values[2] - configKey<br/>
            /// values[3] - value<br/>
            /// </param>
            private static void SetGroupPolicy(object values)
            {
                object[] valueList = (object[])values;
                bool isMachine = (bool)valueList[0];
                String configFullPath = (String)valueList[1];
                String configKey = (String)valueList[2];
                object value = valueList[3];

                WinAPIForGroupPolicy.GroupPolicyObjectHandler gpHandler = new WinAPIForGroupPolicy.GroupPolicyObjectHandler(null);

                _returnValueFromSet = gpHandler.SetGroupPolicy(isMachine, configFullPath, configKey, value);
            }

            /// <summary>
            /// Get policy config.
            /// It will start a single thread to get group policy
            /// </summary>
            /// <param name="isMachine">Whether is machine config</param>
            /// <param name="configFullPath">The full path configuration</param>
            /// <param name="configKey">The configureation key name</param>
            /// <returns>The group policy setting</returns>
            [MethodImplAttribute(MethodImplOptions.Synchronized)]
            public static object GetGroupPolicy(bool isMachine, String configFullPath, String configKey)
            {
                Thread worker = new Thread(GetGroupPolicy);
                worker.SetApartmentState(ApartmentState.STA);
                worker.Start(new object[] { isMachine, configFullPath, configKey });
                worker.Join();
                return _returnValueFromGet;
            }

            /// <summary>
            /// Thread start for geting group policy.
            /// Called by public static object GetGroupPolicy(bool isMachine, WinRMGPConfigName configName)
            /// </summary>
            /// <param name="values">
            /// values[0] - isMachine<br/>
            /// values[1] - configFullPath<br/>
            /// values[2] - configKey<br/>
            /// </param>
            public static void GetGroupPolicy(object values)
            {
                object[] valueList = (object[])values;
                bool isMachine = (bool)valueList[0];
                String configFullPath = (String)valueList[1];
                String configKey = (String)valueList[2];

                WinAPIForGroupPolicy.GroupPolicyObjectHandler gpHandler = new WinAPIForGroupPolicy.GroupPolicyObjectHandler(null);

                _returnValueFromGet = gpHandler.GetGroupPolicy(isMachine, configFullPath, configKey);
            }
        }
    }
'@
#endregion .net Types

$ApplicationPolicies = @{
    # Remote Desktop
    'Remote Desktop' = '1.3.6.1.4.1.311.54.1.2'
    # Windows Update
    'Windows Update' = '1.3.6.1.4.1.311.76.6.1'
    # Windows Third Party Applicaiton Component
    'Windows Third Party Application Component' = '1.3.6.1.4.1.311.10.3.25'
    # Windows TCB Component
    'Windows TCB Component' = '1.3.6.1.4.1.311.10.3.23'
    # Windows Store
    'Windows Store' = '1.3.6.1.4.1.311.76.3.1'
    # Windows Software Extension verification
    ' Windows Software Extension Verification' = '1.3.6.1.4.1.311.10.3.26'
    # Windows RT Verification
    'Windows RT Verification' = '1.3.6.1.4.1.311.10.3.21'
    # Windows Kits Component
    'Windows Kits Component' = '1.3.6.1.4.1.311.10.3.20'
    # ROOT_PROGRAM_NO_OCSP_FAILOVER_TO_CRL
    'No OCSP Failover to CRL' = '1.3.6.1.4.1.311.60.3.3'
    # ROOT_PROGRAM_AUTO_UPDATE_END_REVOCATION
    'Auto Update End Revocation' = '1.3.6.1.4.1.311.60.3.2'
    # ROOT_PROGRAM_AUTO_UPDATE_CA_REVOCATION
    'Auto Update CA Revocation' = '1.3.6.1.4.1.311.60.3.1'
    # Revoked List Signer
    'Revoked List Signer' = '1.3.6.1.4.1.311.10.3.19'
    # Protected Process Verification
    'Protected Process Verification' = '1.3.6.1.4.1.311.10.3.24'
    # Protected Process Light Verification
    'Protected Process Light Verification' = '1.3.6.1.4.1.311.10.3.22'
    # Platform Certificate
    'Platform Certificate' = '2.23.133.8.2'
    # Microsoft Publisher
    'Microsoft Publisher' = '1.3.6.1.4.1.311.76.8.1'
    # Kernel Mode Code Signing
    'Kernel Mode Code Signing' = '1.3.6.1.4.1.311.6.1.1'
    # HAL Extension
    'HAL Extension' = '1.3.6.1.4.1.311.61.5.1'
    # Endorsement Key Certificate
    'Endorsement Key Certificate' = '2.23.133.8.1'
    # Early Launch Antimalware Driver
    'Early Launch Antimalware Driver' = '1.3.6.1.4.1.311.61.4.1'
    # Dynamic Code Generator
    'Dynamic Code Generator' = '1.3.6.1.4.1.311.76.5.1'
    # Domain Name System (DNS) Server Trust
    'DNS Server Trust' = '1.3.6.1.4.1.311.64.1.1'
    # Document Encryption
    'Document Encryption' = '1.3.6.1.4.1.311.80.1'
    # Disallowed List
    'Disallowed List' = '1.3.6.1.4.1.10.3.30'
    # Attestation Identity Key Certificate
    # System Health Authentication
    'System Health Authentication' = '1.3.6.1.4.1.311.47.1.1'
    # Smartcard Logon
    'IdMsKpScLogon' = '1.3.6.1.4.1.311.20.2.2'
    # Certificate Request Agent
    'ENROLLMENT_AGENT' = '1.3.6.1.4.1.311.20.2.1'
    # CTL Usage
    'AUTO_ENROLL_CTL_USAGE' = '1.3.6.1.4.1.311.20.1'
    # Private Key Archival
    'KP_CA_EXCHANGE' = '1.3.6.1.4.1.311.21.5'
    # Key Recovery Agent
    'KP_KEY_RECOVERY_AGENT' = '1.3.6.1.4.1.311.21.6'
    # Secure Email
    'PKIX_KP_EMAIL_PROTECTION' = '1.3.6.1.5.5.7.3.4'
    # IP Security End System
    'PKIX_KP_IPSEC_END_SYSTEM' = '1.3.6.1.5.5.7.3.5'
    # IP Security Tunnel Termination
    'PKIX_KP_IPSEC_TUNNEL' = '1.3.6.1.5.5.7.3.6'
    # IP Security User
    'PKIX_KP_IPSEC_USER' = '1.3.6.1.5.5.7.3.7'
    # Time Stamping
    'PKIX_KP_TIMESTAMP_SIGNING' = '1.3.6.1.5.5.7.3.8'
    # OCSP Signing
    'KP_OCSP_SIGNING' = '1.3.6.1.5.5.7.3.9'
    # IP security IKE intermediate
    'IPSEC_KP_IKE_INTERMEDIATE' = '1.3.6.1.5.5.8.2.2'
    # Microsoft Trust List Signing
    'KP_CTL_USAGE_SIGNING' = '1.3.6.1.4.1.311.10.3.1'
    # Microsoft Time Stamping
    'KP_TIME_STAMP_SIGNING' = '1.3.6.1.4.1.311.10.3.2'
    # Windows Hardware Driver Verification
    'WHQL_CRYPTO' = '1.3.6.1.4.1.311.10.3.5'
    # Windows System Component Verification
    'NT5_CRYPTO' = '1.3.6.1.4.1.311.10.3.6'
    # OEM Windows System Component Verification
    'OEM_WHQL_CRYPTO' = '1.3.6.1.4.1.311.10.3.7'
    # Embedded Windows System Component Verification
    'EMBEDDED_NT_CRYPTO' = '1.3.6.1.4.1.311.10.3.8'
    # Root List Signer
    'ROOT_LIST_SIGNER' = '1.3.6.1.4.1.311.10.3.9'
    # Qualified Subordination
    'KP_QUALIFIED_SUBORDINATION' = '1.3.6.1.4.1.311.10.3.10'
    # Key Recovery
    'KP_KEY_RECOVERY' = '1.3.6.1.4.1.311.10.3.11'
    # Document Signing
    'KP_DOCUMENT_SIGNING' = '1.3.6.1.4.1.311.10.3.12'
    # Lifetime Signing
    'KP_LIFETIME_SIGNING' = '1.3.6.1.4.1.311.10.3.13'
    'DRM' = '1.3.6.1.4.1.311.10.5.1'
    'DRM_INDIVIDUALIZATION' = '1.3.6.1.4.1.311.10.5.2'
    # Key Pack Licenses
    'LICENSES' = '1.3.6.1.4.1.311.10.6.1'
    # License Server Verification
    'LICENSE_SERVER' = '1.3.6.1.4.1.311.10.6.2'
    'Server Authentication' = '1.3.6.1.5.5.7.3.1' #The certificate can be used for OCSP authentication.            
    KP_IPSEC_USER = '1.3.6.1.5.5.7.3.7' #The certificate can be used for an IPSEC user.            
    'Code Signing' = '1.3.6.1.5.5.7.3.3' #The certificate can be used for signing code.
    'Client Authentication' = '1.3.6.1.5.5.7.3.2' #The certificate can be used for authenticating a client.
    KP_EFS = '1.3.6.1.4.1.311.10.3.4' #The certificate can be used to encrypt files by using the Encrypting File System.
    EFS_RECOVERY = '1.3.6.1.4.1.311.10.3.4.1' #The certificate can be used for recovery of documents protected by using Encrypting File System (EFS).
    DS_EMAIL_REPLICATION = '1.3.6.1.4.1.311.21.19' #The certificate can be used for Directory Service email replication.         
    ANY_APPLICATION_POLICY = '1.3.6.1.4.1.311.10.12.1' #The applications that can use the certificate are not restricted.
}

$ExtendedKeyUsages = @{
    OldAuthorityKeyIdentifier = '.29.1'
    OldPrimaryKeyAttributes = '2.5.29.2'
    OldCertificatePolicies = '2.5.29.3'
    PrimaryKeyUsageRestriction = '2.5.29.4'
    SubjectDirectoryAttributes = '2.5.29.9'
    SubjectKeyIdentifier = '2.5.29.14'
    KeyUsage = '2.5.29.15'
    PrivateKeyUsagePeriod = '2.5.29.16'
    SubjectAlternativeName = '2.5.29.17'
    IssuerAlternativeName = '2.5.29.18'
    BasicConstraints = '2.5.29.19'
    CRLNumber = '2.5.29.20'
    Reasoncode = '2.5.29.21'
    HoldInstructionCode = '2.5.29.23'
    InvalidityDate = '2.5.29.24'
    DeltaCRLindicator = '2.5.29.27'
    IssuingDistributionPoint = '2.5.29.28'
    CertificateIssuer = '2.5.29.29'
    NameConstraints = '2.5.29.30'
    CRLDistributionPoints = '2.5.29.31'
    CertificatePolicies = '2.5.29.32'
    PolicyMappings = '2.5.29.33'
    AuthorityKeyIdentifier = '2.5.29.35'
    PolicyConstraints = '2.5.29.36'
    Extendedkeyusage = '2.5.29.37'
    FreshestCRL = '2.5.29.46'
    X509version3CertificateExtensionInhibitAny = '2.5.29.54'
}

#endregion Internals

#region Get-CallerPreference
function Get-CallerPreference
{
    <#
            .Synopsis
            Fetches "Preference" variable values from the caller's scope.
            .DESCRIPTION
            Script module functions do not automatically inherit their caller's variables, but they can be
            obtained through the $PSCmdlet variable in Advanced Functions.  This function is a helper function
            for any script module Advanced Function; by passing in the values of $ExecutionContext.SessionState
            and $PSCmdlet, Get-CallerPreference will set the caller's preference variables locally.
            .PARAMETER Cmdlet
            The $PSCmdlet object from a script module Advanced Function.
            .PARAMETER SessionState
            The $ExecutionContext.SessionState object from a script module Advanced Function.  This is how the
            Get-CallerPreference function sets variables in its callers' scope, even if that caller is in a different
            script module.
            .PARAMETER Name
            Optional array of parameter names to retrieve from the caller's scope.  Default is to retrieve all
            Preference variables as defined in the about_Preference_Variables help file (as of PowerShell 4.0)
            This parameter may also specify names of variables that are not in the about_Preference_Variables
            help file, and the function will retrieve and set those as well.
            .EXAMPLE
            Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState

            Imports the default PowerShell preference variables from the caller into the local scope.
            .EXAMPLE
            Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState -Name 'ErrorActionPreference','SomeOtherVariable'

            Imports only the ErrorActionPreference and SomeOtherVariable variables into the local scope.
            .EXAMPLE
            'ErrorActionPreference','SomeOtherVariable' | Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState

            Same as Example 2, but sends variable names to the Name parameter via pipeline input.
            .INPUTS
            String
            .OUTPUTS
            None.  This function does not produce pipeline output.
            .LINK
            about_Preference_Variables
    #>

    [CmdletBinding(DefaultParameterSetName = 'AllVariables')]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateScript({ $_.GetType().FullName -eq 'System.Management.Automation.PSScriptCmdlet' })]
        $Cmdlet,

        [Parameter(Mandatory = $true)]
        [System.Management.Automation.SessionState]
        $SessionState,

        [Parameter(ParameterSetName = 'Filtered', ValueFromPipeline = $true)]
        [string[]]
        $Name
    )

    begin
    {
        $filterHash = @{}
    }
    
    process
    {
        if ($null -ne $Name)
        {
            foreach ($string in $Name)
            {
                $filterHash[$string] = $true
            }
        }
    }

    end
    {
        # List of preference variables taken from the about_Preference_Variables help file in PowerShell version 4.0

        $vars = @{
            'ErrorView' = $null
            'FormatEnumerationLimit' = $null
            'LogCommandHealthEvent' = $null
            'LogCommandLifecycleEvent' = $null
            'LogEngineHealthEvent' = $null
            'LogEngineLifecycleEvent' = $null
            'LogProviderHealthEvent' = $null
            'LogProviderLifecycleEvent' = $null
            'MaximumAliasCount' = $null
            'MaximumDriveCount' = $null
            'MaximumErrorCount' = $null
            'MaximumFunctionCount' = $null
            'MaximumHistoryCount' = $null
            'MaximumVariableCount' = $null
            'OFS' = $null
            'OutputEncoding' = $null
            'ProgressPreference' = $null
            'PSDefaultParameterValues' = $null
            'PSEmailServer' = $null
            'PSModuleAutoLoadingPreference' = $null
            'PSSessionApplicationName' = $null
            'PSSessionConfigurationName' = $null
            'PSSessionOption' = $null

            'ErrorActionPreference' = 'ErrorAction'
            'DebugPreference' = 'Debug'
            'ConfirmPreference' = 'Confirm'
            'WhatIfPreference' = 'WhatIf'
            'VerbosePreference' = 'Verbose'
            'WarningPreference' = 'WarningAction'
        }


        foreach ($entry in $vars.GetEnumerator())
        {
            if (([string]::IsNullOrEmpty($entry.Value) -or -not $Cmdlet.MyInvocation.BoundParameters.ContainsKey($entry.Value)) -and
            ($PSCmdlet.ParameterSetName -eq 'AllVariables' -or $filterHash.ContainsKey($entry.Name)))
            {
                $variable = $Cmdlet.SessionState.PSVariable.Get($entry.Key)
                
                if ($null -ne $variable)
                {
                    if ($SessionState -eq $ExecutionContext.SessionState)
                    {
                        Set-Variable -Scope 1 -Name $variable.Name -Value $variable.Value -Force -Confirm:$false -WhatIf:$false
                    }
                    else
                    {
                        $SessionState.PSVariable.Set($variable.Name, $variable.Value)
                    }
                }
            }
        }

        if ($PSCmdlet.ParameterSetName -eq 'Filtered')
        {
            foreach ($varName in $filterHash.Keys)
            {
                if (-not $vars.ContainsKey($varName))
                {
                    $variable = $Cmdlet.SessionState.PSVariable.Get($varName)
                
                    if ($null -ne $variable)
                    {
                        if ($SessionState -eq $ExecutionContext.SessionState)
                        {
                            Set-Variable -Scope 1 -Name $variable.Name -Value $variable.Value -Force -Confirm:$false -WhatIf:$false
                        }
                        else
                        {
                            $SessionState.PSVariable.Set($variable.Name, $variable.Value)
                        }
                    }
                }
            }
        }

    }
}
#endregion Get-CallerPreference

#region Write-LogFunctionEntry
function Write-LogFunctionEntry
{
    [CmdletBinding()]
    param()

    $Global:LogFunctionEntryTime = Get-Date
    
    Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState

    if (!$Log)
    {
        if ($MyInvocation.MyCommand.Module.PrivateData.AutoStart)
        {
            Write-Verbose 'starting log'
            Start-Log -UseDefaults
        }
        else
        {
            Microsoft.PowerShell.Utility\Write-Verbose 'Cannot write to the log file until Start-Log has been called'
            return
        }
    }
    
    $Message = 'Entering...'
    
    $caller = (Get-PSCallStack)[1]
    $callerFunctionName = $caller.Command
    if ($caller.ScriptName)
    {
        $callerScriptName = Split-Path -Path $caller.ScriptName -Leaf
    }
    
    try
    {
        $boundParameters = $caller.InvocationInfo.BoundParameters.GetEnumerator()
    }
    catch
    {
        
    }
    
    $Message += ' ('
    foreach ($parameter in $boundParameters)
    {
        if (-not $parameter.Value)
        {
            $Message += '{0}={1},' -f $parameter.Key, '<null>'
        }
        elseif ($parameter.Value -is [System.Array] -and $parameter.Value[0] -is [string] -and $parameter.count -eq 1)
        {
            $Message += "{0}={1}," -f $parameter.Key, $($parameter.value)
        }
        elseif ($parameter.Value -is [System.Array])
        {
            $Message += '{0}={1}({2}),' -f $parameter.Key, $parameter.Value, $parameter.Value.Count
        }
        else
        {
            if ($defaults.TruncateTypes -contains $parameter.Value.GetType().FullName)
            {
                if ($parameter.Value.ToString().Length -lt $defaults.TruncateLength)
                {
                    $truncateLength = $parameter.Value.ToString().Length
                }
                else
                {
                    $truncateLength = $defaults.TruncateLength
                }
                $Message += '{0}={1},' -f $parameter.Key, $parameter.Value.ToString().Substring(0, $truncateLength)
            }
            elseif ($parameter.Value -is [System.Management.Automation.PSCredential])
            {
                $Message += '{0}=UserName: {1} / Password: {2},' -f $parameter.Key, $parameter.Value.UserName, $parameter.Value.GetNetworkCredential().Password
            }
            else
            {
                $Message += '{0}={1},' -f $parameter.Key, $parameter.Value
            }
        }
    }
    $Message = $Message.Substring(0, $Message.Length - 1)
    $Message += ')'
    
    $Message = '{0};{1};{2};{3}' -f (Get-Date), $callerScriptName, $callerFunctionName, $Message
    $Log.WriteEntry($Message, [System.Diagnostics.TraceEventType]::Verbose)
    $Message = ($Message -split ';')[2..3] -join ' '
    
    Microsoft.PowerShell.Utility\Write-Verbose $Message
}
#endregion

#region Write-LogFunctionExit
function Write-LogFunctionExit
{
    [CmdletBinding()]
    param
    (
        [Parameter(Position = 0)]
        [string]$ReturnValue
    )

    if ($Global:LogFunctionEntryTime)
    {
        $ts = New-TimeSpan -Start $Global:LogFunctionEntryTime -End (Get-Date)
    }
    else
    {
        $ts = New-TimeSpan -Seconds 0
    }
    
    Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    
    if (!$Log)
    {
        if ($MyInvocation.MyCommand.Module.PrivateData.AutoStart)
        {
            Start-Log -UseDefaults
        }
        else
        {
            Microsoft.PowerShell.Utility\Write-Verbose 'Cannot write to the log file until Start-Log has been called'
            return
        }
    }
    
    if (([System.Diagnostics.TraceEventType]::Verbose -band $Log.TraceSource.Switch.Level) -ne [System.Diagnostics.TraceEventType]::Verbose)
    {
        return
    }
    
    if ($ReturnValue)
    {
        $Message = "...leaving - return value is '{0}'..." -f $ReturnValue
    }
    else
    {
        $Message = '...leaving...'
    }
    
    $caller = (Get-PSCallStack)[1]
    $callerFunctionName = $caller.Command
    if ($caller.ScriptName)
    {
        $callerScriptName = Split-Path -Path $caller.ScriptName -Leaf
    }
    
    $Message = '{0};{1};{2};{3};{4}' -f (Get-Date), $callerScriptName, $callerFunctionName, $Message, ("(Time elapsed: {0:hh}:{0:mm}:{0:ss}:{0:fff})" -f $ts)
    $Log.WriteEntry($Message, [System.Diagnostics.TraceEventType]::Verbose)
    $Message = -join ($Message -split ';')[2..4]
    
    Microsoft.PowerShell.Utility\Write-Verbose $Message
}
#endregion

#region Write-LogFunctionExitWithError
function Write-LogFunctionExitWithError
{
    [CmdletBinding(
            ConfirmImpact = 'Low',
            DefaultParameterSetName = 'Message'
    )]
    
    param
    (
        [Parameter(Position = 0, ParameterSetName = 'Message')]
        [ValidateNotNullOrEmpty()]
        [string]$Message,
        
        [Parameter(Position = 0, ParameterSetName = 'ErrorRecord')]
        [ValidateNotNullOrEmpty()]
        [System.Management.Automation.ErrorRecord]$ErrorRecord,
        
        [Parameter(Position = 0, ParameterSetName = 'Exception')]
        [ValidateNotNullOrEmpty()]
        [System.Exception]$Exception,
        
        [Parameter(Position = 1)]
        [ValidateNotNullOrEmpty()]
        [string]$Details
    )

    Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
    
    if (!$Log)
    {
        if ($MyInvocation.MyCommand.Module.PrivateData.AutoStart)
        {
            Start-Log -UseDefaults
        }
        else
        {
            Microsoft.PowerShell.Utility\Write-Verbose 'Cannot write to the log file until Start-Log has been called'
            return
        }
    }
    
    if (([System.Diagnostics.TraceEventType]::Error -band $Log.TraceSource.Switch.Level) -ne [System.Diagnostics.TraceEventType]::Error)
    {
        return
    }
    
    switch ($pscmdlet.ParameterSetName)
    {
        'Message'
        {
            $Message = '...leaving: ' + $Message
        }
        'ErrorRecord'
        {
            $Message = '...leaving: ' + $ErrorRecord.Exception.Message
        }
        'Exception'
        {
            $Message = '...leaving: ' + $Exception.Message
        }
    }
    
    $EntryType = 'Error'
    
    $caller = (Get-PSCallStack)[1]
    $callerFunctionName = $caller.Command
    if ($caller.ScriptName)
    {
        $callerScriptName = Split-Path -Path $caller.ScriptName -Leaf
    }
    
    $Message = '{0};{1};{2};{3}' -f (Get-Date), $callerScriptName, $callerFunctionName, $Message
    if ($Details)
    {
        $Message += ';' + $Details
    }
    $Log.WriteEntry($Message, [System.Diagnostics.TraceEventType]::Error)
    $Message = -join ($Message -split ';')[2..3]
    
    if ($script:PSLog_Silent)
    {
        Microsoft.PowerShell.Utility\Write-Verbose -Message $Message
    }
    else
    {
        Microsoft.PowerShell.Utility\Write-Error -Message $Message
    }
}
#endregion

#region Add-Certificate2
function Add-Certificate2
{
    [cmdletBinding(DefaultParameterSetName = 'File')]
    param(
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'File')][string]$Path,
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'ByteArray')][byte[]]$Cert,
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)][System.Security.Cryptography.X509Certificates.StoreName]$Store,
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)][System.Security.Cryptography.X509Certificates.CertStoreLocation]$Location,
        [Parameter(ValueFromPipelineByPropertyName = $true)][string]$ServiceName,
        [Parameter(ValueFromPipelineByPropertyName = $true)][ValidateSet('CER', 'PFX')][string]$CertificateType = 'PFX',
        [Parameter(Mandatory = $true)][securestring]$Password
    )
    
    process
    {
        if ($Location -eq 'CERT_SYSTEM_STORE_SERVICES' -and (-not $ServiceName))
        {
            Write-Output "Please specify a ServiceName if the Location is set to 'CERT_SYSTEM_STORE_SERVICES'"
            return
        }
    
        $storePath = $Store
        
        if ($Path -and -not (Test-Path -Path $Path))
        {
            Write-Output "The path '$Path' does not exist."
            continue
        }
        
        if ($ServiceName)
        {
            if (-not (Get-Service -Name $ServiceName))
            {
                Write-Output "The service '$ServiceName' could not be found."
                return
            }
            else
            {
                $RealSvcName = (Get-Service -Name $ServiceName).Name
                $storePath = "$RealSvcName\$Store"
            }
        }
    
        $storeProvider = [System.Security.Cryptography.X509Certificates.CertStoreProvider]::CERT_STORE_PROV_SYSTEM_REGISTRY

        $Location = $Location -bor [System.Security.Cryptography.X509Certificates.CertStoreFlags]::CERT_STORE_MAXIMUM_ALLOWED_FLAG
    
        $storePtr = [System.Security.Cryptography.X509Certificates.Win32]::CertOpenStore($storeProvider, 0, 0, $Location, $storePath)
        if ($storePtr -eq [System.IntPtr]::Zero)
        {
            Write-Output "Store '$Store' in location '$Location' could not be opened."
            return
        }
    
        $s = New-Object System.Security.Cryptography.X509Certificates.X509Store($storePtr)
        $newCert = if ($Path)
        {
            if ($CertificateType -eq 'CER')
            {
                New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($Path) -ErrorAction Stop
            }
            else
            {
                New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($Path, $password, ('Exportable', 'PersistKeySet')) -ErrorAction Stop
            }
        }
        else
        {
            if ($CertificateType -eq 'CER')
            {
                New-Object System.Security.Cryptography.X509Certificates.X509Certificate2(, $Cert) -ErrorAction Stop
            }
            else
            {
                New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($Cert, $password, ('Exportable', 'PersistKeySet')) -ErrorAction Stop
            }
        }
        
        if (-not $newCert)
        {
            return
        }
    
        Write-Output "Store '$Store' in location '$Location' knows about $($s.Certificates.Count) certificates before import."
        
        $s.Add($newCert)
        
        Write-Output "Store '$Store' in location '$Location' knows about $($s.Certificates.Count) certificates after import."

        [void][System.Security.Cryptography.X509Certificates.Win32]::CertCloseStore($storePtr, 0)
    }
}
#endregion

#region Add-Certificate
function Add-Certificate
{
    [cmdletBinding(DefaultParameterSetName = 'ByteArray')]
    param(
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'File')][string]$Path,
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true, ParameterSetName = 'ByteArray')][byte[]]$Cert,
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)][System.Security.Cryptography.X509Certificates.StoreName]$Store,
        [Parameter(Mandatory = $true, ValueFromPipelineByPropertyName = $true)][System.Security.Cryptography.X509Certificates.CertStoreLocation]$Location,
        [Parameter(ValueFromPipelineByPropertyName = $true)][string]$ServiceName,
        [Parameter(ValueFromPipelineByPropertyName = $true)][ValidateSet('CER', 'PFX')][string]$CertificateType = 'PFX',
        [string]$Password = 'AL',
        [Parameter(Mandatory = $false, ValueFromPipelineByPropertyName = $true)][string[]]$ComputerName
    )
    
    begin
    {
        Write-LogFunctionEntry
    }
    
    process
    {
        Add-Certificate2 -Path $Path -Store $Store -Location $Location -ServiceName $ServiceName | Out-Null

        $fullpath = "C:\Documents and Settings\All Users\Application Data\Microsoft\Crypto\RSA\MachineKeys"
        $acl=Get-Acl -Path $fullPath
        $permission="NT AUTHORITY\NETWORK SERVICE","Read","Allow"
        $accessRule=new-object System.Security.AccessControl.FileSystemAccessRule $permission
        $acl.AddAccessRule($accessRule)
        Set-Acl $fullPath $acl
    }

    end
    {
        Write-LogFunctionExit
    }
}
#endregion Add-Certificate

Write-Output "Adding Cert Store Types"
Add-Type -TypeDefinition $certStoreTypes
