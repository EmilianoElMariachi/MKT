using System;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Management;
using System.ServiceProcess;
using Microsoft.Win32;

namespace ProductDetection
{
    /// <summary>
    /// Generic class for obtaining the Software Licensing Service version
    /// </summary>
    internal static class SLVersion
    {
        /// <summary>
        /// Query WMI to determine the Software Licensing Service version
        /// </summary>
        /// <param name="wmiInfo">WMI Provider and associated data to get the Software Licensing Service version</param>
        /// <returns>Version from SPPWMI Provider</returns>
        internal static string GetSLVersion(string wmiInfo)
        {
            using (ManagementObjectSearcher searcher = new ManagementObjectSearcher(@"root\CIMV2", wmiInfo))
            {
                foreach (ManagementBaseObject queryObj in searcher.Get())
                {
                    return Convert.ToString(queryObj["Version"]);
                }
            }

            throw new Exception("Failed to get SL Version!");
        }
    }

    /// <summary>
    /// Generic class for obtaining OS branding strings
    /// </summary>
    internal static class OSBranding
    {
        /// <summary>
        /// Query winbrand.dll to obtain branding string
        /// </summary>
        /// <param name="ids">IDS branding string to return</param>
        /// <returns>replaces the parameter with the corresponding product string, and returns the new string</returns>
        internal static string GetOSBranding(string ids)
        {
            try
            {
                IntPtr dit = NativeMethods.BrandingFormatString(ids);
                string brd = Marshal.PtrToStringUni(dit);
                Marshal.FreeHGlobal(dit);
                if (brd.Equals(ids, StringComparison.OrdinalIgnoreCase) == false)
                {
                    return brd;
                }
                else
                {
                    return "Unknown";
                }
            }
            catch
            {
                return "Unknown";
            }
        }
    }

    /// <summary>
    /// Group of methods for determining Operating System version and Product Name
    /// </summary>
    public static class OSVersion
    {
        #region Product Name Constants
        public const string Win95 = "Windows 95";
        public const string Win95SE = "Windows 95 SE";
        public const string Win98 = "Windows 98";
        public const string WinME = "Windows ME";
        public const string WinNT351 = "Windows NT 3.51";
        public const string WinNT4 = "Windows NT 4.0";
        public const string Win2000 = "Windows 2000";
        public const string WinXP = "Windows XP";

        public const string WinVista = "Windows Vista";
        public const string WinServer2008 = "Windows Server 2008";

        public const string Win7 = "Windows 7";
        public const string Win7Embedded = "Windows 7 Embedded";
        public const string WinServer2008R2 = "Windows Server 2008 R2";

        public const string Win8 = "Windows 8";
        public const string Win8Embedded = "Windows 8 Embedded";
        public const string WinServer2012 = "Windows Server 2012";

        public const string Win81 = "Windows 8.1";
        public const string Win81Embedded = "Windows 8.1 Embedded";
        public const string WinServer2012R2 = "Windows Server 2012 R2";

        public const string WinTP = "Windows 10 Technical Preview";
        public const string WinServerTP = "Windows Server Technical Preview";

        public const string Win10 = "Windows 10";
        public const string Win10Embedded = "Windows 10 IoT";
        public const string WinServer2016 = "Windows Server 2016";
        public const string WinServer2019 = "Windows Server 2019";
        public const string WinServer2022 = "Windows Server 2022";

        public const string Win11 = "Windows 11";
        #endregion
        /// <summary>
        /// ServerRdsh is a ServerNT but act as WinNT in activation (KMS N count 25)
        /// </summary>
        /// <returns>True if Windows is a ServerRdsh edition, False if otherwise</returns>
        public static bool IsServerRdsh()
        {
            if (GetWindowsBuildNumber() < 17063)
            {
                return false;
            }

            int iSkuNum;
            string sSkuNum = string.Empty;
            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher(@"root\CIMV2", "SELECT * FROM Win32_OperatingSystem");
                foreach (ManagementObject queryObj in searcher.Get())
                {
                    sSkuNum = queryObj["OperatingSystemSKU"].ToString();
                }
                iSkuNum = Convert.ToInt32(sSkuNum);
            }
            catch
            {
                return false;
            }

            return iSkuNum == 175;
        }

        /// <summary>
        /// Determine if the installed version of Windows is a server version
        /// </summary>
        /// <returns>True if Windows is a Server Version, False if Windows is a Client Version</returns>
        public static bool IsWindowsServer()
        {
            if (IsServerRdsh())
            {
                return false;
            }
            else
            {
                return NativeMethods.IsOS(OSType.AnyServer);
            }
        }

        /// <summary>
        /// Call to GetSLVersion() to get the Windows Software Licensing Service Version
        /// </summary>
        /// <returns>Version of Windows Software Licensing Service</returns>
        public static string GetSPPSVCVersion()
        {
            return SLVersion.GetSLVersion("SELECT Version FROM SoftwareLicensingService");
        }

        /// <summary>
        /// Get the Product Name of the installed copy of Windows
        /// </summary>
        /// <returns>Microsoft Windows Product Name</returns>
        public static string GetWindowsName()
        {
            // Get the Kernel32 DLL File Version
            FileVersionInfo osVersionInfo = FileVersionInfo.GetVersionInfo(Environment.GetEnvironmentVariable("windir") + @"\System32\Kernel32.dll");
            
            // Variable to hold our return value
            string operatingSystem = string.Empty;

            switch (osVersionInfo.FileMajorPart)
            {
                case 3:
                    operatingSystem = WinNT351;
                    break;
                case 4:
                    operatingSystem = WinNT4;
                    break;
                case 5:
                    if (osVersionInfo.FileMinorPart == 0)
                    {
                        operatingSystem = Win2000;
                    }
                    else
                    {
                        operatingSystem = WinXP;
                    }
                    break;
                case 6:
                    switch (osVersionInfo.FileMinorPart)
                    {
                        case 0:
                            if (IsWindowsServer() == false)
                            {
                                operatingSystem = WinVista;
                            }
                            else
                            {
                                operatingSystem = WinServer2008;
                            }
                            break;
                        case 1:
                            if (IsWindowsServer() == false)
                            {
                                operatingSystem = Win7;
                            }
                            else
                            {
                                operatingSystem = WinServer2008R2;
                            }
                            break;
                        case 2:
                            if (IsWindowsServer() == false)
                            {
                                operatingSystem = Win8;
                            }
                            else
                            {
                                operatingSystem = WinServer2012;
                            }
                            break;
                        case 3:
                            if (IsWindowsServer() == false)
                            {
                                operatingSystem = Win81;
                            }
                            else
                            {
                                operatingSystem = WinServer2012R2;
                            }
                            break;
                        case 4:
                            if (IsWindowsServer() == false)
                            {
                                operatingSystem = WinTP;
                            }
                            else
                            {
                                operatingSystem = WinServerTP;
                            }
                            break;
                    }
                    break;
                case 10:
                    switch (osVersionInfo.FileMinorPart)
                    {
                        case 0:
                            if (IsWindowsServer() == false)
                            {
                                operatingSystem = Win10;
                            }
                            else
                            {
                                if (osVersionInfo.FileBuildPart >= 20251)
                                {
                                    operatingSystem = WinServer2022;
                                }
                                else if (osVersionInfo.FileBuildPart >= 17763)
                                {
                                    operatingSystem = WinServer2019;
                                }
                                else
                                {
                                    operatingSystem = WinServer2016;
                                }
                            }
                            break;
                    }
                    break;
                case 11:
                    switch (osVersionInfo.FileMinorPart)
                    {
                        case 0:
                            if (IsWindowsServer() == false)
                            {
                                operatingSystem = Win11;
                            }
                            break;
                    }
                    break;
                default:
                    operatingSystem = "Unknown";
                    break;
            }

            if (operatingSystem.Equals(Win10, StringComparison.OrdinalIgnoreCase))
            {
                string brandos = OSBranding.GetOSBranding("%WINDOWS_LONG%");
                if (brandos.IndexOf("Windows 11", StringComparison.OrdinalIgnoreCase) > 0)
                {
                    operatingSystem = Win11;
                }
            }

            // Return the information we've gathered.
            return operatingSystem;
        }

        /// <summary>
        /// Get the Build number of the installed copy of Windows
        /// </summary>
        /// <returns>Microsoft Windows Build number as integer</returns>
        public static int GetWindowsBuildNumber()
        {
            // Get the Kernel32 DLL File Version
            FileVersionInfo osVersionInfo = FileVersionInfo.GetVersionInfo(Environment.GetEnvironmentVariable("windir") + @"\System32\Kernel32.dll");

            int iBldNum = osVersionInfo.FileBuildPart;
            if (iBldNum >= 18362)
            {
                string sBldNum = string.Empty;
                try
                {
                    ManagementObjectSearcher searcher = new ManagementObjectSearcher(@"root\CIMV2", "SELECT * FROM Win32_OperatingSystem");
                    foreach (ManagementObject queryObj in searcher.Get())
                    {
                        sBldNum = queryObj["BuildNumber"].ToString();
                    }
                    if (String.IsNullOrWhiteSpace(sBldNum) == false)
                        iBldNum = Convert.ToInt32(sBldNum);
                }
                catch
                {
                    using (RegistryKey registry = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion", false))
                    {
                        if (registry.GetValue("CurrentBuild") != null)
                        {
                            sBldNum = (string)registry.GetValue("CurrentBuild");
                            iBldNum = Convert.ToInt32(sBldNum);
                        }
                    }
                }
            }

            return iBldNum;
        }

        /// <summary>
        /// Get the Platform Version of the installed copy of Windows
        /// </summary>
        /// <returns>Microsoft Windows Platform version as double</returns>
        public static double GetWindowsNumber()
        {
            // Get the Kernel32 DLL File Version
            FileVersionInfo osVersionInfo = FileVersionInfo.GetVersionInfo(Environment.GetEnvironmentVariable("windir") + @"\System32\Kernel32.dll");

            return double.Parse(osVersionInfo.FileMajorPart + "." + osVersionInfo.FileMinorPart, CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// Determine if the installed copy of Windows is supported based on its platform version number
        /// </summary>
        /// <returns>True if supported, False if unsupported</returns>
        public static bool IsWindowsSupported()
        {
            double windowsNumber = GetWindowsNumber();
            if ((windowsNumber >= 6.0 && windowsNumber <= 11.0))
            {
                return true;
            }
            return false;
        }
    }

    /// <summary>
    /// Group of methods for determining Microsoft Office version
    /// </summary>
    public static class OfficeVersion
    {
        #region Product Name Constants
        public const string Office2003 = "Microsoft Office 2003";
        public const string Office2007 = "Microsoft Office 2007";
        public const string Office2010 = "Microsoft Office 2010";
        public const string Office2013 = "Microsoft Office 2013";
        public const string OffC2R2013 = "Microsoft Office 2013 C2R";
        public const string OffC2R2016 = "Microsoft Office 2016 C2R";
        public const string Office2016 = "Microsoft Office 2016";
        public const string Office2019 = "Microsoft Office 2019";
        public const string Office2021 = "Microsoft Office 2021";
        #endregion
        /// <summary>
        /// Call to GetSLVersion() to get the Office Software Licensing Service Version
        /// </summary>
        /// <returns>Version of Office Software Licensing Service</returns>
        public static string GetOSPPSVCVersion()
        {
            return SLVersion.GetSLVersion("SELECT Version FROM OfficeSoftwareProtectionService");
        }

        /// <summary>
        /// Get the Product Name of the installed copy of Office
        /// </summary>
        /// <returns>Microsoft Office Product Name</returns>
        public static string GetOfficeName()
        {
            // Use Registry Detection for Traditional Installed Office
            if (!IsOfficeVirtual())
            {
                using (RegistryKey officeKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey(@"SOFTWARE\Microsoft\Office", false))
                {
                    using (RegistryKey officeKeyWOW = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32).OpenSubKey(@"SOFTWARE\Microsoft\Office", false))
                    {
                        // Check Office 2016
                        if (officeKeyWOW != null && (officeKey != null && IsOfficeInstalledTraditional(officeKey.OpenSubKey(@"16.0\Common\InstallRoot"), officeKeyWOW.OpenSubKey(@"16.0\Common\InstallRoot"))))
                        {
                            return Office2016;
                        }
                        // Check Office 2013
                        if (officeKeyWOW != null && (officeKey != null && IsOfficeInstalledTraditional(officeKey.OpenSubKey(@"15.0\Common\InstallRoot"), officeKeyWOW.OpenSubKey(@"15.0\Common\InstallRoot"))))
                        {
                            return Office2013;
                        }
                        // Check Office 2010
                        if (officeKeyWOW != null && (officeKey != null && IsOfficeInstalledTraditional(officeKey.OpenSubKey(@"14.0\Common\InstallRoot"), officeKeyWOW.OpenSubKey(@"14.0\Common\InstallRoot"))))
                        {
                            return Office2010;
                        }
                        // Check Office 2007
                        if (officeKeyWOW != null && (officeKey != null && IsOfficeInstalledTraditional(officeKey.OpenSubKey(@"12.0\Common\InstallRoot"), officeKeyWOW.OpenSubKey(@"12.0\Common\InstallRoot"))))
                        {
                            return Office2007;
                        }
                        // Check Office 2003
                        if (officeKeyWOW != null && (officeKey != null && IsOfficeInstalledTraditional(officeKey.OpenSubKey(@"11.0\Common\InstallRoot"), officeKeyWOW.OpenSubKey(@"11.0\Common\InstallRoot"))))
                        {
                            return Office2003;
                        }
                    }
                }
            }
            // Use Registry Detection for Virtually Installed Office
            else
            {
                using (RegistryKey officeKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey(@"SOFTWARE\Microsoft\Office", false))
                {
                    if (officeKey != null)
                    {
                        using (RegistryKey officeKeyC2R = officeKey.OpenSubKey(@"ClickToRun\Configuration", false))
                        {
                            if (officeKeyC2R != null)
                            {
                                string productids = (string)officeKeyC2R.GetValue("ProductReleaseIds");

                                if (String.IsNullOrWhiteSpace(productids) == false)
                                {
                                    if (productids.Contains("2021"))
                                    {
                                        return Office2021;
                                    }
                                    else if (productids.Contains("2019"))
                                    {
                                        return Office2019;
                                    }
                                    else
                                    {
                                        return OffC2R2016;
                                    }
                                }
                            }
                        }

                        using (RegistryKey officeKeyC2R = officeKey.OpenSubKey(@"15.0\ClickToRun\ProductReleaseIDs\Active\culture", false))
                        {
                            if (officeKeyC2R != null)
                            {
                                string xnone = (string)officeKeyC2R.GetValue("x-none");

                                if (String.IsNullOrWhiteSpace(xnone) == false)
                                {
                                    return OffC2R2013;
                                }
                            }
                        }
                    }
                }

                using (RegistryKey officeKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32).OpenSubKey(@"SOFTWARE\Microsoft\Office", false))
                {
                    if (officeKey != null)
                    {
                        using (RegistryKey officeKeyC2R = officeKey.OpenSubKey(@"ClickToRun\Configuration", false))
                        {
                            if (officeKeyC2R != null)
                            {
                                string productids = (string)officeKeyC2R.GetValue("ProductReleaseIds");

                                if (String.IsNullOrWhiteSpace(productids) == false)
                                {
                                    if (productids.Contains("2021"))
                                    {
                                        return Office2021;
                                    }
                                    else if (productids.Contains("2019"))
                                    {
                                        return Office2019;
                                    }
                                    else
                                    {
                                        return OffC2R2016;
                                    }
                                }
                            }
                        }

                        using (RegistryKey officeKeyC2R = officeKey.OpenSubKey(@"15.0\ClickToRun\ProductReleaseIDs\Active\culture", false))
                        {
                            if (officeKeyC2R != null)
                            {
                                string xnone = (string)officeKeyC2R.GetValue("x-none");

                                if (String.IsNullOrWhiteSpace(xnone) == false)
                                {
                                    return OffC2R2013;
                                }
                            }
                        }
                    }
                }
            }

            throw new ApplicationException("Unsupported Microsoft Office Edition!");
        }

        /// <summary>
        /// Helper Method for GetOfficeName()
        /// </summary>
        /// <param name="officeInstallCheck">Registry Key to Check</param>
        /// <param name="officeInstallCheckWOW">WOW Redirected Version of officeInstallCheck</param>
        /// <returns>True if Installed, False if Not Installed</returns>
        private static bool IsOfficeInstalledTraditional(RegistryKey officeInstallCheck, RegistryKey officeInstallCheckWOW)
        {
            // Check if the Passed Keys Contain a Path Value that isn't null
            if (officeInstallCheck != null)
            {
                string installPath = (string)officeInstallCheck.GetValue("Path");
                // If a path is stored here, Office is installed
                if (String.IsNullOrWhiteSpace(installPath) == false)
                {
                    return true;
                }
            }
            else if (officeInstallCheckWOW != null)
            {
                string installPath = (string)officeInstallCheckWOW.GetValue("Path");
                // If a path is stored here, Office is installed
                if (String.IsNullOrWhiteSpace(installPath) == false)
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Get the installation of Microsoft Office from the registry
        /// </summary>
        /// <returns>String Representation of the Path to Microsoft Office Installation</returns>
        public static string GetInstallationPath()
        {      
            int officeNumber = GetOfficeNumber();
            // Traditional Check
            if (!IsOfficeVirtual())
            {
                switch (Architecture.GetOfficeArch())
                {
                    case Architecture.X86:
                    case Architecture.X64:
                        {
                            using (RegistryKey installRoot = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey("SOFTWARE\\Microsoft\\Office\\" + officeNumber + ".0\\Common\\InstallRoot", false))
                            {
                                if (installRoot != null)
                                {
                                    return (string)installRoot.GetValue("Path");
                                }
                            }
                            break;
                        }
                    case Architecture.WOW:
                        {
                            using (RegistryKey installRoot = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32).OpenSubKey("SOFTWARE\\Microsoft\\Office\\" + officeNumber + ".0\\Common\\InstallRoot", false))
                            {
                                if (installRoot != null)
                                {
                                    return (string)installRoot.GetValue("Path");
                                }
                            }
                            break;
                        }
                }
            }
            // Virtual Check
            else
            {
                using (RegistryKey installRoot = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey("SOFTWARE\\Microsoft\\Office\\ClickToRun", false))
                {
                    if (installRoot != null)
                    {
                        return (string)installRoot.GetValue("InstallPath") + "\\Office16" + Path.DirectorySeparatorChar;
                    }
                }
                using (RegistryKey installRoot = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32).OpenSubKey("SOFTWARE\\Microsoft\\Office\\ClickToRun", false))
                {
                    if (installRoot != null)
                    {
                        return (string)installRoot.GetValue("InstallPath") + "\\Office16" + Path.DirectorySeparatorChar;
                    }
                }
                using (RegistryKey installRoot = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey("SOFTWARE\\Microsoft\\Office\\15.0\\ClickToRun", false))
                {
                    if (installRoot != null)
                    {
                        switch (Architecture.GetOfficeArch())
                        {
                            case Architecture.X86:
                            case Architecture.X64:
                                {
                                    return (string)Environment.GetEnvironmentVariable("SystemDrive") + "\\Program Files\\Microsoft Office\\Office15" + Path.DirectorySeparatorChar;
                                }
                            case Architecture.WOW:
                                {
                                    return (string)Environment.GetEnvironmentVariable("SystemDrive") + "\\Program Files (x86)\\Microsoft Office\\Office15" + Path.DirectorySeparatorChar;
                                }
                        }
                    }
                }
            }

            throw new Exception("Failed to get Microsoft Office Installation Path!");
        }

        /// <summary>
        /// Get the Platform Version of the installed copy of Office
        /// </summary>
        /// <returns>Microsoft Office Platform version number</returns>
        public static int GetOfficeNumber()
        {
            switch (GetOfficeName())
            {
                case Office2003:
                    return 11;
                case Office2007:
                    return 12;
                case Office2010:
                    return 14;
                case Office2013:
                case OffC2R2013:
                    return 15;
                case OffC2R2016:
                case Office2016:
                case Office2019:
                case Office2021:
                    return 16;
                default:
                    throw new ApplicationException("Unsupported Microsoft Office Edition!");
            }
        }

        /// <summary>
        /// Determine if the installed copy of Office is supported based on its platform version number
        /// </summary>
        /// <returns>True if supported, False if unsupported</returns>
        public static bool IsOfficeSupported()
        {
            try
            {
                int officeNumber = GetOfficeNumber();
                if (officeNumber >= 14 && officeNumber <= 16)
                {
                    return true;
                }
                return false;
            }
            catch (ApplicationException)
            {
                return false;
            }
        }

        /// <summary>
        /// Check if Microsoft Office is using SPPSVC from the OS instead of bundled OSPPSVC.
        /// </summary>
        /// <returns>True if using SPPSVC, False if using OSPPSVC</returns>
        public static bool IsOfficeSPP()
        {
            try
            {
                return (GetOfficeNumber() >= 15 && OSVersion.GetWindowsNumber() >= 6.2);
            }
            catch (ApplicationException)
            {
                return false;
            }
        }

        /// <summary>
        /// Check if Microsoft Office 16 is using Click To Run or is a Traditional install.
        /// </summary>
        /// <returns>True if using Click To Run, False if using Traditional</returns>
        public static bool IsOfficeVirtual()
        {
            if (ChkService("ClickToRunSvc") == false && ChkService("OfficeSvc") == false)
            {
                return false;
            }
            using (RegistryKey officeKeyC2R = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey("SOFTWARE\\Microsoft\\Office\\ClickToRun", false))
            {
                if (officeKeyC2R != null)
                {
                    try
                    {
                        string packageGUID = officeKeyC2R.GetValue("PackageGUID").ToString();
                        return (!String.IsNullOrWhiteSpace(packageGUID));
                    }
                    catch (NullReferenceException)
                    {
                        return false;
                    }                    
                }
            }
            using (RegistryKey officeKeyC2R = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32).OpenSubKey("SOFTWARE\\Microsoft\\Office\\ClickToRun", false))
            {
                if (officeKeyC2R != null)
                {
                    try
                    {
                        string packageGUID = officeKeyC2R.GetValue("PackageGUID").ToString();
                        return (!String.IsNullOrWhiteSpace(packageGUID));
                    }
                    catch (NullReferenceException)
                    {
                        return false;
                    }                    
                }
            }
            using (RegistryKey officeKeyC2R = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey("SOFTWARE\\Microsoft\\Office\\15.0\\ClickToRun", false))
            {
                if (officeKeyC2R != null)
                {
                    string packageGUID = officeKeyC2R.GetValue("PackageGUID").ToString();
                    if (String.IsNullOrWhiteSpace(packageGUID) == false)
                    {
                        return true;
                    }
                    return false;
                }
            }
            using (RegistryKey officeKeyC2R = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32).OpenSubKey("SOFTWARE\\Microsoft\\Office\\15.0\\ClickToRun", false))
            {
                if (officeKeyC2R != null)
                {
                    string packageGUID = officeKeyC2R.GetValue("PackageGUID").ToString();
                    if (String.IsNullOrWhiteSpace(packageGUID) == false)
                    {
                        return true;
                    }
                    return false;
                }
            }
            return false;
        }

        /// <summary>
        /// Check is a Service is Registered
        /// </summary>
        /// <param name="serviceName">Name of the Service to search for</param>
        /// <returns>True if Registered, False if not Registered</returns>
        public static bool ChkService(string serviceName)
        {
            foreach (ServiceController controller in ServiceController.GetServices())
            {
                if (controller.ServiceName == serviceName)
                {
                    return true;
                }
            }
            return false;
        }
    }
}