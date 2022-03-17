using System;
using System.IO;
using Microsoft.Win32;

namespace ProductDetection
{
    /// <summary>
    /// Group of methods for determining the Processor Architecture of Microsoft Office or Windows.
    /// </summary>
    public static class Architecture
    {
        #region Architecture Constants
        public const string X86 = "x86";
        public const string X64 = "x64";
        public const string WOW = "x86-64";
        public const string A64 = "ARM64";
        public const string CHPE = "x86-ARM64";
        public const string ARMX = "x64-ARM64";
        #endregion

        public static bool IsArmChk()
        {
            return File.Exists(Environment.GetEnvironmentVariable("windir") + @"\SysArm32\cmd.exe");
        }

        /// <summary>
        /// Determine whether the operating system is 32 or 64 bit
        /// </summary>
        /// <returns>String value representation of the OS Architecture</returns>
        public static string GetOSArch()
        {
            // x86 or ARM64 build < 21277
            if (Environment.Is64BitOperatingSystem == false)
            {
                if (IsArmChk() == false)
                {
                    return X86;
                }
                return A64;
            }

            // x64 or ARM64 build > 21277
            if (Environment.Is64BitOperatingSystem)
            {
                if (IsArmChk() == false)
                {
                    return X64;
                }
                return A64;
            }
            if (Environment.Is64BitProcess)
            {
                if (IsArmChk() == false)
                {
                    return X64;
                }
                return A64;
            }

            throw new ApplicationException("Unsupported Windows OS Architecture!");
        }

        /// <summary>
        /// Determine whether the installed copy of Microsoft Office is 32 bit, 64 bit, or running under WOW64 emulation
        /// </summary>
        /// <returns>String value representation of the Office Architecture</returns>
        public static string GetOfficeArch()
        {
            // x86 or ARM64 build < 21277
            if (Environment.Is64BitOperatingSystem == false)
            {
                if (IsArmChk() == false)
                {
                    return X86;
                }
                return CHPE;
            }

            // Use Registry Detection
            if (!OfficeVersion.IsOfficeVirtual())
            {
                // Traditional Check
                string installRoot;
                int officeNumber = OfficeVersion.GetOfficeNumber();
                switch (officeNumber)
                {
                    // Office 2003 and Office 2007 Don't have 64 Bit Edition, so they must be WOW
                    case 11:
                    case 12:
                        return WOW;
                    case 14:
                    case 15:
                    case 16:
                        installRoot = "SOFTWARE\\Microsoft\\Office\\" + officeNumber + ".0\\Common\\InstallRoot";
                        break;
                    default:
                        throw new ApplicationException("Unsupported Microsoft Office Edition!");
                }
                // x86-64 Check
                using (RegistryKey officeArchCheck = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32).OpenSubKey(installRoot, false))
                {
                    if (officeArchCheck != null)
                    {
                        string installPath = (string)officeArchCheck.GetValue("Path");
                        // If a path is stored here, Office is running under WOW64
                        if (String.IsNullOrWhiteSpace(installPath) == false)
                        {
                            return WOW;
                        }
                    }
                }
                // x64 Check
                using (RegistryKey officeArchCheck = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey(installRoot, false))
                {
                    if (officeArchCheck != null)
                    {
                        string installPath = (string)officeArchCheck.GetValue("Path");
                        // If a path is stored here, Office is running under x64
                        if (String.IsNullOrWhiteSpace(installPath) == false)
                        {
                            return X64;
                        }
                    }
                }
            }
            else
            {
                // Office 16 C2R Check
                using (RegistryKey officeArchCheck = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32).OpenSubKey(@"SOFTWARE\Microsoft\Office\ClickToRun\Configuration", false))
                {
                    if (officeArchCheck != null)
                    {
                        string platform = (string)officeArchCheck.GetValue("Platform");
                        if (platform.ToLower() == "x86")
                        {
                            if (IsArmChk() == false)
                            {
                                return WOW;
                            }
                            return CHPE;
                        }
                        if (platform.ToLower() == "x64")
                        {
                            return ARMX;
                        }
                    }
                }
                using (RegistryKey officeArchCheck = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey(@"SOFTWARE\Microsoft\Office\ClickToRun\Configuration", false))
                {
                    if (officeArchCheck != null)
                    {
                        string platform = (string)officeArchCheck.GetValue("Platform");
                        if (platform.ToLower() == "x86")
                        {
                            if (IsArmChk() == false)
                            {
                                return WOW;
                            }
                            return CHPE;
                        }
                        if (platform.ToLower() == "x64")
                        {
                            if (IsArmChk() == false)
                            {
                                return X64;
                            }
                            return ARMX;
                        }
                    }
                }
                // Office 15 C2R Check
                using (RegistryKey officeArchCheck = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32).OpenSubKey(@"SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration", false))
                {
                    if (officeArchCheck != null)
                    {
                        string platform = (string)officeArchCheck.GetValue("Platform");
                        if (platform.ToLower() == "x86")
                        {
                            return WOW;
                        }
                        return X64;
                    }
                }
                using (RegistryKey officeArchCheck = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey(@"SOFTWARE\Microsoft\Office\15.0\ClickToRun\Configuration", false))
                {
                    if (officeArchCheck != null)
                    {
                        string platform = (string)officeArchCheck.GetValue("Platform");
                        if (platform.ToLower() == "x86")
                        {
                            return WOW;
                        }
                        return X64;
                    }
                }
            }
            throw new ApplicationException("Unsupported Microsoft Office Architecture!");
        }
    }
}