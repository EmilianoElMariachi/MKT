using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using Common;
using OfficeUninstaller.Properties;
using ProductDetection;
using SharpCompress;
using SharpCompress.Archives;
using SharpCompress.Common;
using SharpCompress.Readers;

namespace OfficeUninstaller
{
    /// <summary>
    /// Group of Methods for Uninstalling Microsoft Office
    /// </summary>
    public static class OfficeUninstaller
    {
        /// <summary>
        /// Names of Scripts used to uninstall a specific Microsoft Office Product
        /// </summary>
        private const string Script2003FileName = "OffScrub03.vbs";
        private const string Script2007FileName = "OffScrub07.vbs";
        private const string Script2010FileName = "OffScrub10.vbs";
        private const string Script2013MSIFileName = "OffScrub15msi.vbs";
        private const string Script2016MSIFileName = "OffScrub16msi.vbs";
        private const string ScriptC2RFileName = "OffScrubC2R.vbs";

        /// <summary>
        /// Run a script to Uninstall Microsoft Office.
        /// </summary>
        /// <param name="productToUninstall">Microsoft Office Product to Uninstall</param>
        /// <returns>Results from Uninstaller Script</returns>
        public static string Uninstaller(string productToUninstall)
        {
            // Create Base Path
            string path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
            Directory.CreateDirectory(path);

            // Extract OfficeScrub.7z
            byte[] OfficeScrub = Resources.OffScrub;
            using (IArchive aScrub = ArchiveFactory.Open(new MemoryStream(OfficeScrub)))
            {
                IReader reader = aScrub.ExtractAllEntries();
                while (reader.MoveToNextEntry())
                {
                    if (!reader.Entry.IsDirectory)
                    {
                        reader.WriteEntryToDirectory(CommonUtilities.EscapePath(path), new ExtractionOptions() { ExtractFullPath = true, Overwrite = true });
                    }
                }
            }

            using (StringWriter output = new StringWriter())
            {
                if (productToUninstall == OfficeVersion.Office2003)
                {
                    // Name of MSI Uninstaller
                    string saveAsName = CommonUtilities.EscapePath(path + Path.DirectorySeparatorChar + Script2003FileName);

                    // Run Uninstaller Script
                    output.WriteLine("----------------------------------------");
                    output.WriteLine("Running Microsoft Office 2003 MSI Uninstall Script");
                    output.WriteLine("----------------------------------------");
                    Result result = CommonUtilities.ExecuteCommand("cscript " + saveAsName + " All /DELETEUSERSETTINGS /FORCE /NOCANCEL /OSE", true);

                    // Return Output
                    if (result.HasError)
                    {
                        output.WriteLine(result.Error);
                    }
                    else
                    {
                        output.WriteLine(result.Output);
                    }
                    output.Write("----------------------------------------");
                }
                else if (productToUninstall == OfficeVersion.Office2007)
                {
                    // Name of MSI Uninstaller
                    string saveAsName = CommonUtilities.EscapePath(path + Path.DirectorySeparatorChar + Script2007FileName);

                    // Run Uninstaller Script
                    output.WriteLine("----------------------------------------");
                    output.WriteLine("Running Microsoft Office 2007 MSI Uninstall Script");
                    output.WriteLine("----------------------------------------");
                    Result result = CommonUtilities.ExecuteCommand("cscript " + saveAsName + " All /DELETEUSERSETTINGS /FORCE /NOCANCEL /OSE", true);

                    // Return Output
                    if (result.HasError)
                    {
                        output.WriteLine(result.Error);
                    }
                    else
                    {
                        output.WriteLine(result.Output);
                    }
                    output.Write("----------------------------------------");
                }
                else if (productToUninstall == OfficeVersion.Office2010)
                {
                    // Name of MSI Uninstaller
                    string saveAsName = CommonUtilities.EscapePath(path + Path.DirectorySeparatorChar + Script2010FileName);

                    // Run Uninstaller Script
                    output.WriteLine("----------------------------------------");
                    output.WriteLine("Running Microsoft Office 2010 MSI Uninstall Script");
                    output.WriteLine("----------------------------------------");
                    Result result = CommonUtilities.ExecuteCommand("cscript " + saveAsName + " All /DELETEUSERSETTINGS /FORCE /NOCANCEL /OSE", true);

                    // Return Output
                    if (result.HasError)
                    {
                        output.WriteLine(result.Error);
                    }
                    else
                    {
                        output.WriteLine(result.Output);
                    }
                    output.Write("----------------------------------------");
                }
                else if (productToUninstall == OfficeVersion.Office2013 || productToUninstall == OfficeVersion.OffC2R2013)
                {
                    // Name of MSI Uninstaller
                    string saveAsNameMSI = CommonUtilities.EscapePath(path + Path.DirectorySeparatorChar + Script2013MSIFileName);

                    // Run Traditional Uninstaller Script
                    output.WriteLine("----------------------------------------");
                    output.WriteLine("Running Microsoft Office 2013 MSI Uninstall Script");
                    output.WriteLine("----------------------------------------");
                    Result msiResult = CommonUtilities.ExecuteCommand("cscript " + saveAsNameMSI + " All /DELETEUSERSETTINGS /FORCE /NOCANCEL /NOREBOOT /OSE /REMOVELYNC /REMOVEOSPP", true);
                    if (msiResult.HasError)
                    {
                        output.WriteLine(msiResult.Error);
                    }
                    else
                    {
                        output.WriteLine(msiResult.Output);
                    }

                    // Name of Virtual Uninstaller
                    string saveAsNameC2R = CommonUtilities.EscapePath(path + Path.DirectorySeparatorChar + ScriptC2RFileName);

                    // Run Virtual Uninstaller Script
                    output.WriteLine("----------------------------------------");
                    output.WriteLine("Running Microsoft Office 2013 Click To Run Uninstall Script");
                    output.WriteLine("----------------------------------------");
                    Result virtualResult = CommonUtilities.ExecuteCommand("cscript " + saveAsNameC2R + " ALL /NOCANCEL /OSE", true);
                    if (virtualResult.HasError)
                    {
                        output.WriteLine(virtualResult.Error);
                    }
                    else
                    {
                        output.WriteLine(virtualResult.Output);
                    }
                    output.Write("----------------------------------------");
                }
                else if (productToUninstall == OfficeVersion.Office2016 || productToUninstall == OfficeVersion.OffC2R2016 || productToUninstall == OfficeVersion.Office2019 || productToUninstall == OfficeVersion.Office2021)
                {
                    // Name of MSI Uninstaller
                    string saveAsNameMSI = CommonUtilities.EscapePath(path + Path.DirectorySeparatorChar + Script2016MSIFileName);

                    // Run Traditional Uninstaller Script
                    output.WriteLine("----------------------------------------");
                    output.WriteLine("Running Microsoft Office 2016 MSI Uninstall Script");
                    output.WriteLine("----------------------------------------");
                    Result msiResult = CommonUtilities.ExecuteCommand("cscript " + saveAsNameMSI + " All /DELETEUSERSETTINGS /FORCE /NOCANCEL /NOREBOOT /OSE /REMOVELYNC /REMOVEOSPP", true);
                    if (msiResult.HasError)
                    {
                        output.WriteLine(msiResult.Error);
                    }
                    else
                    {
                        output.WriteLine(msiResult.Output);
                    }

                    // Name of Virtual Uninstaller
                    string saveAsNameC2R = CommonUtilities.EscapePath(path + Path.DirectorySeparatorChar + ScriptC2RFileName);

                    // Run Virtual Uninstaller Script
                    output.WriteLine("----------------------------------------");
                    output.WriteLine("Running Microsoft Office 2016-2021 Click To Run Uninstall Script");
                    output.WriteLine("----------------------------------------");
                    Result virtualResult = CommonUtilities.ExecuteCommand("cscript " + saveAsNameC2R + " ALL /NOCANCEL /OSE", true);
                    if (virtualResult.HasError)
                    {
                        output.WriteLine(virtualResult.Error);
                    }
                    else
                    {
                        output.WriteLine(virtualResult.Output);
                    }
                    output.Write("----------------------------------------");
                }
                else
                {
                    throw new Exception("Could not find an uninstaller script for this Microsoft Office Edition!");
                }

                // Delete Temporary Folder and Return Output
                CommonUtilities.FolderDelete(path);
                return output.ToString();
            }
        }
    }
}