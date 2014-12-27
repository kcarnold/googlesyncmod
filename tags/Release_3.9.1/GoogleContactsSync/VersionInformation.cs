using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Management;
using System.Management.Instrumentation;

namespace GoContactSyncMod
{
    static class VersionInformation
    {
        public enum OutlookMainVersion
        {
            Outlook2002,
            Outlook2003,
            Outlook2007,
            Outlook2010,
            Outlook2013,
            OutlookUnknownVersion,
            OutlookNoInstance
        }

        public static OutlookMainVersion GetOutlookVersion(Microsoft.Office.Interop.Outlook.Application appVersion)
        {
            if (appVersion == null)
                appVersion = new Microsoft.Office.Interop.Outlook.Application();

            switch (appVersion.Version.ToString().Substring(0,2))
            {
                case "10":
                    return OutlookMainVersion.Outlook2002;
                case "11":
                    return OutlookMainVersion.Outlook2003;
                case "12":
                    return OutlookMainVersion.Outlook2007;
                case "14":
                    return OutlookMainVersion.Outlook2010;
                case "15":
                    return OutlookMainVersion.Outlook2013;
                default:
                    {
                        if (appVersion != null)
                        {
                            Marshal.ReleaseComObject(appVersion);
                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                        }
                        return OutlookMainVersion.OutlookUnknownVersion;
                    }
            }
     
        }

        /// <summary>
        /// detect windows main version
        /// </summary>
        public static string GetWindowsVersionName()
        {
            using (ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_OperatingSystem"))
            {
                foreach (ManagementObject managementObject in searcher.Get())
                {
                    /*//iterate trough all properties
                    foreach (PropertyData prop in managementObject.Properties)
                    {
                        Console.WriteLine("{0}: {1}", prop.Name, prop.Value);
                    }
                     */
                    return (string)managementObject["Caption"];
                }
            }
            return "Unknown Windows Version";
        }
    }
}
