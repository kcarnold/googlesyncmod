﻿using System;
using System.Windows.Forms;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Globalization;

namespace GoContactSyncMod
{
    static class ErrorHandler
    {

        private static string OSInfo
        {
            get
            {
                return VersionInformation.GetWindowsMainVersion().ToString();
            }
        }

        private static string OutlookInfo
        {
            get
            {
                return VersionInformation.GetOutlookVersion(Syncronizer.OutlookApplication).ToString();
            }
        }

        // TODO: Write a nice error dialog, that maybe supports directly E-Mail sending as bugreport
        public static void Handle(Exception ex)
        {
            Logger.Log(ex.Message, EventType.Error);
            //AppendSyncConsoleText(Logger.GetText());
            Logger.Log("Sync failed.", EventType.Error);

            try
            {
                SettingsForm.Instance.ShowBalloonToolTip("Error", ex.Message, ToolTipIcon.Error, 5000, true);
                /*
				Program.Instance.notifyIcon.BalloonTipTitle = "Error";
				Program.Instance.notifyIcon.BalloonTipText = ex.Message;
				Program.Instance.notifyIcon.BalloonTipIcon = ToolTipIcon.Error;
				Program.Instance.notifyIcon.ShowBalloonTip(5000);
                 */
            }
            catch (Exception)
            {
                // this can fail if form was disposed or not created yet, so catch the exception - balloon is not that important to risk followup error
            }
            string message = "Sorry, an unexpected error occured.\nPlease support us fixing this problem.\n\n1. Ensure that you use the latest release of GCSM. You can download the latest version here:\nhttps://sourceforge.net/projects/googlesyncmod/files/latest/download.\n\n2.If the problem still exists, go to\nhttps://sourceforge.net/projects/googlesyncmod/ and use the Tracker!\nPlease check first if error has already been reported.\nProgram Version: {0}\n\nError Details:\n{1}\n\nOS Version: {2}\nOutlook Version: {3}";
            message = string.Format(message, AssemblyVersion, ex.ToString(), OSInfo, OutlookInfo);
            
            try
            {
                Clipboard.SetText(message);
                message += "/n/nHint: This error message is automatically copied to the clipboard.\n";
            
            }
            catch (Exception e)
            {
                Logger.Log("Message couldn't be copied to clipboard: " + e.Message, EventType.Debug);
            }
            //Logger.Log(message, EventType.Debug);
            MessageBox.Show(message, Application.ProductName);
        }

        private static string AssemblyVersion
        {
            get
            {
                return Assembly.GetExecutingAssembly().GetName().Version.ToString();
            }
        }
    }
}