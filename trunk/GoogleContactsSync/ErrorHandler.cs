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
                return VersionInformation.GetWindowsVersionName();
            }
        }

        private static string OutlookInfo
        {
            get
            {
                return VersionInformation.GetOutlookVersion(Synchronizer.OutlookApplication).ToString();
            }
        }

        // TODO: Write a nice error dialog, that maybe supports directly email sending as bugreport
        public static void Handle(Exception ex)
        {
            //save user culture
            CultureInfo oldCI = Thread.CurrentThread.CurrentCulture;
            //set culture to english for exception messages
            Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture ("en-US");
            Thread.CurrentThread.CurrentUICulture=new CultureInfo("en-US");

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
            catch (Exception exc)
            {               
                // this can fail if form was disposed or not created yet, so catch the exception - balloon is not that important to risk followup error
                Logger.Log("Error showing Balloon: " + exc.Message, EventType.Error);
            }
            //create and show error information
            ErrorDialog errorDialog = new ErrorDialog();
            errorDialog.setErrorText(ex);
            errorDialog.Show();

            //set user culture
            Thread.CurrentThread.CurrentCulture = oldCI;
            Thread.CurrentThread.CurrentUICulture = oldCI;
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