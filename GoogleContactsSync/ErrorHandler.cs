﻿using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Reflection;

namespace GoContactSyncMod
{
    static class ErrorHandler
    {
		// TODO: Write a nice error dialog, that maybe supports directly E-Mail sending as bugreport
		public static void Handle(Exception ex)
        {
			Logger.Log(ex.Message, EventType.Error);
			//AppendSyncConsoleText(Logger.GetText());
			Logger.Log("Sync failed.", EventType.Error);

			try
			{
				Program.Instance.notifyIcon.BalloonTipTitle = "Error";
				Program.Instance.notifyIcon.BalloonTipText = ex.Message;
				Program.Instance.notifyIcon.BalloonTipIcon = ToolTipIcon.Error;
				Program.Instance.notifyIcon.ShowBalloonTip(5000);
			}
			catch (Exception)
			{
				// this can fail if form was disposed or not created yet, so catch the exception - balloon is not that important to risk followup error
			}

            string message = "Sorry, an unexpected error occured.\nPlease support us fixing this problem. Go to\nhttps://sourceforge.net/projects/googlesyncmod/ and use the Tracker!\nHint: You can copy this message by pressing CTRL-C in the dialog box.\nPlease check first if error has already been reported.\nProgram Version: {0}\n\nError Details:\n{1}";
			message = string.Format(message, AssemblyVersion, ex.ToString());
            MessageBox.Show(message, "GO Contact Sync Mod");
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