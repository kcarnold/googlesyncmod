using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace GoContactSyncMod
{
    static class Program
    {
		private static SettingsForm instance;

        public const int HWND_BROADCAST = 0xffff;
        public static readonly int WM_SHOWME = RegisterWindowMessage("WM_SHOWME");
        [DllImport("user32")]
        public static extern bool PostMessage(IntPtr hwnd, int msg, IntPtr wparam, IntPtr lparam);
        [DllImport("user32")]
        public static extern int RegisterWindowMessage(string message);


        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            //prevent more than one instance of the program
            bool ok;
            System.Threading.Mutex m = new System.Threading.Mutex(true, "acbbbc09-f76c-4874-aaff-4f3353a5a5a6", out ok);

            if (!ok)
            {
                //Message.Create((IntPtr)HWND_BROADCAST, WM_SHOWME, IntPtr.Zero, IntPtr.Zero);                
                PostMessage((IntPtr)HWND_BROADCAST, WM_SHOWME, IntPtr.Zero, IntPtr.Zero);                
                //MessageBox.Show("Another instance of Go Contact Sync Mod is already running.","GO Contact Sync Mod",MessageBoxButtons.OK);
                return;
            }
            AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(CurrentDomain_UnhandledException);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
			instance = new SettingsForm();
            Application.Run(instance);
            GC.KeepAlive(m);
        }

		internal static SettingsForm Instance
		{
			get { return instance; }
		}

        /// <summary>
        /// Fallback. If there is some try/catch missing we will handle it here, just before the application quits unhandled
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            if (e.ExceptionObject is Exception)
                ErrorHandler.Handle((Exception)e.ExceptionObject);
        }
    }
}