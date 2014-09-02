using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Threading;
using System.Text.RegularExpressions;
using System.Diagnostics;
using System.IO;
using System.Runtime.Remoting;
using System.Runtime.InteropServices;
using System.Collections;


namespace GoContactSyncMod
{
	internal partial class SettingsForm : Form
	{
		internal Syncronizer _sync;
		private SyncOption _syncOption;
		private DateTime lastSync;
		private bool requestClose = false;
        private bool boolShowBalloonTip = true;

        public const string AppRootKey = @"Software\Webgear\GOContactSync";

        private ProxySettingsForm _proxy = new ProxySettingsForm();

        private string _syncContactsFolder = "";
        private string _syncNotesFolder = "";
        private string _syncProfile
        {
            get
            {
                RegistryKey regKeyAppRoot = Registry.CurrentUser.CreateSubKey(AppRootKey);
                return (regKeyAppRoot.GetValue("SyncProfile") != null) ?
                       (string)regKeyAppRoot.GetValue("SyncProfile") : null;
            }
            set
            {
                RegistryKey regKeyAppRoot = Registry.CurrentUser.CreateSubKey(AppRootKey);
                if (!string.IsNullOrEmpty(value))
                    regKeyAppRoot.SetValue("SyncProfile", value);
            }
        }

        //register window for lock/unlock messages of workstation
        private bool registered = false;

		delegate void TextHandler(string text);
		delegate void SwitchHandler(bool value);

		public SettingsForm()
		{
			InitializeComponent();
            Text = Text + " - " + Application.ProductVersion;
			Logger.LogUpdated += new Logger.LogUpdatedHandler(Logger_LogUpdated);
            ContactsMatcher.NotificationReceived += new ContactsMatcher.NotificationHandler(OnNotificationReceived);
            NotesMatcher.NotificationReceived += new NotesMatcher.NotificationHandler(OnNotificationReceived);
			PopulateSyncOptionBox();

            fillSyncFolderItems();

            if (fillSyncProfileItems()) 
                LoadSettings(cmbSyncProfile.Text);
            else 
                LoadSettings(null);

            TimerSwitch(true);
			lastSyncLabel.Text = "Not synced";

			ValidateSyncButton();

            // requires Windows XP or higher
            bool XpOrHigher = Environment.OSVersion.Platform == PlatformID.Win32NT &&
                                (Environment.OSVersion.Version.Major > 5 ||
                                    (Environment.OSVersion.Version.Major == 5 &&
                                     Environment.OSVersion.Version.Minor >= 1));

            if (XpOrHigher)
                registered = WTSRegisterSessionNotification(Handle, 0);
		}

        ~SettingsForm()
        {
            if(registered)
            {
                WTSUnRegisterSessionNotification(Handle);
                registered = false;
            }
            Logger.Close();
        }

		private void PopulateSyncOptionBox()
		{
			string str;
			for (int i = 0; i < 20; i++)
			{
				str = ((SyncOption)i).ToString();
				if (str == i.ToString())
					break;

				// format (to add space before capital)
				MatchCollection matches = Regex.Matches(str, "[A-Z]");
				for (int k = 0; k < matches.Count; k++)
				{
					str = str.Replace(str[matches[k].Index].ToString(), " " + str[matches[k].Index]);
					matches = Regex.Matches(str, "[A-Z]");
				}
				str = str.Replace("  ", " ");
				// fix start
				str = str.Substring(1);

				syncOptionBox.Items.Add(str);
			}
		}
        private void fillSyncFolderItems()
        {
            this.contactFoldersComboBox.Visible = true;
            this.noteFoldersComboBox.Visible = true;
            this.cmbSyncProfile.Visible = true;
            ArrayList outlookContactFolders = new ArrayList();
            ArrayList outlookNoteFolders = new ArrayList();
            try
            {
                this.contactFoldersComboBox.BeginUpdate();
                this.contactFoldersComboBox.Items.Clear();

                Microsoft.Office.Interop.Outlook.Folders folders = Syncronizer.OutlookNameSpace.Folders;
                foreach (Microsoft.Office.Interop.Outlook.Folder folder in folders)
                {
                    try
                    {
                        foreach (Microsoft.Office.Interop.Outlook.MAPIFolder mapi in folder.Folders)
                        {
                            if (mapi.DefaultItemType == Microsoft.Office.Interop.Outlook.OlItemType.olContactItem)
                            {
                                bool isDefaultFolder = mapi.EntryID.Equals(Syncronizer.OutlookNameSpace.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderContacts).EntryID);
                                outlookContactFolders.Add(new OutlookFolder(folder.Name + " - " + mapi.Name, mapi.EntryID, isDefaultFolder));
                            }
                            if (mapi.DefaultItemType == Microsoft.Office.Interop.Outlook.OlItemType.olNoteItem)
                            {
                                bool isDefaultFolder = mapi.EntryID.Equals(Syncronizer.OutlookNameSpace.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderNotes).EntryID);
                                outlookNoteFolders.Add(new OutlookFolder(folder.Name + " - " + mapi.Name, mapi.EntryID, isDefaultFolder));
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        Logger.Log("Error getting available Outlook folders: " + e.Message, EventType.Warning);
                    }
                }
                outlookContactFolders.Sort();
                outlookNoteFolders.Sort();

                this.contactFoldersComboBox.DataSource = outlookContactFolders;
                this.contactFoldersComboBox.DisplayMember = "DisplayName";
                this.contactFoldersComboBox.ValueMember = "FolderID";

                this.noteFoldersComboBox.DataSource = outlookNoteFolders;
                this.noteFoldersComboBox.DisplayMember = "DisplayName";
                this.noteFoldersComboBox.ValueMember = "FolderID";

                this.contactFoldersComboBox.EndUpdate();
                this.noteFoldersComboBox.EndUpdate();

                this.contactFoldersComboBox.SelectedValue = "";
                this.noteFoldersComboBox.SelectedValue = "";

                //Select Default Folder per Default
                foreach (OutlookFolder folder in contactFoldersComboBox.Items)
                    if (folder.IsDefaultFolder)
                    {
                        this.contactFoldersComboBox.SelectedValue = folder.FolderID;
                        break;
                    }

                //Select Default Folder per Default
                foreach (OutlookFolder folder in noteFoldersComboBox.Items)
                    if (folder.IsDefaultFolder)
                    {
                        this.noteFoldersComboBox.SelectedItem = folder;
                        break;
                    }


            }
            catch (Exception e)
            {
                Logger.Log("Error getting available Outlook folders: " + e.Message, EventType.Warning);
            }
        }

        private void ClearSettings()
        {
            SetSyncOption(0);
            UserName.Text = Password.Text = "";
            autoSyncCheckBox.Checked = runAtStartupCheckBox.Checked = reportSyncResultCheckBox.Checked = false;
            autoSyncInterval.Value = 120;
            _proxy.ClearSettings();
        }
        // Fill lists of sync profiles
        private bool fillSyncProfileItems()
        {
            RegistryKey regKeyAppRoot = Registry.CurrentUser.CreateSubKey(AppRootKey);
            bool vReturn = false;

            cmbSyncProfile.Items.Clear();
            cmbSyncProfile.Items.Add("[Add new profile...]");

            foreach (string subKeyName in regKeyAppRoot.GetSubKeyNames())
            {
                cmbSyncProfile.Items.Add(subKeyName);
            }

            if (string.IsNullOrEmpty(_syncProfile))
                _syncProfile = "Default";

            if (cmbSyncProfile.Items.Count == 1)
                cmbSyncProfile.Items.Add(_syncProfile);
            else
                vReturn = true;

            cmbSyncProfile.Items.Add("[Configuration manager...]");
            cmbSyncProfile.Text = _syncProfile;

            return vReturn;
        }
                

        private void LoadSettings(string _profile)
        {
            RegistryKey regKeyAppRoot = Registry.CurrentUser.CreateSubKey(AppRootKey  + (_profile != null ? ('\\' + _profile) : "")  );

            if (regKeyAppRoot.GetValue("SyncOption") != null)
            {
                _syncOption = (SyncOption)regKeyAppRoot.GetValue("SyncOption");
                SetSyncOption((int)_syncOption);
            }

            if (regKeyAppRoot.GetValue("Username") != null)
            {
                UserName.Text = regKeyAppRoot.GetValue("Username") as string;
                if (regKeyAppRoot.GetValue("Password") != null)
                    Password.Text = Encryption.DecryptPassword(UserName.Text, regKeyAppRoot.GetValue("Password") as string);
            }
            if (regKeyAppRoot.GetValue("AutoSync") != null)
                autoSyncCheckBox.Checked = Convert.ToBoolean(regKeyAppRoot.GetValue("AutoSync"));
            if (regKeyAppRoot.GetValue("AutoSyncInterval") != null)
                autoSyncInterval.Value = Convert.ToDecimal(regKeyAppRoot.GetValue("AutoSyncInterval"));
            if (regKeyAppRoot.GetValue("AutoStart") != null)
                runAtStartupCheckBox.Checked = Convert.ToBoolean(regKeyAppRoot.GetValue("AutoStart"));
            if (regKeyAppRoot.GetValue("ReportSyncResult") != null)
                reportSyncResultCheckBox.Checked = Convert.ToBoolean(regKeyAppRoot.GetValue("ReportSyncResult"));            
            if (regKeyAppRoot.GetValue("SyncDeletion") != null)
                btSyncDelete.Checked = Convert.ToBoolean(regKeyAppRoot.GetValue("SyncDeletion"));
            if (regKeyAppRoot.GetValue("PromptDeletion") != null)
                btPromptDelete.Checked = Convert.ToBoolean(regKeyAppRoot.GetValue("PromptDeletion"));        
            if (regKeyAppRoot.GetValue("SyncNotes") != null)
                btSyncNotes.Checked = Convert.ToBoolean(regKeyAppRoot.GetValue("SyncNotes"));
            if (regKeyAppRoot.GetValue("SyncContacts") != null)
                btSyncContacts.Checked = Convert.ToBoolean(regKeyAppRoot.GetValue("SyncContacts"));
            if (regKeyAppRoot.GetValue("SyncContactsFolder") != null)
                contactFoldersComboBox.SelectedValue = regKeyAppRoot.GetValue("SyncContactsFolder") as string;
            if (regKeyAppRoot.GetValue("SyncNotesFolder") != null)
                noteFoldersComboBox.SelectedValue = regKeyAppRoot.GetValue("SyncNotesFolder") as string;

            autoSyncCheckBox_CheckedChanged(null, null);
            btSyncContacts_CheckedChanged(null, null);
            btSyncNotes_CheckedChanged(null, null);

            _proxy.LoadSettings(_profile);
        }

		private void SaveSettings()
		{
            SaveSettings(cmbSyncProfile.Text);
		}

        private void SaveSettings(string profile)
        {
            if (!string.IsNullOrEmpty(profile))
            {
                _syncProfile = cmbSyncProfile.Text;
                RegistryKey regKeyAppRoot = Registry.CurrentUser.CreateSubKey(AppRootKey + "\\" + profile);
                regKeyAppRoot.SetValue("SyncOption", (int)_syncOption);
 
                if (!string.IsNullOrEmpty(UserName.Text))
                {
                    regKeyAppRoot.SetValue("Username", UserName.Text);
                    if (!string.IsNullOrEmpty(Password.Text))
                        regKeyAppRoot.SetValue("Password", Encryption.EncryptPassword(UserName.Text, Password.Text));
                }
                regKeyAppRoot.SetValue("AutoSync", autoSyncCheckBox.Checked.ToString());
                regKeyAppRoot.SetValue("AutoSyncInterval", autoSyncInterval.Value.ToString());
                regKeyAppRoot.SetValue("AutoStart", runAtStartupCheckBox.Checked);
                regKeyAppRoot.SetValue("ReportSyncResult", reportSyncResultCheckBox.Checked);
                regKeyAppRoot.SetValue("SyncDeletion", btSyncDelete.Checked);
                regKeyAppRoot.SetValue("PromptDeletion", btPromptDelete.Checked);
                regKeyAppRoot.SetValue("SyncNotes", btSyncNotes.Checked);
                regKeyAppRoot.SetValue("SyncContacts", btSyncContacts.Checked);

                if (btSyncContacts.Checked && contactFoldersComboBox.SelectedValue != null)
                    regKeyAppRoot.SetValue("SyncContactsFolder", contactFoldersComboBox.SelectedValue.ToString());
                if (btSyncNotes.Checked && noteFoldersComboBox.SelectedValue != null)
                    regKeyAppRoot.SetValue("SyncNotesFolder", noteFoldersComboBox.SelectedValue.ToString());

                _proxy.SaveSettings(cmbSyncProfile.Text);
            }
        }


        private bool ValidCredentials
		{
			get
			{
				bool userNameIsValid = Regex.IsMatch(UserName.Text, @"^(?'id'[a-z0-9\'\%\._\+\-]+)@(?'domain'[a-z0-9\'\%\._\+\-]+)\.(?'ext'[a-z]{2,6})$", RegexOptions.IgnoreCase);
				bool passwordIsValid = Password.Text.Length != 0;
                bool syncProfileIsValid = (cmbSyncProfile.SelectedIndex > 0 && cmbSyncProfile.SelectedIndex < cmbSyncProfile.Items.Count);
                bool syncContactFolderIsValid = (contactFoldersComboBox.SelectedIndex >= 0 && contactFoldersComboBox.SelectedIndex < contactFoldersComboBox.Items.Count) || !btSyncContacts.Checked;
                bool syncNotesFolderIsValid = (noteFoldersComboBox.SelectedIndex >= 0 && noteFoldersComboBox.SelectedIndex < noteFoldersComboBox.Items.Count) || !btSyncNotes.Checked;

				setBgColor(UserName, userNameIsValid);
				setBgColor(Password, passwordIsValid);
                return userNameIsValid && passwordIsValid && syncProfileIsValid && syncContactFolderIsValid && syncNotesFolderIsValid;
			}
		}

		private void setBgColor(TextBox box, bool isValid)
		{
			if (!isValid)
				box.BackColor = Color.LightPink;
			else
				box.BackColor = Color.LightGreen;
		}

		private void syncButton_Click(object sender, EventArgs e)
		{
			Sync();
		}
		private void Sync()
		{
			try
			{
				if (!ValidCredentials)
					return;

				ThreadStart starter = new ThreadStart(Sync_ThreadStarter);
				Thread thread = new Thread(starter);
				thread.Start();

				// wait for thread to start
				while (!thread.IsAlive)
					Thread.Sleep(1);
			}
			catch (Exception ex)
			{
				ErrorHandler.Handle(ex);
			}
		}

		private void Sync_ThreadStarter()
		{
            try
            {
                TimerSwitch(false);

                //if the contacts or notes folder has changed ==> Reset matches (to not delete contacts or notes on the one or other side)                
                RegistryKey regKeyAppRoot = Registry.CurrentUser.CreateSubKey(AppRootKey + "\\" + _syncProfile);
                string syncContactsFolder = regKeyAppRoot.GetValue("SyncContactsFolder") as string;
                string syncNotesFolder = regKeyAppRoot.GetValue("SyncNotesFolder") as string;

                //only reset notes if NotesFolder changed and reset contacts if ContactsFolder changed
                bool syncContacts = !string.IsNullOrEmpty(syncContactsFolder) && !syncContactsFolder.Equals(_syncContactsFolder) && btSyncContacts.Checked;
                bool syncNotes = !string.IsNullOrEmpty(syncNotesFolder) && !syncNotesFolder.Equals(_syncNotesFolder) && btSyncNotes.Checked;                                
                if (syncContacts || syncNotes)
                    ResetMatches(syncContacts, syncNotes);
                
                //Then save the Contacts and Notes Folders used at last sync
                if (btSyncContacts.Checked)
                    regKeyAppRoot.SetValue("SyncContactsFolder", _syncContactsFolder);
                if (btSyncNotes.Checked)
                    regKeyAppRoot.SetValue("SyncNotesFolder", _syncNotesFolder);

                SetLastSyncText("Syncing...");
                notifyIcon.Text = Application.ProductName + "\nSyncing...";
                SetFormEnabled(false);

                if (_sync == null)
                {
                    _sync = new Syncronizer();
                    _sync.DuplicatesFound += new Syncronizer.DuplicatesFoundHandler(OnDuplicatesFound);
                    _sync.ErrorEncountered += new Syncronizer.ErrorNotificationHandler(OnErrorEncountered);
                }

                Logger.ClearLog();
                SetSyncConsoleText("");
                Logger.Log("Sync started (" + _syncProfile + ").", EventType.Information);
                //SetSyncConsoleText(Logger.GetText());
                _sync.SyncProfile = _syncProfile;
                _sync.SyncContactsFolder  = _syncContactsFolder;
                _sync.SyncNotesFolder = _syncNotesFolder;

                _sync.SyncOption = _syncOption;
                _sync.SyncDelete = btSyncDelete.Checked;
                _sync.PromptDelete = btPromptDelete.Checked && btSyncDelete.Checked;
                _sync.SyncNotes = btSyncNotes.Checked;
                _sync.SyncContacts = btSyncContacts.Checked;

                if (!_sync.SyncContacts && !_sync.SyncNotes)
                {
                    SetLastSyncText("Sync failed.");
                    notifyIcon.Text = Application.ProductName + "\nSync failed";

                    string messageText = "Neither notes nor contacts are switched on for syncing. Please choose at least one option. Sync aborted!";
                    Logger.Log(messageText, EventType.Error);
                    ShowForm();
                    Program.Instance.ShowBalloonToolTip("Error", messageText, ToolTipIcon.Error, 5000);
                    return;
                }


                _sync.LoginToGoogle(UserName.Text, Password.Text);
                _sync.LoginToOutlook();

                _sync.Sync();

                lastSync = DateTime.Now;
                SetLastSyncText("Last synced at " + lastSync.ToString());

                string message = string.Format("Sync complete.\r\n Synced:  {1} out of {0}.\r\n Deleted:  {2}.\r\n Skipped: {3}.\r\n Errors:    {4}.", _sync.TotalCount, _sync.SyncedCount, _sync.DeletedCount, _sync.SkippedCount, _sync.ErrorCount);
                Logger.Log(message, EventType.Information);
                if (reportSyncResultCheckBox.Checked)
                {
                    /*
                    notifyIcon.BalloonTipTitle = Application.ProductName;
                    notifyIcon.BalloonTipText = string.Format("{0}. {1}", DateTime.Now, message);
                    */
                    ToolTipIcon icon;
                    if (_sync.ErrorCount > 0)
                        icon = ToolTipIcon.Error;
                    else if (_sync.SkippedCount > 0)
                        icon = ToolTipIcon.Warning;
                    else
                        icon = ToolTipIcon.Info;
                    /*notifyIcon.ShowBalloonTip(5000);
                    */
                    ShowBalloonToolTip(Application.ProductName,
                        string.Format("{0}. {1}", DateTime.Now, message),
                        icon,
                        5000);

                }
                string toolTip = string.Format("{0}\nLast sync: {1}", Application.ProductName, DateTime.Now.ToString("dd.MM. HH:mm"));
                if (_sync.ErrorCount + _sync.SkippedCount > 0)
                    toolTip += string.Format("\nWarnings: {0}.", _sync.ErrorCount + _sync.SkippedCount);
                if (toolTip.Length >= 64)
                    toolTip = toolTip.Substring(0, 63);
                notifyIcon.Text = toolTip;
            }
            catch (Google.GData.Client.GDataRequestException ex)
            {
                SetLastSyncText("Sync failed.");
                notifyIcon.Text = Application.ProductName + "\nSync failed";

                //string responseString = (null != ex.InnerException) ? ex.ResponseString : ex.Message;

                if (ex.InnerException is System.Net.WebException)
                {
                    string message = "Cannot connect to Google, please check for available internet connection and proxy settings if applicable: " + ex.InnerException.Message + "\r\n" + ex.ResponseString;
                    Logger.Log(message, EventType.Warning);
                    Program.Instance.ShowBalloonToolTip("Error", message, ToolTipIcon.Error, 5000);
                }
                else
                {
                    ErrorHandler.Handle(ex);
                }
            }
            catch (Google.GData.Client.InvalidCredentialsException)
            {
                SetLastSyncText("Sync failed.");
                notifyIcon.Text = Application.ProductName + "\nSync failed";

                string message = "The credentials (Google Account username and/or password) are invalid, please correct them in the settings form before you sync again";
                Logger.Log(message, EventType.Error);
                ShowForm();
                Program.Instance.ShowBalloonToolTip("Error", message, ToolTipIcon.Error, 5000);

            }
            catch (Exception ex)
            {
                SetLastSyncText("Sync failed.");
                notifyIcon.Text = Application.ProductName + "\nSync failed";

                if (ex is COMException)
                {
                    string message = "Outlook exception, please assure that Outlook is running and not closed when syncing";
                    Logger.Log(message + ": " + ex.Message, EventType.Warning);
                    Program.Instance.ShowBalloonToolTip("Error", message, ToolTipIcon.Error, 5000);
                }
                else
                {
                    ErrorHandler.Handle(ex);
                }
            }							
			finally
			{                        
                lastSync = DateTime.Now;
                TimerSwitch(true);
				SetFormEnabled(true);
                if (_sync != null)
                {
                    _sync.LogoffOutlook();
                    _sync.LogoffGoogle();
                    _sync = null;
                }
			}
		}

        public void ShowBalloonToolTip(string title, string message, ToolTipIcon icon, int timeout)
        {
            //if user is active on workstation
            if(boolShowBalloonTip)
            {
                notifyIcon.BalloonTipTitle = title;
			    notifyIcon.BalloonTipText = message;
			    notifyIcon.BalloonTipIcon = icon;
			    notifyIcon.ShowBalloonTip(timeout);
            }
        }

		void Logger_LogUpdated(string Message)
		{
			AppendSyncConsoleText(Message);
		}

		void OnErrorEncountered(string title, Exception ex, EventType eventType)
		{
			// do not show ErrorHandler, as there may be multiple exceptions that would nag the user
			Logger.Log(ex.ToString(), EventType.Error);
			string message = String.Format("Error Saving Contact: {0}.\nPlease report complete ErrorMessage from Log to the Tracker\nat https://sourceforge.net/tracker/?group_id=369321", ex.Message);
            ShowBalloonToolTip(title,message,ToolTipIcon.Error,5000);
			/*notifyIcon.BalloonTipTitle = title;
			notifyIcon.BalloonTipText = message;
			notifyIcon.BalloonTipIcon = ToolTipIcon.Error;
			notifyIcon.ShowBalloonTip(5000);*/
		}

		void OnDuplicatesFound(string title, string message)
		{
            Logger.Log(message, EventType.Warning);
            ShowBalloonToolTip(title,message,ToolTipIcon.Warning,5000);
            /*
			notifyIcon.BalloonTipTitle = title;
			notifyIcon.BalloonTipText = message;
			notifyIcon.BalloonTipIcon = ToolTipIcon.Warning;
			notifyIcon.ShowBalloonTip(5000);
             */
		}

        void OnNotificationReceived(string message)
        {
            SetLastSyncText(message);           
        }

		public void SetFormEnabled(bool enabled)
		{
			if (this.InvokeRequired)
			{
				SwitchHandler h = new SwitchHandler(SetFormEnabled);
				this.Invoke(h, new object[] { enabled });
			}
			else
			{
				resetMatchesLinkLabel.Enabled = enabled;
				settingsGroupBox.Enabled = enabled;
				syncButton.Enabled = enabled;
			}
		}
		public void SetLastSyncText(string text)
		{
			if (this.InvokeRequired)
			{
				TextHandler h = new TextHandler(SetLastSyncText);
				this.Invoke(h, new object[] { text });
			}
			else
				lastSyncLabel.Text = text;
		}
		public void SetSyncConsoleText(string text)
		{
			if (this.InvokeRequired)
			{
				TextHandler h = new TextHandler(SetSyncConsoleText);
				this.Invoke(h, new object[] { text });
			}
			else
            {
				syncConsole.Text = text;
                //Scroll to bottom to always see the last log entry
                syncConsole.SelectionStart = syncConsole.TextLength;
                syncConsole.ScrollToCaret();
            }

		}
		public void AppendSyncConsoleText(string text)
		{
			if (this.InvokeRequired)
			{
				TextHandler h = new TextHandler(AppendSyncConsoleText);
				this.Invoke(h, new object[] { text });
			}
			else
            {
				syncConsole.Text += text;
                //Scroll to bottom to always see the last log entry
                syncConsole.SelectionStart = syncConsole.TextLength;
                syncConsole.ScrollToCaret();
            }
		}
		public void TimerSwitch(bool value)
		{
			if (this.InvokeRequired)
			{
				SwitchHandler h = new SwitchHandler(TimerSwitch);
				this.Invoke(h, new object[] { value });
			}
			else
			{
                //If PC resumes or unlocks or is started, give him 90 seconds to recover everything before the sync starts
                if (lastSync <= DateTime.Now.AddSeconds(90) - new TimeSpan(0, (int)autoSyncInterval.Value, 0))
                    lastSync = DateTime.Now.AddSeconds(90) - new TimeSpan(0, (int)autoSyncInterval.Value, 0);
				autoSyncInterval.Enabled = autoSyncCheckBox.Checked && value;
				syncTimer.Enabled = autoSyncCheckBox.Checked && value;
				nextSyncLabel.Visible = autoSyncCheckBox.Checked &&  value;          				
			}
		}

        //to detect if the user locks or unlocks the workstation
        [DllImport("wtsapi32.dll")]
        private static extern bool WTSRegisterSessionNotification(IntPtr hWnd, int dwFlags);

        [DllImport("wtsapi32.dll")]
        private static extern bool WTSUnRegisterSessionNotification(IntPtr hWnd);

		// Fix for WinXP and older systems, that do not continue with shutdown until all programs have closed
		// FormClosing would hold system shutdown, when it sets the cancel to true
		private const int WM_QUERYENDSESSION = 0x11;

        //Code to find out if workstation is locked
        private const int WM_WTSSESSION_CHANGE = 0x02B1;
        private const int WTS_SESSION_LOCK = 0x7;
        private const int WTS_SESSION_UNLOCK = 0x8;

        //Code to find if workstation is resumed
        const int WM_POWERBROADCAST = 0x0218;
        const int PBT_APMQUERYSUSPEND = 0x0000;
        const int PBT_APMQUERYSTANDBY = 0x0001;
        const int PBT_APMQUERYSUSPENDFAILED = 0x0002;
        const int PBT_APMQUERYSTANDBYFAILED = 0x0003;
        const int PBT_APMSUSPEND = 0x0004;
        const int PBT_APMSTANDBY = 0x0005;
        const int PBT_APMRESUMECRITICAL = 0x0006;
        const int PBT_APMRESUMESUSPEND = 0x0007;
        const int PBT_APMRESUMESTANDBY = 0x0008;       
        const int PBT_APMRESUMEAUTOMATIC = 0x0012;        

        
        /*
        protected void OnSessionLock()
        {
            Logger.Log("Locked at " + DateTime.Now + Environment.NewLine, EventType.Information);
        }

        protected void OnSessionUnlock()
        {
            Logger.Log("Unlocked at " + DateTime.Now + Environment.NewLine, EventType.Information);
        }
        */

        protected override void WndProc(ref System.Windows.Forms.Message m)
		{
            //Logger.Log(m.Msg, EventType.Information);
            switch(m.Msg) 
            {
                
                case WM_QUERYENDSESSION:
                    requestClose = true;
                    break;
                case WM_WTSSESSION_CHANGE:
                    {
                        if (m.WParam.ToInt32() == WTS_SESSION_LOCK)
                        {
                            //Logger.Log("\nBenutzer aktiv -> ToolTip", EventType.Information);
                            //OnSessionLock();
                            boolShowBalloonTip = false; // Do something when locked
                        }
                        else if (m.WParam.ToInt32() == WTS_SESSION_UNLOCK)
                        {
                            //Logger.Log("\nBenutzer inaktiv -> kein ToolTip", EventType.Information);
                            //OnSessionUnlock();
                            boolShowBalloonTip = true; // Do something when unlocked
                            TimerSwitch(true);
                        }
                     break;
                    }                
                case WM_POWERBROADCAST:
                    {
                        if (m.WParam.ToInt32() == PBT_APMRESUMEAUTOMATIC ||
                            m.WParam.ToInt32() == PBT_APMRESUMECRITICAL ||
                            m.WParam.ToInt32() == PBT_APMRESUMESTANDBY ||
                            m.WParam.ToInt32() == PBT_APMRESUMESUSPEND ||
                            m.WParam.ToInt32() == PBT_APMQUERYSTANDBYFAILED ||
                            m.WParam.ToInt32() == PBT_APMQUERYSTANDBYFAILED)
                        {                            
                            TimerSwitch(true);
                        }
                        else if (m.WParam.ToInt32() == PBT_APMSUSPEND ||
                                 m.WParam.ToInt32() == PBT_APMSTANDBY ||
                                 m.WParam.ToInt32() == PBT_APMQUERYSTANDBY ||
                                 m.WParam.ToInt32() == PBT_APMQUERYSUSPEND)
                        {
                            TimerSwitch(false);
                        }
                            

                        break;
                    }
                default:
                    break;
            }
            /*
			if (m.Msg == WM_QUERYENDSESSION)
				requestClose = true;
            if (m.Msg == SESSIONCHANGEMESSAGE)
            {
                if (m.WParam.ToInt32() == SESSIONLOCKPARAM)
                    OnSessionLock(); // Do something when locked
                else if (m.WParam.ToInt32() == SESSIONUNLOCKPARAM)
                    OnSessionUnlock(); // Do something when unlocked
            }*/
			// If this is WM_QUERYENDSESSION, the form must exit and not just hide

            if (m.Msg == Program.WM_SHOWME)
                ShowForm();
			base.WndProc(ref m);
		} 

		private void SettingsForm_FormClosing(object sender, FormClosingEventArgs e)
		{
			if (!requestClose)
			{
				SaveSettings();
				e.Cancel = true;
			}
			HideForm();
		}
		private void SettingsForm_FormClosed(object sender, FormClosedEventArgs e)
		{
			try
			{
				if (_sync != null)
					_sync.LogoffOutlook();

				SaveSettings();

				notifyIcon.Dispose();
			}
			catch (Exception ex)
			{
				ErrorHandler.Handle(ex);
			}
		}

		private void syncOptionBox_SelectedIndexChanged(object sender, EventArgs e)
		{
			try
			{
				int index = syncOptionBox.SelectedIndex;
				if (index == -1)
					return;

				SetSyncOption(index);
			}
			catch (Exception ex)
			{
				ErrorHandler.Handle(ex);
			}
		}
		private void SetSyncOption(int index)
		{
			_syncOption = (SyncOption)index;
			for (int i = 0; i < syncOptionBox.Items.Count; i++)
			{
				if (i == index)
					syncOptionBox.SetItemCheckState(i, CheckState.Checked);
				else
					syncOptionBox.SetItemCheckState(i, CheckState.Unchecked);
			}
		}

		private void SettingsForm_Resize(object sender, EventArgs e)
		{
			if (WindowState == FormWindowState.Minimized)
				Hide();

		}

		private void notifyIcon_MouseDoubleClick(object sender, MouseEventArgs e)
		{
			if (WindowState == FormWindowState.Normal)
				HideForm();
			else
				ShowForm();
		}

		private void autoSyncCheckBox_CheckedChanged(object sender, EventArgs e)
		{
            lastSync = DateTime.Now.AddSeconds(90) - new TimeSpan(0, (int)autoSyncInterval.Value, 0);
            TimerSwitch(true);
		}

		private void syncTimer_Tick(object sender, EventArgs e)
		{
			
			TimeSpan syncTime = DateTime.Now - lastSync;
			TimeSpan limit = new TimeSpan(0, (int)autoSyncInterval.Value, 0);
            if (syncTime < limit)
            {
                TimeSpan diff = limit - syncTime;
                string str = "Next sync in";
                if (diff.Hours != 0)
                    str += " " + diff.Hours + " h";
                if (diff.Minutes != 0 || diff.Hours != 0)
                    str += " " + diff.Minutes + " min";
                if (diff.Seconds != 0)
                    str += " " + diff.Seconds + " s";
                nextSyncLabel.Text = str;
            }
            else
            {
                Sync();
            }
		}

		private void resetMatchesLinkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

			// force deactivation to show up
			Application.DoEvents();
			try
			{
                ResetMatches(btSyncContacts.Checked,btSyncNotes.Checked);                
			}
			catch (Exception ex)
            {
                SetLastSyncText("Reset Matches failed");
                Logger.Log("Reset Matches failed", EventType.Error);
				ErrorHandler.Handle(ex);
			}
			finally
			{                
                lastSync = DateTime.Now;
                TimerSwitch(true);
				SetFormEnabled(true);
                this.hideButton.Enabled = true;
                if (_sync != null)
                {
                    _sync.LogoffOutlook();
                    _sync.LogoffGoogle();
                    _sync = null;
                }
			}
		}

        private void ResetMatches(bool syncContacts, bool syncNotes)
        {
            TimerSwitch(false);
            SetLastSyncText("Resetting matches...");
            notifyIcon.Text = Application.ProductName + "\nResetting matches...";
            SetFormEnabled(false);
            //this.hideButton.Enabled = false;

            if (_sync == null)
            {
                _sync = new Syncronizer();
            }

            Logger.ClearLog();
            SetSyncConsoleText("");
            Logger.Log("Reset Matches started  (" + _syncProfile + ").", EventType.Information);

            _sync.SyncNotes = syncNotes;
            _sync.SyncContacts = syncContacts;

            _sync.SyncContactsFolder = _syncContactsFolder;
            _sync.SyncNotesFolder = _syncNotesFolder;
            _sync.SyncProfile = _syncProfile;

            _sync.LoginToGoogle(UserName.Text, Password.Text);
            _sync.LoginToOutlook();

            

            //Load matches, but match them by properties, not sync id

            if (_sync.SyncContacts)
            {
                _sync.LoadContacts();
                _sync.ResetContactMatches();
            }


            if (_sync.SyncNotes)
            {
                _sync.LoadNotes();
                _sync.ResetNoteMatches();
            }



            lastSync = DateTime.Now;
            SetLastSyncText("Matches reset at " + lastSync.ToString());
            Logger.Log("Matches reset.", EventType.Information);
        }

        private delegate void InvokeCallback(); 

        private void ShowForm()
        {
            if (this.InvokeRequired)
            {
                Invoke(new InvokeCallback(ShowForm));
            }
            else
            {
                Show();
                Activate();
                WindowState = FormWindowState.Normal;
            }
        }
		private void HideForm()
		{
			WindowState = FormWindowState.Minimized;
			Hide();
		}

		private void toolStripMenuItem1_Click(object sender, EventArgs e)
		{
			ShowForm();
            this.Activate();
		}
		private void toolStripMenuItem3_Click(object sender, EventArgs e)
		{
			HideForm();
		}
		private void toolStripMenuItem2_Click(object sender, EventArgs e)
		{
			requestClose = true;
			Close();
		}
		private void toolStripMenuItem5_Click(object sender, EventArgs e)
		{
			AboutBox about = new AboutBox();
			about.Show();
		}
		private void toolStripMenuItem4_Click(object sender, EventArgs e)
		{
			Sync();
		}

		private void SettingsForm_Load(object sender, EventArgs e)
		{
			if (string.IsNullOrEmpty(UserName.Text) ||
				string.IsNullOrEmpty(Password.Text) ||
                string.IsNullOrEmpty(cmbSyncProfile.Text) ||
                string.IsNullOrEmpty(contactFoldersComboBox.Text) )
			{
				// this is the first load, show form
				ShowForm();
				UserName.Focus();
                ShowBalloonToolTip(Application.ProductName,
                        "Application started and visible in your PC's system tray, click on this balloon or the icon below to open the settings form and enter your Google credentials there.",
                        ToolTipIcon.Info,
                        5000);
			}
			else
				HideForm();
		}

		private void runAtStartupCheckBox_CheckedChanged(object sender, EventArgs e)
		{
			RegistryKey regKeyAppRoot = Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run");

			if (runAtStartupCheckBox.Checked)
			{
				// add to registry
				regKeyAppRoot.SetValue("GoogleContactSync", Application.ExecutablePath);
			}
			else
			{
				// remove from registry
				regKeyAppRoot.DeleteValue("GoogleContactSync");
			}
		}

		private void UserName_TextChanged(object sender, EventArgs e)
		{
			ValidateSyncButton();
		}
		private void Password_TextChanged(object sender, EventArgs e)
		{
			ValidateSyncButton();
		}

		private void ValidateSyncButton()
		{
			syncButton.Enabled = ValidCredentials;
		}

		private void deleteDuplicatesButton_Click(object sender, EventArgs e)
		{
			//DeleteDuplicatesForm f = new DeleteDuplicatesForm(_sync
		}		       

		private void Donate_Click(object sender, EventArgs e)
		{
			System.Diagnostics.Process.Start("https://sourceforge.net/project/project_donations.php?group_id=369321");
		}

		private void Donate_MouseEnter(object sender, EventArgs e)
		{
			Donate.BackColor = System.Drawing.Color.LightGray;
		}

		private void Donate_MouseLeave(object sender, EventArgs e)
		{
			Donate.BackColor = System.Drawing.Color.Transparent;
		}

		private void hideButton_Click(object sender, EventArgs e)
		{
			this.Close();
		}

		private void proxySettingsLinkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
		{
            if (_proxy != null) _proxy.ShowDialog();
        }

		private void SettingsForm_HelpButtonClicked(object sender, CancelEventArgs e)
		{
			ShowHelp();
		}

		private void SettingsForm_HelpRequested(object sender, HelpEventArgs hlpevent)
		{
			ShowHelp();
		}

		private void ShowHelp()
		{
			// go to the page showing the help and howto instructions
			Process.Start("http://googlesyncmod.sourceforge.net/");
		}

        private void btSyncContacts_CheckedChanged(object sender, EventArgs e)
        {
            if (!btSyncContacts.Checked && !btSyncNotes.Checked)
            {
                MessageBox.Show("Neither notes nor contacts are switched on for syncing. Please choose at least one option (automatically switched on notes for syncing now).", "No sync switched on");
                btSyncNotes.Checked = true;
            }
            contactFoldersComboBox.Visible = btSyncContacts.Checked;
        }

        private void btSyncNotes_CheckedChanged(object sender, EventArgs e)
        {
            if (!btSyncContacts.Checked && !btSyncNotes.Checked)
            {
                MessageBox.Show("Neither notes nor contacts are switched on for syncing. Please choose at least one option (automatically switched on contacts for syncing now).", "No sync switched on");
                btSyncContacts.Checked = true;
            }
            noteFoldersComboBox.Visible = btSyncNotes.Checked;
        }
    	
        private void cmbSyncProfile_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox comboBox = (ComboBox)sender;

            if ((0 == comboBox.SelectedIndex) || (comboBox.SelectedIndex == (comboBox.Items.Count - 1))) {
                ConfigurationManagerForm _configs = new ConfigurationManagerForm();

                if (0 == comboBox.SelectedIndex && _configs != null)
                {
                    _syncProfile = _configs.AddProfile();
                    ClearSettings();
                }
                
                if (comboBox.SelectedIndex == (comboBox.Items.Count - 1) && _configs != null)
                    _configs.ShowDialog();

                fillSyncProfileItems();

                comboBox.Text = _syncProfile;
                SaveSettings();
            }
            if (comboBox.SelectedIndex < 0)
                MessageBox.Show("Please select Sync Profile.", "No sync switched on");
            else
            {
                //ClearSettings();
                LoadSettings(comboBox.Text);
                _syncProfile = comboBox.Text;
            }

            ValidateSyncButton();
        }

        private void contacFoldersComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            string message = "Select the Outlook Contacts folder you want to sync";
            if ((sender as ComboBox).SelectedIndex >= 0 && (sender as ComboBox).SelectedIndex < (sender as ComboBox).Items.Count && (sender as ComboBox).SelectedItem is OutlookFolder)
            {
                _syncContactsFolder = (sender as ComboBox).SelectedValue.ToString();
                toolTip.SetToolTip((sender as ComboBox), message + ":\r\n" + ((OutlookFolder)(sender as ComboBox).SelectedItem).DisplayName);
            }
            else
            {
                _syncContactsFolder = "";
                toolTip.SetToolTip((sender as ComboBox), message);
            }
            ValidateSyncButton();

            
        }

        private void noteFoldersComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            string message = "Select the Outlook Notes folder you want to sync";
            if ((sender as ComboBox).SelectedIndex >= 0 && (sender as ComboBox).SelectedIndex < (sender as ComboBox).Items.Count && (sender as ComboBox).SelectedItem is OutlookFolder)
            {
                _syncNotesFolder = (sender as ComboBox).SelectedValue.ToString();
                toolTip.SetToolTip((sender as ComboBox), message + ":\r\n" + ((OutlookFolder)(sender as ComboBox).SelectedItem).DisplayName);
            }
            else
            {
                _syncNotesFolder = "";
                toolTip.SetToolTip((sender as ComboBox), message);
            }

            ValidateSyncButton();

            
        }

        private void btSyncDelete_CheckedChanged(object sender, EventArgs e)
        {
            btPromptDelete.Visible = btSyncDelete.Checked;
            btPromptDelete.Checked = btSyncDelete.Checked;
        }        

	}
}