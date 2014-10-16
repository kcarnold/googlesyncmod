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
using System.Globalization;


namespace GoContactSyncMod
{
	internal partial class SettingsForm : Form
	{
        //Singleton-Object
        #region Singleton Definition

        private static volatile SettingsForm instance;
        private static object syncRoot = new Object();

        public static SettingsForm Instance
        {
            get
            {
                if (instance == null)
                {
                    lock (syncRoot)
                    {
                        if (instance == null)
                            instance = new SettingsForm();
                    }
                }

                return instance;
            }
        }
        #endregion

        internal Syncronizer sync;
		private SyncOption syncOption;
		private DateTime lastSync;
		private bool requestClose = false;
        private bool boolShowBalloonTip = true;

        public const string AppRootKey = @"Software\Webgear\GOContactSync";

        private ProxySettingsForm _proxy = new ProxySettingsForm();

        private string syncContactsFolder = "";
        private string syncNotesFolder = "";
        private string syncAppointmentsFolder = "";
        private string Timezone = "";

        private string syncProfile
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


        private int executing; // make this static if you want this one-caller-only to
        // all objects instead of a single object

        Thread syncThread;

        //register window for lock/unlock messages of workstation
        //private bool registered = false;

		delegate void TextHandler(string text);
		delegate void SwitchHandler(bool value);
        delegate void IconHandler();

        private Icon IconError = GoContactSyncMod.Properties.Resources.sync_error;
        private Icon Icon0 = GoContactSyncMod.Properties.Resources.sync;
        private Icon Icon30 = GoContactSyncMod.Properties.Resources.sync_30;
        private Icon Icon60 = GoContactSyncMod.Properties.Resources.sync_60;
        private Icon Icon90 = GoContactSyncMod.Properties.Resources.sync_90;
        private Icon Icon120 = GoContactSyncMod.Properties.Resources.sync_120;
        private Icon Icon150 = GoContactSyncMod.Properties.Resources.sync_150;
        private Icon Icon180 = GoContactSyncMod.Properties.Resources.sync_180;
        private Icon Icon210 = GoContactSyncMod.Properties.Resources.sync_210;
        private Icon Icon240 = GoContactSyncMod.Properties.Resources.sync_240;
        private Icon Icon270 = GoContactSyncMod.Properties.Resources.sync_270;
        private Icon Icon300 = GoContactSyncMod.Properties.Resources.sync_300;
        private Icon Icon330 = GoContactSyncMod.Properties.Resources.sync_330;

		private SettingsForm()
		{
			InitializeComponent();
            Text = Text + " - " + Application.ProductVersion;
			Logger.LogUpdated += new Logger.LogUpdatedHandler(Logger_LogUpdated);
            Logger.Log("Started application " + Application.ProductName + " " + Application.ProductVersion + " on " + VersionInformation.GetWindowsVersionName(), EventType.Information);
            ContactsMatcher.NotificationReceived += new ContactsMatcher.NotificationHandler(OnNotificationReceived);
            NotesMatcher.NotificationReceived += new NotesMatcher.NotificationHandler(OnNotificationReceived);
            AppointmentsMatcher.NotificationReceived += new AppointmentsMatcher.NotificationHandler(OnNotificationReceived);
			PopulateSyncOptionBox();            

            if (fillSyncProfileItems()) 
                LoadSettings(cmbSyncProfile.Text);
            else 
                LoadSettings(null);

            TimerSwitch(true);
			lastSyncLabel.Text = "Not synced";

			ValidateSyncButton();

            //Register Session Lock Event
            SystemEvents.SessionSwitch += new SessionSwitchEventHandler(SystemEvents_SessionSwitch);
            //Register Power Mode Event
            SystemEvents.PowerModeChanged += new PowerModeChangedEventHandler(SystemEvents_PowerModeSwitch);
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
            lock (syncRoot)
            {                
                if (this.contactFoldersComboBox.DataSource == null || this.noteFoldersComboBox.DataSource == null || this.appointmentFoldersComboBox.DataSource == null ||
                    this.contactFoldersComboBox.Items.Count == 0 || this.noteFoldersComboBox.Items.Count == 0 || this.appointmentFoldersComboBox.Items.Count == 0)
                {
                    Logger.Log("Loading Outlook folders...", EventType.Information);

                    this.contactFoldersComboBox.Visible = btSyncContacts.Checked;
                    this.noteFoldersComboBox.Visible = btSyncNotes.Checked;
                    this.appointmentFoldersComboBox.Visible = this.futureMonthTextBox.Visible = this.pastMonthTextBox.Visible = this.appointmentTimezonesComboBox.Visible = btSyncAppointments.Checked;
                    this.cmbSyncProfile.Visible = true;
                    ArrayList outlookContactFolders = new ArrayList();
                    ArrayList outlookNoteFolders = new ArrayList();
                    ArrayList outlookAppointmentFolders = new ArrayList();
                    try
                    {
                        Cursor = Cursors.WaitCursor;
                        SuspendLayout();

                        this.contactFoldersComboBox.BeginUpdate();
                        this.noteFoldersComboBox.BeginUpdate();
                        this.appointmentFoldersComboBox.BeginUpdate();
                        this.contactFoldersComboBox.DataSource = null;
                        this.noteFoldersComboBox.DataSource = null;
                        this.appointmentFoldersComboBox.DataSource = null;
                        //this.contactFoldersComboBox.Items.Clear();

                        Microsoft.Office.Interop.Outlook.Folders folders = Syncronizer.OutlookNameSpace.Folders;
                        foreach (Microsoft.Office.Interop.Outlook.Folder folder in folders)
                        {
                            try
                            {
                                GetOutlookMAPIFolders(outlookContactFolders, outlookNoteFolders, outlookAppointmentFolders, folder);
                            }
                            catch (Exception e)
                            {
                                Logger.Log("Error getting available Outlook folders: " + e.Message, EventType.Warning);
                            }
                        }
                                                                        

                        if (outlookContactFolders != null && outlookContactFolders.Count > 0)
                        {
                            outlookContactFolders.Sort();
                            this.contactFoldersComboBox.DataSource = outlookContactFolders;
                            this.contactFoldersComboBox.DisplayMember = "DisplayName";
                            this.contactFoldersComboBox.ValueMember = "FolderID";
                        }

                        if (outlookNoteFolders != null && outlookNoteFolders.Count > 0)
                        {
                            outlookNoteFolders.Sort();
                            this.noteFoldersComboBox.DataSource = outlookNoteFolders;
                            this.noteFoldersComboBox.DisplayMember = "DisplayName";
                            this.noteFoldersComboBox.ValueMember = "FolderID";
                        }

                        if (outlookAppointmentFolders != null && outlookAppointmentFolders.Count > 0)
                        {
                            outlookAppointmentFolders.Sort();
                            this.appointmentFoldersComboBox.DataSource = outlookAppointmentFolders;
                            this.appointmentFoldersComboBox.DisplayMember = "DisplayName";
                            this.appointmentFoldersComboBox.ValueMember = "FolderID";
                        }
                        this.contactFoldersComboBox.EndUpdate();
                        this.noteFoldersComboBox.EndUpdate();
                        this.appointmentFoldersComboBox.EndUpdate();

                        this.contactFoldersComboBox.SelectedValue = "";
                        this.noteFoldersComboBox.SelectedValue = "";
                        this.appointmentFoldersComboBox.SelectedValue = "";

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

                        //Select Default Folder per Default
                        foreach (OutlookFolder folder in appointmentFoldersComboBox.Items)
                            if (folder.IsDefaultFolder)
                            {
                                this.appointmentFoldersComboBox.SelectedItem = folder;
                                break;
                            }

                        Logger.Log("Loaded Outlook folders.", EventType.Information);

                    }
                    catch (Exception e)
                    {
                        Logger.Log("Error getting available Outlook folders: " + e.Message, EventType.Warning);
                    }
                    finally
                    {
                        Cursor = Cursors.Default;
                        ResumeLayout();
                    }

                    LoadSettingsFolders(syncProfile);
                }
            }
        }

        public static void GetOutlookMAPIFolders(ArrayList outlookContactFolders, ArrayList outlookNoteFolders, ArrayList outlookAppointmentFolders, Microsoft.Office.Interop.Outlook.MAPIFolder folder)
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
                if (mapi.DefaultItemType == Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem)
                {
                    bool isDefaultFolder = mapi.EntryID.Equals(Syncronizer.OutlookNameSpace.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderCalendar).EntryID);
                    outlookAppointmentFolders.Add(new OutlookFolder(folder.Name + " - " + mapi.Name, mapi.EntryID, isDefaultFolder));
                }

                if (mapi.DefaultItemType == Microsoft.Office.Interop.Outlook.OlItemType.olContactItem ||
                    mapi.DefaultItemType == Microsoft.Office.Interop.Outlook.OlItemType.olNoteItem ||
                    mapi.DefaultItemType == Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem)
                    GetOutlookMAPIFolders(outlookContactFolders, outlookNoteFolders, outlookAppointmentFolders, mapi);

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

            if (string.IsNullOrEmpty(syncProfile))
                syncProfile = "Default";

            if (cmbSyncProfile.Items.Count == 1)
                cmbSyncProfile.Items.Add(syncProfile);
            else
                vReturn = true;

            cmbSyncProfile.Items.Add("[Configuration manager...]");
            cmbSyncProfile.Text = syncProfile;

            return vReturn;
        }
                

        private void LoadSettings(string _profile)
        {
            RegistryKey regKeyAppRoot = Registry.CurrentUser.CreateSubKey(AppRootKey  + (_profile != null ? ('\\' + _profile) : "")  );

            if (regKeyAppRoot.GetValue("SyncOption") != null)
            {
                syncOption = (SyncOption)regKeyAppRoot.GetValue("SyncOption");
                SetSyncOption((int)syncOption);
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
            if (regKeyAppRoot.GetValue("SyncAppointmentsMonthsInPast") != null)
                pastMonthTextBox.Text = regKeyAppRoot.GetValue("SyncAppointmentsMonthsInPast") as string;
            if (regKeyAppRoot.GetValue("SyncAppointmentsMonthsInFuture") != null)
                futureMonthTextBox.Text = regKeyAppRoot.GetValue("SyncAppointmentsMonthsInFuture") as string;
            if (regKeyAppRoot.GetValue("SyncAppointmentsTimezone") != null)
                appointmentTimezonesComboBox.Text = regKeyAppRoot.GetValue("SyncAppointmentsTimezone") as string;
            if (regKeyAppRoot.GetValue("SyncAppointments") != null)
                btSyncAppointments.Checked = Convert.ToBoolean(regKeyAppRoot.GetValue("SyncAppointments"));
            if (regKeyAppRoot.GetValue("SyncNotes") != null)
                btSyncNotes.Checked = Convert.ToBoolean(regKeyAppRoot.GetValue("SyncNotes"));
            if (regKeyAppRoot.GetValue("SyncContacts") != null)
                btSyncContacts.Checked = Convert.ToBoolean(regKeyAppRoot.GetValue("SyncContacts"));
            if (regKeyAppRoot.GetValue("UseFileAs") != null)
                chkUseFileAs.Checked = Convert.ToBoolean(regKeyAppRoot.GetValue("UseFileAs"));

            LoadSettingsFolders(_profile);

            autoSyncCheckBox_CheckedChanged(null, null);
            btSyncContacts_CheckedChanged(null, null);
            btSyncNotes_CheckedChanged(null, null);

            _proxy.LoadSettings(_profile);
        }

        private void LoadSettingsFolders(string _profile)
        {

            RegistryKey regKeyAppRoot = Registry.CurrentUser.CreateSubKey(AppRootKey + (_profile != null ? ('\\' + _profile) : ""));

            if (regKeyAppRoot.GetValue("SyncContactsFolder") != null)
                contactFoldersComboBox.SelectedValue = regKeyAppRoot.GetValue("SyncContactsFolder") as string;
            if (regKeyAppRoot.GetValue("SyncNotesFolder") != null)
                noteFoldersComboBox.SelectedValue = regKeyAppRoot.GetValue("SyncNotesFolder") as string;
            if (regKeyAppRoot.GetValue("SyncAppointmentsFolder") != null)
                appointmentFoldersComboBox.SelectedValue = regKeyAppRoot.GetValue("SyncAppointmentsFolder") as string;
        }

		private void SaveSettings()
		{
            SaveSettings(cmbSyncProfile.Text);
		}

        private void SaveSettings(string profile)
        {
            if (!string.IsNullOrEmpty(profile))
            {
                syncProfile = cmbSyncProfile.Text;
                RegistryKey regKeyAppRoot = Registry.CurrentUser.CreateSubKey(AppRootKey + "\\" + profile);
                regKeyAppRoot.SetValue("SyncOption", (int)syncOption);
 
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
                regKeyAppRoot.SetValue("SyncAppointmentsMonthsInPast", pastMonthTextBox.Text);
                regKeyAppRoot.SetValue("SyncAppointmentsMonthsInFuture", futureMonthTextBox.Text);
                regKeyAppRoot.SetValue("SyncAppointmentsTimeZone", appointmentTimezonesComboBox.Text);
                regKeyAppRoot.SetValue("SyncAppointments", btSyncAppointments.Checked);
                regKeyAppRoot.SetValue("SyncNotes", btSyncNotes.Checked);
                regKeyAppRoot.SetValue("SyncContacts", btSyncContacts.Checked);
                regKeyAppRoot.SetValue("UseFileAs", chkUseFileAs.Checked);

                //if (btSyncContacts.Checked && contactFoldersComboBox.SelectedValue != null)
                //    regKeyAppRoot.SetValue("SyncContactsFolder", contactFoldersComboBox.SelectedValue.ToString());
                //if (btSyncNotes.Checked && noteFoldersComboBox.SelectedValue != null)
                //    regKeyAppRoot.SetValue("SyncNotesFolder", noteFoldersComboBox.SelectedValue.ToString());

                _proxy.SaveSettings(cmbSyncProfile.Text);
            }
        }


	    private bool ValidSyncFolders
	    {
            get
            {
                bool syncContactFolderIsValid = (contactFoldersComboBox.SelectedIndex >= 0
                                                 && contactFoldersComboBox.SelectedIndex < contactFoldersComboBox.Items.Count) || !btSyncContacts.Checked;
                bool syncNoteFolderIsValid = (noteFoldersComboBox.SelectedIndex >= 0 && noteFoldersComboBox.SelectedIndex < noteFoldersComboBox.Items.Count)
                                             || !btSyncNotes.Checked;
                bool syncAppointmentFolderIsValid = (appointmentFoldersComboBox.SelectedIndex >= 0 && appointmentFoldersComboBox.SelectedIndex < appointmentFoldersComboBox.Items.Count)
                                            || !btSyncAppointments.Checked;

                //ToDo: Coloring doesn'T Work for these combos
                setBgColor(contactFoldersComboBox, syncContactFolderIsValid);
                setBgColor(noteFoldersComboBox, syncNoteFolderIsValid);
                setBgColor(appointmentFoldersComboBox, syncAppointmentFolderIsValid);

                return syncContactFolderIsValid && syncNoteFolderIsValid && syncAppointmentFolderIsValid;
            }


	    }

	    private bool ValidCredentials
		{
			get
			{
				bool userNameIsValid = Regex.IsMatch(UserName.Text, @"^(?'id'[a-z0-9\'\%\._\+\-]+)@(?'domain'[a-z0-9\'\%\._\+\-]+)\.(?'ext'[a-z]{2,6})$", RegexOptions.IgnoreCase);
				bool passwordIsValid = !string.IsNullOrEmpty(Password.Text.Trim());
                bool syncProfileIsValid = (cmbSyncProfile.SelectedIndex > 0 && cmbSyncProfile.SelectedIndex < cmbSyncProfile.Items.Count-1);
               

				setBgColor(UserName, userNameIsValid);
				setBgColor(Password, passwordIsValid);
                setBgColor(cmbSyncProfile, syncProfileIsValid);
               


                if (!userNameIsValid)
                    toolTip.SetToolTip(UserName, "User is of wrong format, should be full Google Mail address, e.g. user@googelmail.com");
                else
                    toolTip.SetToolTip(UserName, String.Empty);
                if (!passwordIsValid)
                    toolTip.SetToolTip(Password, "Password is empty, please provide your Google Mail password");
                else
                    toolTip.SetToolTip(Password, String.Empty);               
                                             

                return userNameIsValid && passwordIsValid && syncProfileIsValid;
			}
		}

		private void setBgColor(Control box, bool isValid)
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
                    //return;
                    throw new Exception("Gmail Credentials are incomplete or incorrect!");

                fillSyncFolderItems();

                if (!ValidSyncFolders)
                    throw new Exception("At least one sync folder is not selected or invalid!");


                //IconTimerSwitch(true);
				ThreadStart starter = new ThreadStart(Sync_ThreadStarter);
				syncThread = new Thread(starter);
                syncThread.Start();

				// wait for thread to start
                for (int i = 0; !syncThread.IsAlive && i < 10; i++)
                    Thread.Sleep(1000);//DoNothing, until the thread was started, but only wait maximum 10 seconds
                
			}
			catch (Exception ex)
			{
                TimerSwitch(false);
                ShowForm();
				ErrorHandler.Handle(ex);
			}
		}
     
        [STAThread]
		private void Sync_ThreadStarter()
		{
            //==>Instead of lock, use Interlocked to exit the code, if already another thread is calling the same
            bool won = false;

            try
            {

                won = Interlocked.CompareExchange(ref executing, 1, 0) == 0;
                if (won)
                {

                    TimerSwitch(false);

                    //if the contacts or notes folder has changed ==> Reset matches (to not delete contacts or notes on the one or other side)                
                    RegistryKey regKeyAppRoot = Registry.CurrentUser.CreateSubKey(AppRootKey + "\\" + syncProfile);
                    string oldSyncContactsFolder = regKeyAppRoot.GetValue("SyncContactsFolder") as string;
                    string oldSyncNotesFolder = regKeyAppRoot.GetValue("SyncNotesFolder") as string;
                    string oldSyncAppointmentsFolder = regKeyAppRoot.GetValue("SyncAppointmentsFolder") as string;

                    //only reset notes if NotesFolder changed and reset contacts if ContactsFolder changed
                    bool syncContacts = !string.IsNullOrEmpty(oldSyncContactsFolder) && !oldSyncContactsFolder.Equals(this.syncContactsFolder) && btSyncContacts.Checked;
                    bool syncNotes = !string.IsNullOrEmpty(oldSyncNotesFolder) && !oldSyncNotesFolder.Equals(this.syncNotesFolder) && btSyncNotes.Checked;
                    bool syncAppointments = !string.IsNullOrEmpty(oldSyncAppointmentsFolder) && !oldSyncAppointmentsFolder.Equals(this.syncAppointmentsFolder) && btSyncAppointments.Checked;
                    if (syncContacts || syncNotes || syncAppointments)
                        ResetMatches(syncContacts, syncNotes, syncAppointments);

                    //Then save the Contacts and Notes Folders used at last sync
                    if (btSyncContacts.Checked)
                        regKeyAppRoot.SetValue("SyncContactsFolder", this.syncContactsFolder);
                    if (btSyncNotes.Checked)
                        regKeyAppRoot.SetValue("SyncNotesFolder", this.syncNotesFolder);
                    if (btSyncAppointments.Checked)
                        regKeyAppRoot.SetValue("SyncAppointmentsFolder", this.syncAppointmentsFolder);

                    SetLastSyncText("Syncing...");
                    notifyIcon.Text = Application.ProductName + "\nSyncing...";                    
                    //System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SettingsForm));
                    //notifyIcon.Icon = ((System.Drawing.Icon)(resources.GetObject("notifyIcon.Icon")));                    
                    IconTimerSwitch(true);

                    SetFormEnabled(false);

                    if (sync == null)
                    {
                        sync = new Syncronizer();
                        sync.DuplicatesFound += new Syncronizer.DuplicatesFoundHandler(OnDuplicatesFound);
                        sync.ErrorEncountered += new Syncronizer.ErrorNotificationHandler(OnErrorEncountered);
                    }

                    Logger.ClearLog();
                    SetSyncConsoleText("");
                    Logger.Log("Sync started (" + syncProfile + ").", EventType.Information);
                    //SetSyncConsoleText(Logger.GetText());
                    sync.SyncProfile = syncProfile;
                    Syncronizer.SyncContactsFolder = this.syncContactsFolder;
                    Syncronizer.SyncNotesFolder = this.syncNotesFolder;
                    Syncronizer.SyncAppointmentsFolder = this.syncAppointmentsFolder;
                    Syncronizer.MonthsInPast = Convert.ToUInt16(this.pastMonthTextBox.Text);
                    Syncronizer.MonthsInFuture = Convert.ToUInt16(this.futureMonthTextBox.Text);
                    Syncronizer.Timezone = this.Timezone;

                    sync.SyncOption = syncOption;
                    sync.SyncDelete = btSyncDelete.Checked;
                    sync.PromptDelete = btPromptDelete.Checked && btSyncDelete.Checked;
                    sync.UseFileAs = chkUseFileAs.Checked;
                    sync.SyncNotes = btSyncNotes.Checked;
                    sync.SyncContacts = btSyncContacts.Checked;
                    sync.SyncAppointments = btSyncAppointments.Checked;                    

                    if (!sync.SyncContacts && !sync.SyncNotes && !sync.SyncAppointments)
                    {
                        SetLastSyncText("Sync failed.");
                        notifyIcon.Text = Application.ProductName + "\nSync failed";

                        string messageText = "Neither notes nor contacts  nor appointments are switched on for syncing. Please choose at least one option. Sync aborted!";
                        Logger.Log(messageText, EventType.Error);
                        ShowForm();
                        ShowBalloonToolTip("Error", messageText, ToolTipIcon.Error, 5000, true);
                        return;
                    }


                    sync.LoginToGoogle(UserName.Text, Password.Text);
                    sync.LoginToOutlook();

                    sync.Sync();

                    lastSync = DateTime.Now;
                    SetLastSyncText("Last synced at " + lastSync.ToString());

                    string message = string.Format("Sync complete.\r\n Synced:  {1} out of {0}.\r\n Deleted:  {2}.\r\n Skipped: {3}.\r\n Errors:    {4}.", sync.TotalCount, sync.SyncedCount, sync.DeletedCount, sync.SkippedCount, sync.ErrorCount);
                    Logger.Log(message, EventType.Information);
                    if (reportSyncResultCheckBox.Checked)
                    {
                        /*
                        notifyIcon.BalloonTipTitle = Application.ProductName;
                        notifyIcon.BalloonTipText = string.Format("{0}. {1}", DateTime.Now, message);
                        */
                        ToolTipIcon icon;
                        if (sync.ErrorCount > 0)
                            icon = ToolTipIcon.Error;
                        else if (sync.SkippedCount > 0)
                            icon = ToolTipIcon.Warning;
                        else
                            icon = ToolTipIcon.Info;
                        /*notifyIcon.ShowBalloonTip(5000);
                        */
                        ShowBalloonToolTip(Application.ProductName,
                            string.Format("{0}. {1}", DateTime.Now, message),
                            icon,
                            5000, false);

                    }
                    string toolTip = string.Format("{0}\nLast sync: {1}", Application.ProductName, DateTime.Now.ToString("dd.MM. HH:mm"));
                    if (sync.ErrorCount + sync.SkippedCount > 0)
                        toolTip += string.Format("\nWarnings: {0}.", sync.ErrorCount + sync.SkippedCount);
                    if (toolTip.Length >= 64)
                        toolTip = toolTip.Substring(0, 63);
                    notifyIcon.Text = toolTip;
                }
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
                    ShowBalloonToolTip("Error", message, ToolTipIcon.Error, 5000, true);
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
                ShowBalloonToolTip("Error", message, ToolTipIcon.Error, 5000, true);

            }
            catch (Exception ex)
            {
                SetLastSyncText("Sync failed.");
                notifyIcon.Text = Application.ProductName + "\nSync failed";

                if (ex is COMException)
                {
                    string message = "Outlook exception, please assure that Outlook is running and not closed when syncing";
                    Logger.Log(message + ": " + ex.Message + "\r\n" + ex.StackTrace, EventType.Warning);
                    ShowBalloonToolTip("Error", message, ToolTipIcon.Error, 5000, true);
                }
                else
                {
                    ErrorHandler.Handle(ex);
                }
            }							
			finally
			{
                if (won)
                {
                    Interlocked.Exchange(ref executing, 0);

                    lastSync = DateTime.Now;
                    TimerSwitch(true);
                    SetFormEnabled(true);
                    if (sync != null)
                    {
                        sync.LogoffOutlook();
                        sync.LogoffGoogle();
                        sync = null;
                    }

                    IconTimerSwitch(false);
                }
			}
		}

        public void ShowBalloonToolTip(string title, string message, ToolTipIcon icon, int timeout, bool error)
        {
            //if user is active on workstation
            if(boolShowBalloonTip)
            {
                notifyIcon.BalloonTipTitle = title;
			    notifyIcon.BalloonTipText = message;
			    notifyIcon.BalloonTipIcon = icon;
			    notifyIcon.ShowBalloonTip(timeout);
            }

            string iconText = title + ": " + message;
            if (!string.IsNullOrEmpty(iconText))
                notifyIcon.Text = (iconText).Substring(0, iconText.Length >=63? 63:iconText.Length);

            if (error) 
                notifyIcon.Icon = this.IconError;
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
            ShowBalloonToolTip(title,message,ToolTipIcon.Error,5000, true);
			/*notifyIcon.BalloonTipTitle = title;
			notifyIcon.BalloonTipText = message;
			notifyIcon.BalloonTipIcon = ToolTipIcon.Error;
			notifyIcon.ShowBalloonTip(5000);*/
		}

		void OnDuplicatesFound(string title, string message)
		{
            Logger.Log(message, EventType.Warning);
            ShowBalloonToolTip(title,message,ToolTipIcon.Warning,5000, true);
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
                cancelButton.Enabled = !enabled;
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
                //If PC resumes or unlocks or is started, give him 5 minutes to recover everything before the sync starts
                if (lastSync <= DateTime.Now.AddSeconds(300) - new TimeSpan(0, (int)autoSyncInterval.Value, 0))
                    lastSync = DateTime.Now.AddSeconds(300) - new TimeSpan(0, (int)autoSyncInterval.Value, 0);
				autoSyncInterval.Enabled = autoSyncCheckBox.Checked && value;
				syncTimer.Enabled = autoSyncCheckBox.Checked && value;
				nextSyncLabel.Visible = autoSyncCheckBox.Checked &&  value;          				
			}
		}

        

        protected override void WndProc(ref System.Windows.Forms.Message m)
		{
            //Logger.Log(m.Msg, EventType.Information);
            switch(m.Msg) 
            {
                //System shutdown
                case NativeMethods.WM_QUERYENDSESSION:
                    requestClose = true;
                    break;
                /*case NativeMethods.WM_WTSSESSION_CHANGE:
                    {
                        int value = m.WParam.ToInt32();
                        //User Session locked
                        if (value == NativeMethods.WTS_SESSION_LOCK)
                        {
                            Console.WriteLine("Session Lock",EventType.Information);
                            //OnSessionLock();
                            boolShowBalloonTip = false; // Do something when locked
                        }
                        //User Session unlocked
                        else if (value == NativeMethods.WTS_SESSION_UNLOCK)
                        {
                            Console.WriteLine("Session Unlock", EventType.Information);
                            //OnSessionUnlock();
                            boolShowBalloonTip = true; // Do something when unlocked
                            TimerSwitch(true);
                        }
                     break;
                    }
                
                
                case NativeMethods.WM_POWERBROADCAST:
                    {
                        if (m.WParam.ToInt32() == NativeMethods.PBT_APMRESUMEAUTOMATIC ||
                            m.WParam.ToInt32() == NativeMethods.PBT_APMRESUMECRITICAL ||
                            m.WParam.ToInt32() == NativeMethods.PBT_APMRESUMESTANDBY ||
                            m.WParam.ToInt32() == NativeMethods.PBT_APMRESUMESUSPEND ||
                            m.WParam.ToInt32() == NativeMethods.PBT_APMQUERYSTANDBYFAILED ||
                            m.WParam.ToInt32() == NativeMethods.PBT_APMQUERYSTANDBYFAILED)
                        {                            
                            TimerSwitch(true);
                        }
                        else if (m.WParam.ToInt32() == NativeMethods.PBT_APMSUSPEND ||
                                 m.WParam.ToInt32() == NativeMethods.PBT_APMSTANDBY ||
                                 m.WParam.ToInt32() == NativeMethods.PBT_APMQUERYSTANDBY ||
                                 m.WParam.ToInt32() == NativeMethods.PBT_APMQUERYSUSPEND)
                        {
                            TimerSwitch(false);
                        }
                            

                        break;
                    }*/
                default:
                    break;
            }
            //Show Window from Tray
            if (m.Msg == NativeMethods.WM_GCSM_SHOWME)
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
				if (sync != null)
					sync.LogoffOutlook();


                Logger.Log("Closed application.", EventType.Information);
                Logger.Close();

				SaveSettings();
                
                //unregister event handler
                SystemEvents.SessionSwitch -= SystemEvents_SessionSwitch;
                SystemEvents.PowerModeChanged -= SystemEvents_PowerModeSwitch;
				
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
                TimerSwitch(false);
                ShowForm();
				ErrorHandler.Handle(ex);
			}
		}
		private void SetSyncOption(int index)
		{
			syncOption = (SyncOption)index;
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
            //if (WindowState == FormWindowState.Normal)
            //    HideForm();
            //else
				ShowForm();
		}

		private void autoSyncCheckBox_CheckedChanged(object sender, EventArgs e)
		{
            lastSync = DateTime.Now.AddSeconds(300) - new TimeSpan(0, (int)autoSyncInterval.Value, 0);
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
                this.cancelButton.Enabled = false; //Cancel is only working for sync currently, not for reset
                ResetMatches(btSyncContacts.Checked,btSyncNotes.Checked, btSyncAppointments.Checked);
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
                if (sync != null)
                {
                    sync.LogoffOutlook();
                    sync.LogoffGoogle();
                    sync = null;
                }
			}
		}

        private void ResetMatches(bool syncContacts, bool syncNotes, bool syncAppointments)
        {
            TimerSwitch(false);

            SetLastSyncText("Resetting matches...");
            notifyIcon.Text = Application.ProductName + "\nResetting matches...";

            fillSyncFolderItems();

            SetFormEnabled(false);            

            if (sync == null)
            {
                sync = new Syncronizer();
            }

            Logger.ClearLog();
            SetSyncConsoleText("");
            Logger.Log("Reset Matches started  (" + syncProfile + ").", EventType.Information);

            sync.SyncNotes = syncNotes;
            sync.SyncContacts = syncContacts;
            sync.SyncAppointments = syncAppointments;

            Syncronizer.SyncContactsFolder = syncContactsFolder;
            Syncronizer.SyncNotesFolder = syncNotesFolder;
            Syncronizer.SyncAppointmentsFolder = syncAppointmentsFolder;
            sync.SyncProfile = syncProfile;

            sync.LoginToGoogle(UserName.Text, Password.Text);
            sync.LoginToOutlook();

           
            if (sync.SyncAppointments)
            {
                bool deleteOutlookAppointments = false;
                bool deleteGoogleAppointments = false;                
                switch (MessageBox.Show(this, "Do you want to delete all Outlook Calendar entries?", Application.ProductName, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2))
                {
                    case DialogResult.Yes: deleteOutlookAppointments = true; break;
                    case DialogResult.No: deleteOutlookAppointments = false; break;
                    default: return;
                }
                switch (MessageBox.Show(this, "Do you want to delete all Google Calendar entries?", Application.ProductName, MessageBoxButtons.YesNoCancel,MessageBoxIcon.Question,MessageBoxDefaultButton.Button2 ))
                {
                    case DialogResult.Yes: deleteGoogleAppointments = true; break;
                    case DialogResult.No: deleteGoogleAppointments = false; break;
                    default: return;
                }
                
                sync.LoadAppointments();
                sync.ResetAppointmentMatches(deleteOutlookAppointments, deleteGoogleAppointments);
            }

            if (sync.SyncContacts)
            {
                sync.LoadContacts();
                sync.ResetContactMatches();
            }


            if (sync.SyncNotes)
            {
                sync.LoadNotes();
                sync.ResetNoteMatches();
            }

            



            lastSync = DateTime.Now;
            SetLastSyncText("Matches reset at " + lastSync.ToString());
            Logger.Log("Matches reset.", EventType.Information);
        }

        private delegate void InvokeCallback();

        public delegate DialogResult InvokeConflict(ConflictResolverForm conflictResolverForm); 

        public DialogResult ShowConflictDialog(ConflictResolverForm conflictResolverForm)
        {
            if (this.InvokeRequired)
            {
                return (DialogResult) Invoke(new InvokeConflict(ShowConflictDialog), new object[] {conflictResolverForm});
            }
            else
            {
                DialogResult res = conflictResolverForm.ShowDialog(this);
                                
                notifyIcon.Icon = this.Icon0;

                return res;

            }
        }

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
                fillSyncFolderItems();
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
                string.IsNullOrEmpty(cmbSyncProfile.Text) /*||
                string.IsNullOrEmpty(contactFoldersComboBox.Text)*/ )
			{
				// this is the first load, show form
				ShowForm();
				UserName.Focus();
                ShowBalloonToolTip(Application.ProductName,
                        "Application started and visible in your PC's system tray, click on this balloon or the icon below to open the settings form and enter your Google credentials there.",
                        ToolTipIcon.Info,
                        5000, false);
			}
			else
				HideForm();
		}

		private void runAtStartupCheckBox_CheckedChanged(object sender, EventArgs e)
		{
            string regKey = @"Software\Microsoft\Windows\CurrentVersion\Run";
            try
            {
                RegistryKey regKeyAppRoot = Registry.CurrentUser.CreateSubKey(regKey);

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
            catch (Exception ex)
            {
                TimerSwitch(false);
                ShowForm();
                ErrorHandler.Handle(new Exception("Error saving 'Run program at startup' settings into Registry key '" + regKey + "'",ex));
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
			syncButton.Enabled = ValidCredentials && ValidSyncFolders;
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
            if (_proxy != null) _proxy.ShowDialog(this);
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
            if (!btSyncContacts.Checked && !btSyncNotes.Checked && !btSyncAppointments.Checked)
            {
                MessageBox.Show("Neither notes nor contacts nor appointments are switched on for syncing. Please choose at least one option (automatically switched on notes for syncing now).", "No sync switched on");
                btSyncNotes.Checked = true;
            }
            contactFoldersComboBox.Visible = btSyncContacts.Checked;
        }

        private void btSyncNotes_CheckedChanged(object sender, EventArgs e)
        {
            if (!btSyncContacts.Checked && !btSyncNotes.Checked && !btSyncAppointments.Checked)
            {
                MessageBox.Show("Neither notes nor contacts nor appointments are switched on for syncing. Please choose at least one option (automatically switched on appointments for syncing now).", "No sync switched on");
                btSyncAppointments.Checked = true;
            }
            noteFoldersComboBox.Visible = btSyncNotes.Checked;
        }

        private void btSyncAppointments_CheckedChanged(object sender, EventArgs e)
        {
            if (!btSyncContacts.Checked && !btSyncNotes.Checked && !btSyncAppointments.Checked)
            {
                MessageBox.Show("Neither notes nor contacts nor appointments are switched on for syncing. Please choose at least one option (automatically switched on contacts for syncing now).", "No sync switched on");
                btSyncContacts.Checked = true;
            }
            appointmentFoldersComboBox.Visible = btSyncAppointments.Checked;
            pastMonthTextBox.Visible = futureMonthTextBox.Visible = appointmentTimezonesComboBox.Visible = btSyncAppointments.Checked;
        }
    	
        private void cmbSyncProfile_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox comboBox = (ComboBox)sender;

            if ((0 == comboBox.SelectedIndex) || (comboBox.SelectedIndex == (comboBox.Items.Count - 1))) {
                ConfigurationManagerForm _configs = new ConfigurationManagerForm();

                if (0 == comboBox.SelectedIndex && _configs != null)
                {
                    syncProfile = _configs.AddProfile();
                    ClearSettings();
                }
                
                if (comboBox.SelectedIndex == (comboBox.Items.Count - 1) && _configs != null)
                    _configs.ShowDialog(this);

                fillSyncProfileItems();

                comboBox.Text = syncProfile;
                SaveSettings();
            }
            if (comboBox.SelectedIndex < 0)
                MessageBox.Show("Please select Sync Profile.", "No sync switched on");
            else
            {
                //ClearSettings();
                LoadSettings(comboBox.Text);
                syncProfile = comboBox.Text;
            }

            ValidateSyncButton();
        }

        private void contacFoldersComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            string message = "Select the Outlook Contacts folder you want to sync";
            if ((sender as ComboBox).SelectedIndex >= 0 && (sender as ComboBox).SelectedIndex < (sender as ComboBox).Items.Count && (sender as ComboBox).SelectedItem is OutlookFolder)
            {
                syncContactsFolder = (sender as ComboBox).SelectedValue.ToString();
                toolTip.SetToolTip((sender as ComboBox), message + ":\r\n" + ((OutlookFolder)(sender as ComboBox).SelectedItem).DisplayName);
            }
            else
            {
                syncContactsFolder = "";
                toolTip.SetToolTip((sender as ComboBox), message);
            }
            ValidateSyncButton();

            
        }

        private void noteFoldersComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            string message = "Select the Outlook Notes folder you want to sync";
            if ((sender as ComboBox).SelectedIndex >= 0 && (sender as ComboBox).SelectedIndex < (sender as ComboBox).Items.Count && (sender as ComboBox).SelectedItem is OutlookFolder)
            {
                syncNotesFolder = (sender as ComboBox).SelectedValue.ToString();
                toolTip.SetToolTip((sender as ComboBox), message + ":\r\n" + ((OutlookFolder)(sender as ComboBox).SelectedItem).DisplayName);
            }
            else
            {
                syncNotesFolder = "";
                toolTip.SetToolTip((sender as ComboBox), message);
            }

            ValidateSyncButton();

            
        }

        private void appointmentFoldersComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            string message = "Select the Outlook Appointments folder you want to sync";
            if ((sender as ComboBox).SelectedIndex >= 0 && (sender as ComboBox).SelectedIndex < (sender as ComboBox).Items.Count && (sender as ComboBox).SelectedItem is OutlookFolder)
            {
                syncAppointmentsFolder = (sender as ComboBox).SelectedValue.ToString();
                toolTip.SetToolTip((sender as ComboBox), message + ":\r\n" + ((OutlookFolder)(sender as ComboBox).SelectedItem).DisplayName);
            }
            else
            {
                syncAppointmentsFolder = "";
                toolTip.SetToolTip((sender as ComboBox), message);
            }

            ValidateSyncButton();


        }

        private void btSyncDelete_CheckedChanged(object sender, EventArgs e)
        {
            btPromptDelete.Visible = btSyncDelete.Checked;
            btPromptDelete.Checked = btSyncDelete.Checked;
        }

        private void pictureBoxExit_Click(object sender, EventArgs e)
        {
            if (DialogResult.Yes == MessageBox.Show("Do you really want to exit " + Application.ProductName + "? This will also stop the service performing automatic synchronizaton in the background. If you only want to hide the settings form, use the 'Hide' Button instead.", "Exit " + Application.ProductName, MessageBoxButtons.YesNo, MessageBoxIcon.Question))
            {
                CancelButton_Click(sender, EventArgs.Empty); //Close running thread
                requestClose = true;
                Close();
            }
        }

        private void SystemEvents_PowerModeSwitch(Object sender, PowerModeChangedEventArgs e)
        {
            if (e.Mode == PowerModes.Suspend)
            {
                TimerSwitch(false);
            }
            else if (e.Mode == PowerModes.Resume)
            {
                TimerSwitch(true);
            }
        }

        private void SystemEvents_SessionSwitch(object sender, SessionSwitchEventArgs e)
        {
            if (e.Reason == SessionSwitchReason.SessionLock)
            {
                boolShowBalloonTip = false;
            }
            else if (e.Reason == SessionSwitchReason.SessionUnlock)
            {
                boolShowBalloonTip = true;
                TimerSwitch(true);
            }
        }

        private void autoSyncInterval_ValueChanged(object sender, EventArgs e)
        {
            TimerSwitch(true);
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            KillSyncThread();
        }

        [System.Security.Permissions.SecurityPermission(System.Security.Permissions.SecurityAction.Demand, ControlThread = true)]
        private void KillSyncThread()
        {
            if (syncThread != null && syncThread.IsAlive)
                syncThread.Abort();
        }

        #region syncing icon
        public void IconTimerSwitch(bool value)
        {
            if (this.InvokeRequired)
            {
                SwitchHandler h = new SwitchHandler(IconTimerSwitch);
                this.Invoke(h, new object[] { value });
            }
            else
            {      
                if (value) //Reset Icon to default icon as starting point for the syncing icon
                    notifyIcon.Icon = this.Icon0;
                iconTimer.Enabled = value;
            }
        }

        private void iconTimer_Tick(object sender, EventArgs e)
        {
            showNextIcon();
        }

        private void showNextIcon()
        {
            if (this.InvokeRequired)
            {
                IconHandler h = new IconHandler(showNextIcon);
                this.Invoke(h, new object[] {});
            }
            else
                this.notifyIcon.Icon = GetNextIcon(this.notifyIcon.Icon); ;
        }

        
        
        private Icon GetNextIcon(Icon currentIcon)
        {
            if (currentIcon == IconError) //Don't change the icon anymore, once an error occurred
                return IconError;
            if (currentIcon == Icon30)
                return Icon60;
            else if (currentIcon == Icon60)
                return Icon90;
            else if (currentIcon == Icon90)
                return Icon120;
            else if (currentIcon == Icon120)
                return Icon150;
            else if (currentIcon == Icon150)
                return Icon180;
            else if (currentIcon == Icon180)
                return Icon210;
            else if (currentIcon == Icon210)
                return Icon240;
            else if (currentIcon == Icon240)
                return Icon270;
            else if (currentIcon == Icon270)
                return Icon300;
            else if (currentIcon == Icon300)
                return Icon330;
            else if (currentIcon == Icon330)
                return Icon0;
            else
                return Icon30;
        }
        #endregion

        private void futureMonthTextBox_Validating(object sender, CancelEventArgs e)
        {
            ushort value;
            if (!ushort.TryParse(futureMonthTextBox.Text, out value))
            {
                MessageBox.Show("only positive integer numbers or 0 (i.e. all) allowed");
                futureMonthTextBox.Text = "0";
                e.Cancel = true;
            }

        }

        private void pastMonthTextBox_Validating(object sender, CancelEventArgs e)
        {
            ushort value;
            if (!ushort.TryParse(pastMonthTextBox.Text, out value))
            {
                MessageBox.Show("only positive integer numbers or 0 (i.e. all) allowed");
                pastMonthTextBox.Text = "1";
                e.Cancel = true;
            }
        }

        private void appointmentTimezonesComboBox_TextChanged(object sender, EventArgs e)
        {
            this.Timezone = appointmentTimezonesComboBox.Text;
        }


    }
}