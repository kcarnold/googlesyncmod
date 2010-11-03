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

namespace WebGear.GoogleContactsSync
{
    internal partial class SettingsForm : Form
    {
        internal Syncronizer _sync;
        private SyncOption _syncOption;
        private DateTime lastSync;
        private bool requestClose = false;

        delegate void TextHandler(string text);
        delegate void SwitchHandler(bool value);

        public SettingsForm()
        {
            InitializeComponent();

            PopulateSyncOptionBox();

            LoadSettings();

            lastSync = DateTime.Now.AddSeconds(15) - new TimeSpan(0, (int)autoSyncInterval.Value, 0);
            lastSyncLabel.Text = "Not synced";

            ValidateSyncButton();
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

        private void LoadSettings()
        {
            // default
            SetSyncOption(0);

            // load
            RegistryKey regKeyAppRoot = Registry.CurrentUser.CreateSubKey(@"Software\Webgear\GOContactSync");
            if (regKeyAppRoot.GetValue("SyncOption") != null)
            {
                _syncOption = (SyncOption)regKeyAppRoot.GetValue("SyncOption");
                SetSyncOption((int)_syncOption);
            }
            if (regKeyAppRoot.GetValue("SyncProfile") != null)
                tbSyncProfile.Text = (string)regKeyAppRoot.GetValue("SyncProfile");
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

            autoSyncCheckBox_CheckedChanged(null, null);
        }
        private void SaveSettings()
        {
            RegistryKey regKeyAppRoot = Registry.CurrentUser.CreateSubKey(@"Software\Webgear\GOContactSync");
            regKeyAppRoot.SetValue("SyncOption", (int)_syncOption);
            if (!string.IsNullOrEmpty(tbSyncProfile.Text))
                regKeyAppRoot.SetValue("SyncProfile", tbSyncProfile.Text);
            if (!string.IsNullOrEmpty(UserName.Text))
            {
                regKeyAppRoot.SetValue("Username", UserName.Text);
                if (!string.IsNullOrEmpty(Password.Text))
                    regKeyAppRoot.SetValue("Password", Encryption.EncryptPassword(UserName.Text, Password.Text));
            }
            regKeyAppRoot.SetValue("AutoSync", autoSyncCheckBox.Checked.ToString());
            regKeyAppRoot.SetValue("AutoSyncInterval", autoSyncInterval.Value.ToString());
            regKeyAppRoot.SetValue("AutoStart", runAtStartupCheckBox.Checked);
        }

        private bool ValidCredentials
        {
            get
            {
                bool userNameIsValid = Regex.IsMatch(UserName.Text, @"^(?'id'[a-z0-9\'\%\._\+\-]+)@(?'domain'[a-z0-9\'\%\._\+\-]+)\.(?'ext'[a-z]{2,6})$", RegexOptions.IgnoreCase);
                bool passwordIsValid = Password.Text.Length != 0;
                bool syncProfileNameIsValid = tbSyncProfile.Text.Length != 0;

                setBgColor(UserName, userNameIsValid);
                setBgColor(Password, passwordIsValid);
                setBgColor(tbSyncProfile, syncProfileNameIsValid);
                return userNameIsValid && passwordIsValid && syncProfileNameIsValid;
            }
        }
        private void setBgColor(TextBox box, bool isValid)
        {
            if (!isValid)
                box.BackColor = Color.LightPink;
            else
                box.BackColor = Color.LightGreen;

        }


        private void button4_Click(object sender, EventArgs e)
        {
            Sync();
        }
        private void Sync()
        {
            if (!ValidCredentials)
                return;

            //Sync_ThreadStarter();

            ThreadStart starter = new ThreadStart(Sync_ThreadStarter);
            Thread thread = new Thread(starter);
            thread.Start();

            // wait for thread to start
            while (!thread.IsAlive)
                Thread.Sleep(1);
        }

        private void Sync_ThreadStarter()
        {
            TimerSwitch(false);
            SetLastSyncText("Syncing...");
            SetFormEnabled(false);

            if (_sync == null)
            {
                _sync = new Syncronizer();
                _sync.Logger = new Logger();
                _sync.DuplicatesFound += new Syncronizer.NotificationHandler(_sync_DuplicatesFound);
                _sync.ErrorEncountered += new Syncronizer.NotificationHandler(_sync_ErrorEncountered);
                _sync.Logger.LogUpdated += new Logger.LogUpdatedHandler(Logger_LogUpdated);
            }

            _sync.Logger.ClearLog();
            SetSyncConsoleText("");
            _sync.Logger.Log("Sync started.", EventType.Information);
            //SetSyncConsoleText(_sync.Logger.GetText());
            _sync.SyncProfile = tbSyncProfile.Text;
            _sync.SyncOption = _syncOption;

            try
            {
                _sync.LoginToGoogle(UserName.Text, Password.Text);
                _sync.LoginToOutlook();

                _sync.Sync();

                SetLastSyncText("Last synced at " + lastSync.ToString());
                _sync.Logger.Log("Sync complete.", EventType.Information);
                //SetSyncConsoleText(_sync.Logger.GetText());

                notifyIcon.BalloonTipTitle = "Complete";
                notifyIcon.BalloonTipText = string.Format("{0}. Sync complete.\n Synced: {2} out of {1}.\n Deleted: {3}.", DateTime.Now, _sync.TotalCount, _sync.SyncedCount, _sync.DeletedCount);
                notifyIcon.BalloonTipIcon = ToolTipIcon.Info;
                notifyIcon.ShowBalloonTip(5000);
            }
            catch (Exception ex)
            {
                _sync.Logger.Log(ex.Message, EventType.Error);
                //AppendSyncConsoleText(_sync.Logger.GetText());
                _sync.Logger.Log("Sync failed.", EventType.Error);

                notifyIcon.BalloonTipTitle = "Error";
                notifyIcon.BalloonTipText = ex.Message;
                notifyIcon.BalloonTipIcon = ToolTipIcon.Error;
                notifyIcon.ShowBalloonTip(5000);

            }
            lastSync = DateTime.Now;
            TimerSwitch(true);
            SetFormEnabled(true);
        }

        void Logger_LogUpdated(string Message)
        {
            AppendSyncConsoleText(Message);
        }

        void _sync_ErrorEncountered(string title, string message, EventType eventType)
        {
            notifyIcon.BalloonTipTitle = title;
            notifyIcon.BalloonTipText = message;
            notifyIcon.BalloonTipIcon = ToolTipIcon.Error;
            notifyIcon.ShowBalloonTip(5000);
        }

        void _sync_DuplicatesFound(string title, string message, EventType eventType)
        {
            notifyIcon.BalloonTipTitle = title;
            notifyIcon.BalloonTipText = message;
            notifyIcon.BalloonTipIcon = ToolTipIcon.Warning;
            notifyIcon.ShowBalloonTip(5000);
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
                settingsGroupBox.Enabled = enabled;
                actionsTableLayout.Enabled = enabled;
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
                syncConsole.Text = text;
        }
        public void AppendSyncConsoleText(string text)
        {
            if (this.InvokeRequired)
            {
                TextHandler h = new TextHandler(AppendSyncConsoleText);
                this.Invoke(h, new object[] { text });
            }
            else
                syncConsole.Text += text;
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

                if (value)
                {
                    if (autoSyncCheckBox.Checked)
                    {
                        autoSyncInterval.Enabled = autoSyncCheckBox.Checked;
                        syncTimer.Enabled = autoSyncCheckBox.Checked;
                        nextSyncLabel.Visible = autoSyncCheckBox.Checked;
                    }
                }
                else
                {
                    autoSyncInterval.Enabled = value;
                    syncTimer.Enabled = value;
                    nextSyncLabel.Visible = value;
                }
            }
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
            if (_sync != null)
                _sync.LogoffOutlook();

            SaveSettings();

            notifyIcon.Dispose();
        }

        private void syncOptionBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            int index = syncOptionBox.SelectedIndex;
            if (index == -1)
                return;

            SetSyncOption(index);
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
            autoSyncInterval.Enabled = autoSyncCheckBox.Checked;
            syncTimer.Enabled = autoSyncCheckBox.Checked;
            nextSyncLabel.Visible = autoSyncCheckBox.Checked;
        }

        private void syncTimer_Tick(object sender, EventArgs e)
        {
            if (lastSync != null)
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
                    return;
                }
            }
            Sync();
        }

        private void resetMatchesButton_Click(object sender, EventArgs e)
        {
            if (_sync == null)
            {
                _sync = new Syncronizer();
                _sync.Logger = new Logger();
            }

            _sync.LoginToGoogle(UserName.Text, Password.Text);
            _sync.LoginToOutlook();

            _sync.Load();

            _sync.ResetMatches();

            AppendSyncConsoleText(DateTime.Now + " Contacts unlinked" + Environment.NewLine);
        }

        private void ShowForm()
        {
            Show();
            WindowState = FormWindowState.Normal;
        }
        private void HideForm()
        {
            WindowState = FormWindowState.Minimized;
            Hide();
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            ShowForm();
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
                string.IsNullOrEmpty(tbSyncProfile.Text))
            {
                // this is the first load, show form
                ShowForm();
                UserName.Focus(); 
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

        private void tbSyncProfile_TextChanged(object sender, EventArgs e)
        {
            ValidateSyncButton();
        }

        private void btSyncDelete_CheckedChanged(object sender, EventArgs e)
        {
            if (_sync == null)
            {
                _sync = new Syncronizer();
                _sync.Logger = new Logger();
            }

            _sync.SyncDelete = btSyncDelete.Checked;
        }
    }

    //internal class EventLogger : ILogger
    //{
    //    private TextBox _box;
    //    public EventLogger(TextBox box)
    //    {
    //        _box = box;
    //    }

    //    #region ILogger Members

    //    public void Log(string message, EventType eventType)
    //    {
    //        _box.Text += message;
    //    }

    //    #endregion
    //}
}