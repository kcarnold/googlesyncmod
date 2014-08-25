namespace GoContactSyncMod
{
    partial class SettingsForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SettingsForm));
            this.Password = new System.Windows.Forms.TextBox();
            this.UserName = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.syncButton = new System.Windows.Forms.Button();
            this.syncOptionBox = new System.Windows.Forms.CheckedListBox();
            this.notifyIcon = new System.Windows.Forms.NotifyIcon(this.components);
            this.systemTrayMenu = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.toolStripMenuItem4 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem3 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.toolStripMenuItem5 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem2 = new System.Windows.Forms.ToolStripMenuItem();
            this.autoSyncInterval = new System.Windows.Forms.NumericUpDown();
            this.autoSyncCheckBox = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.reportSyncResultCheckBox = new System.Windows.Forms.CheckBox();
            this.runAtStartupCheckBox = new System.Windows.Forms.CheckBox();
            this.nextSyncLabel = new System.Windows.Forms.Label();
            this.syncTimer = new System.Windows.Forms.Timer(this.components);
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.appointmentTimezonesComboBox = new System.Windows.Forms.ComboBox();
            this.futureMonthTextBox = new System.Windows.Forms.TextBox();
            this.pastMonthTextBox = new System.Windows.Forms.TextBox();
            this.btSyncAppointments = new System.Windows.Forms.CheckBox();
            this.appointmentFoldersComboBox = new System.Windows.Forms.ComboBox();
            this.btSyncNotes = new System.Windows.Forms.CheckBox();
            this.btSyncContacts = new System.Windows.Forms.CheckBox();
            this.btPromptDelete = new System.Windows.Forms.CheckBox();
            this.noteFoldersComboBox = new System.Windows.Forms.ComboBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.btSyncDelete = new System.Windows.Forms.CheckBox();
            this.cmbSyncProfile = new System.Windows.Forms.ComboBox();
            this.contactFoldersComboBox = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.lastSyncLabel = new System.Windows.Forms.Label();
            this.logGroupBox = new System.Windows.Forms.GroupBox();
            this.syncConsole = new System.Windows.Forms.TextBox();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.chkUseFileAs = new System.Windows.Forms.CheckBox();
            this.proxySettingsLinkLabel = new System.Windows.Forms.LinkLabel();
            this.toolTip = new System.Windows.Forms.ToolTip(this.components);
            this.resetMatchesLinkLabel = new System.Windows.Forms.LinkLabel();
            this.Donate = new System.Windows.Forms.PictureBox();
            this.pictureBoxExit = new System.Windows.Forms.PictureBox();
            this.settingsGroupBox = new System.Windows.Forms.GroupBox();
            this.actionsTableLayout = new System.Windows.Forms.TableLayoutPanel();
            this.cancelButton = new System.Windows.Forms.Button();
            this.hideButton = new System.Windows.Forms.Button();
            this.MainPanel = new System.Windows.Forms.Panel();
            this.MainSplitter = new System.Windows.Forms.Splitter();
            this.iconTimer = new System.Windows.Forms.Timer(this.components);
            this.systemTrayMenu.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.autoSyncInterval)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.logGroupBox.SuspendLayout();
            this.groupBox4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.Donate)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxExit)).BeginInit();
            this.settingsGroupBox.SuspendLayout();
            this.actionsTableLayout.SuspendLayout();
            this.MainPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // Password
            // 
            this.Password.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.Password.Location = new System.Drawing.Point(100, 44);
            this.Password.Name = "Password";
            this.Password.PasswordChar = '*';
            this.Password.Size = new System.Drawing.Size(489, 21);
            this.Password.TabIndex = 3;
            this.toolTip.SetToolTip(this.Password, "Type in your Google Mail Password");
            this.Password.TextChanged += new System.EventHandler(this.Password_TextChanged);
            // 
            // UserName
            // 
            this.UserName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.UserName.Location = new System.Drawing.Point(100, 18);
            this.UserName.Name = "UserName";
            this.UserName.Size = new System.Drawing.Size(489, 21);
            this.UserName.TabIndex = 1;
            this.toolTip.SetToolTip(this.UserName, "Type in your Google Mail User Name (full name)");
            this.UserName.TextChanged += new System.EventHandler(this.UserName_TextChanged);
            // 
            // label3
            // 
            this.label3.Location = new System.Drawing.Point(7, 44);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(75, 16);
            this.label3.TabIndex = 2;
            this.label3.Text = "&Password:";
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(7, 21);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(75, 17);
            this.label2.TabIndex = 0;
            this.label2.Text = "&User:";
            // 
            // syncButton
            // 
            this.syncButton.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.syncButton.Location = new System.Drawing.Point(3, 3);
            this.syncButton.Name = "syncButton";
            this.syncButton.Size = new System.Drawing.Size(79, 25);
            this.syncButton.TabIndex = 1;
            this.syncButton.Text = "S&ync";
            this.syncButton.UseVisualStyleBackColor = true;
            this.syncButton.Click += new System.EventHandler(this.syncButton_Click);
            // 
            // syncOptionBox
            // 
            this.syncOptionBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.syncOptionBox.BackColor = System.Drawing.SystemColors.Control;
            this.syncOptionBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.syncOptionBox.CheckOnClick = true;
            this.syncOptionBox.FormattingEnabled = true;
            this.syncOptionBox.IntegralHeight = false;
            this.syncOptionBox.Location = new System.Drawing.Point(7, 163);
            this.syncOptionBox.Name = "syncOptionBox";
            this.syncOptionBox.Size = new System.Drawing.Size(583, 110);
            this.syncOptionBox.TabIndex = 4;
            this.toolTip.SetToolTip(this.syncOptionBox, resources.GetString("syncOptionBox.ToolTip"));
            this.syncOptionBox.SelectedIndexChanged += new System.EventHandler(this.syncOptionBox_SelectedIndexChanged);
            // 
            // notifyIcon
            // 
            this.notifyIcon.BalloonTipIcon = System.Windows.Forms.ToolTipIcon.Warning;
            this.notifyIcon.ContextMenuStrip = this.systemTrayMenu;
            this.notifyIcon.Icon = global::GoContactSyncMod.Properties.Resources.sync;
            this.notifyIcon.Text = "GO Contact Sync Mod";
            this.notifyIcon.Visible = true;
            this.notifyIcon.BalloonTipClicked += new System.EventHandler(this.toolStripMenuItem1_Click);
            this.notifyIcon.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.notifyIcon_MouseDoubleClick);
            // 
            // systemTrayMenu
            // 
            this.systemTrayMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItem4,
            this.toolStripSeparator2,
            this.toolStripMenuItem1,
            this.toolStripMenuItem3,
            this.toolStripSeparator1,
            this.toolStripMenuItem5,
            this.toolStripMenuItem2});
            this.systemTrayMenu.Name = "systemTrayMenu";
            this.systemTrayMenu.Size = new System.Drawing.Size(108, 126);
            // 
            // toolStripMenuItem4
            // 
            this.toolStripMenuItem4.Name = "toolStripMenuItem4";
            this.toolStripMenuItem4.Size = new System.Drawing.Size(107, 22);
            this.toolStripMenuItem4.Text = "Sync";
            this.toolStripMenuItem4.Click += new System.EventHandler(this.toolStripMenuItem4_Click);
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(104, 6);
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(107, 22);
            this.toolStripMenuItem1.Text = "Show";
            this.toolStripMenuItem1.Click += new System.EventHandler(this.toolStripMenuItem1_Click);
            // 
            // toolStripMenuItem3
            // 
            this.toolStripMenuItem3.Name = "toolStripMenuItem3";
            this.toolStripMenuItem3.Size = new System.Drawing.Size(107, 22);
            this.toolStripMenuItem3.Text = "Hide";
            this.toolStripMenuItem3.Click += new System.EventHandler(this.toolStripMenuItem3_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(104, 6);
            // 
            // toolStripMenuItem5
            // 
            this.toolStripMenuItem5.Name = "toolStripMenuItem5";
            this.toolStripMenuItem5.Size = new System.Drawing.Size(107, 22);
            this.toolStripMenuItem5.Text = "About";
            this.toolStripMenuItem5.Click += new System.EventHandler(this.toolStripMenuItem5_Click);
            // 
            // toolStripMenuItem2
            // 
            this.toolStripMenuItem2.Name = "toolStripMenuItem2";
            this.toolStripMenuItem2.Size = new System.Drawing.Size(107, 22);
            this.toolStripMenuItem2.Text = "Exit";
            this.toolStripMenuItem2.Click += new System.EventHandler(this.toolStripMenuItem2_Click);
            // 
            // autoSyncInterval
            // 
            this.autoSyncInterval.Location = new System.Drawing.Point(108, 91);
            this.autoSyncInterval.Maximum = new decimal(new int[] {
            1440,
            0,
            0,
            0});
            this.autoSyncInterval.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.autoSyncInterval.Name = "autoSyncInterval";
            this.autoSyncInterval.Size = new System.Drawing.Size(49, 21);
            this.autoSyncInterval.TabIndex = 3;
            this.autoSyncInterval.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.autoSyncInterval.Value = new decimal(new int[] {
            120,
            0,
            0,
            0});
            this.autoSyncInterval.ValueChanged += new System.EventHandler(this.autoSyncInterval_ValueChanged);
            // 
            // autoSyncCheckBox
            // 
            this.autoSyncCheckBox.AutoSize = true;
            this.autoSyncCheckBox.Location = new System.Drawing.Point(14, 42);
            this.autoSyncCheckBox.Name = "autoSyncCheckBox";
            this.autoSyncCheckBox.Size = new System.Drawing.Size(84, 17);
            this.autoSyncCheckBox.TabIndex = 1;
            this.autoSyncCheckBox.Text = "&Auto Sync";
            this.autoSyncCheckBox.UseVisualStyleBackColor = true;
            this.autoSyncCheckBox.CheckedChanged += new System.EventHandler(this.autoSyncCheckBox_CheckedChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(10, 93);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(89, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Sync &Interval:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(164, 93);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(34, 13);
            this.label4.TabIndex = 4;
            this.label4.Text = "mins";
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.reportSyncResultCheckBox);
            this.groupBox1.Controls.Add(this.runAtStartupCheckBox);
            this.groupBox1.Controls.Add(this.nextSyncLabel);
            this.groupBox1.Controls.Add(this.autoSyncInterval);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.autoSyncCheckBox);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Location = new System.Drawing.Point(6, 394);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(597, 138);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Automation";
            // 
            // reportSyncResultCheckBox
            // 
            this.reportSyncResultCheckBox.AutoSize = true;
            this.reportSyncResultCheckBox.Location = new System.Drawing.Point(14, 65);
            this.reportSyncResultCheckBox.Name = "reportSyncResultCheckBox";
            this.reportSyncResultCheckBox.Size = new System.Drawing.Size(226, 17);
            this.reportSyncResultCheckBox.TabIndex = 6;
            this.reportSyncResultCheckBox.Text = "Re&port Sync Result in System Tray";
            this.reportSyncResultCheckBox.UseVisualStyleBackColor = true;
            // 
            // runAtStartupCheckBox
            // 
            this.runAtStartupCheckBox.AutoSize = true;
            this.runAtStartupCheckBox.Location = new System.Drawing.Point(14, 21);
            this.runAtStartupCheckBox.Name = "runAtStartupCheckBox";
            this.runAtStartupCheckBox.Size = new System.Drawing.Size(160, 17);
            this.runAtStartupCheckBox.TabIndex = 0;
            this.runAtStartupCheckBox.Text = "&Run program at startup";
            this.runAtStartupCheckBox.UseVisualStyleBackColor = true;
            this.runAtStartupCheckBox.CheckedChanged += new System.EventHandler(this.runAtStartupCheckBox_CheckedChanged);
            // 
            // nextSyncLabel
            // 
            this.nextSyncLabel.AutoSize = true;
            this.nextSyncLabel.Location = new System.Drawing.Point(10, 117);
            this.nextSyncLabel.Name = "nextSyncLabel";
            this.nextSyncLabel.Size = new System.Drawing.Size(79, 13);
            this.nextSyncLabel.TabIndex = 5;
            this.nextSyncLabel.Text = "Next Sync in";
            // 
            // syncTimer
            // 
            this.syncTimer.Interval = 1000;
            this.syncTimer.Tick += new System.EventHandler(this.syncTimer_Tick);
            // 
            // groupBox2
            // 
            this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox2.Controls.Add(this.appointmentTimezonesComboBox);
            this.groupBox2.Controls.Add(this.futureMonthTextBox);
            this.groupBox2.Controls.Add(this.pastMonthTextBox);
            this.groupBox2.Controls.Add(this.btSyncAppointments);
            this.groupBox2.Controls.Add(this.appointmentFoldersComboBox);
            this.groupBox2.Controls.Add(this.btSyncNotes);
            this.groupBox2.Controls.Add(this.btSyncContacts);
            this.groupBox2.Controls.Add(this.btPromptDelete);
            this.groupBox2.Controls.Add(this.noteFoldersComboBox);
            this.groupBox2.Controls.Add(this.panel1);
            this.groupBox2.Controls.Add(this.btSyncDelete);
            this.groupBox2.Controls.Add(this.cmbSyncProfile);
            this.groupBox2.Controls.Add(this.contactFoldersComboBox);
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Controls.Add(this.syncOptionBox);
            this.groupBox2.Location = new System.Drawing.Point(6, 115);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(597, 279);
            this.groupBox2.TabIndex = 1;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Sync Options";
            // 
            // appointmentTimezonesComboBox
            // 
            this.appointmentTimezonesComboBox.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.appointmentTimezonesComboBox.FormattingEnabled = true;
            this.appointmentTimezonesComboBox.Items.AddRange(new object[] {
            "Africa/Abidjan",
            "Africa/Accra",
            "Africa/Addis_Ababa",
            "Africa/Asmera",
            "Africa/Bamako",
            "Africa/Bangui",
            "Africa/Banjul",
            "Africa/Blantyre",
            "Africa/Brazzaville",
            "Africa/Bujumbura",
            "Africa/Cairo",
            "Africa/Conakry",
            "Africa/Dakar",
            "Africa/Dar_es_Salaam",
            "Africa/Djibouti",
            "Africa/Douala",
            "Africa/El_Aaiun",
            "Africa/Freetown",
            "Africa/Gaborone",
            "Africa/Harare",
            "Africa/Johannesburg",
            "Africa/Kampala",
            "Africa/Kigali",
            "Africa/Kinshasa",
            "Africa/Lagos",
            "Africa/Libreville",
            "Africa/Lome",
            "Africa/Luanda",
            "Africa/Lubumbashi",
            "Africa/Lusaka",
            "Africa/Malabo",
            "Africa/Maputo",
            "Africa/Maseru",
            "Africa/Mbabane",
            "Africa/Mogadishu",
            "Africa/Monrovia",
            "Africa/Monrovia",
            "Africa/Nairobi",
            "Africa/Ndjamena",
            "Africa/Niamey",
            "Africa/Nouakchott",
            "Africa/Ouagadougou",
            "Africa/Porto-Novo",
            "Africa/Sao_Tome",
            "Africa/Tunis",
            "America/Adak",
            "America/Anchorage",
            "America/Anguilla",
            "America/Antigua",
            "America/Argentina/San_Luis",
            "America/Aruba",
            "America/Asuncion",
            "America/Barbados",
            "America/Belize",
            "America/Bogota",
            "America/Buenos_Aires",
            "America/Caracas",
            "America/Cayenne",
            "America/Cayman",
            "America/Chicago",
            "America/Costa_Rica",
            "America/Curacao",
            "America/Denver",
            "America/Dominica",
            "America/Edmonton",
            "America/El_Salvador",
            "America/Godthab",
            "America/Goose_Bay",
            "America/Grand_Turk",
            "America/Grenada",
            "America/Guadeloupe",
            "America/Guatemala",
            "America/Guayaquil",
            "America/Guyana",
            "America/Halifax",
            "America/Havana",
            "America/Hermosillo",
            "America/Jamaica",
            "America/Juneau",
            "America/Kralendijk",
            "America/La_Paz",
            "America/Lima",
            "America/Los_Angeles",
            "America/Lower_Princes",
            "America/Manaus",
            "America/Marigot",
            "America/Martinique",
            "America/Mexico_City",
            "America/Miquelon",
            "America/Montevideo",
            "America/Montserrat",
            "America/Nassau",
            "America/New_York",
            "America/Noronha",
            "America/Panama",
            "America/Paramaribo",
            "America/Paramaribo",
            "America/Port-au-Prince",
            "America/Port_of_Spain",
            "America/Puerto_Rico",
            "America/Rio_Branco",
            "America/Santiago",
            "America/Santo_Domingo",
            "America/Sao_Paulo",
            "America/Scoresbysund",
            "America/Scoresbysund",
            "America/St_Johns",
            "America/St_Kitts",
            "America/St_Lucia",
            "America/St_Thomas",
            "America/St_Vincent",
            "America/Tegucigalpa",
            "America/Thule",
            "America/Tijuana",
            "America/Toronto",
            "America/Tortola",
            "America/Vancouver",
            "America/Winnipeg",
            "America/Yakutat",
            "Antarctica/Casey",
            "Antarctica/Davis",
            "Antarctica/DumontDUrville",
            "Antarctica/Macquarie",
            "Antarctica/Mawson",
            "Antarctica/McMurdo",
            "Antarctica/Palmer",
            "Antarctica/Rothera",
            "Antarctica/Syowa",
            "Antarctica/Vostok",
            "Asia/Aden",
            "Asia/Almaty",
            "Asia/Almaty",
            "Asia/Amman",
            "Asia/Anadyr",
            "Asia/Aqtau",
            "Asia/Aqtau",
            "Asia/Aqtobe",
            "Asia/Aqtobe",
            "Asia/Aqtobe",
            "Asia/Ashgabat",
            "Asia/Ashgabat",
            "Asia/Baghdad",
            "Asia/Bahrain",
            "Asia/Baku",
            "Asia/Baku",
            "Asia/Bangkok",
            "Asia/Beirut",
            "Asia/Bishkek",
            "Asia/Bishkek",
            "Asia/Brunei",
            "Asia/Calcutta",
            "Asia/Choibalsan",
            "Asia/Chongqing",
            "Asia/Colombo",
            "Asia/Colombo",
            "Asia/Damascus",
            "Asia/Dhaka",
            "Asia/Dhaka",
            "Asia/Dili",
            "Asia/Dubai",
            "Asia/Dushanbe",
            "Asia/Dushanbe",
            "Asia/Harbin",
            "Asia/Hong_Kong",
            "Asia/Hovd",
            "Asia/Irkutsk",
            "Asia/Jakarta",
            "Asia/Jayapura",
            "Asia/Jerusalem",
            "Asia/Kabul",
            "Asia/Kamchatka",
            "Asia/Karachi",
            "Asia/Karachi",
            "Asia/Kashgar",
            "Asia/Katmandu",
            "Asia/Krasnoyarsk",
            "Asia/Kuala_Lumpur",
            "Asia/Kuching",
            "Asia/Kuching",
            "Asia/Kuwait",
            "Asia/Macau",
            "Asia/Macau",
            "Asia/Magadan",
            "Asia/Makassar",
            "Asia/Manila",
            "Asia/Muscat",
            "Asia/Nicosia",
            "Asia/Novosibirsk",
            "Asia/Omsk",
            "Asia/Oral",
            "Asia/Oral",
            "Asia/Phnom_Penh",
            "Asia/Pyongyang",
            "Asia/Qatar",
            "Asia/Qyzylorda",
            "Asia/Qyzylorda",
            "Asia/Rangoon",
            "Asia/Riyadh",
            "Asia/Saigon",
            "Asia/Sakhalin",
            "Asia/Samarkand",
            "Asia/Seoul",
            "Asia/Shanghai",
            "Asia/Singapore",
            "Asia/Taipei",
            "Asia/Tashkent",
            "Asia/Tashkent",
            "Asia/Tbilisi",
            "Asia/Tbilisi",
            "Asia/Tehran",
            "Asia/Thimphu",
            "Asia/Tokyo",
            "Asia/Ulaanbaatar",
            "Asia/Urumqi",
            "Asia/Vientiane",
            "Asia/Vladivostok",
            "Asia/Yakutsk",
            "Asia/Yekaterinburg",
            "Asia/Yekaterinburg",
            "Asia/Yerevan",
            "Asia/Yerevan",
            "Atlantic/Azores",
            "Atlantic/Bermuda",
            "Atlantic/Canary",
            "Atlantic/Cape_Verde",
            "Atlantic/Faeroe",
            "Atlantic/Reykjavik",
            "Atlantic/South_Georgia",
            "Atlantic/St_Helena",
            "Atlantic/Stanley",
            "Australia/Adelaide",
            "Australia/Eucla",
            "Australia/Lord_Howe",
            "Australia/Perth",
            "Australia/Sydney",
            "Europe/Amsterdam",
            "Europe/Andorra",
            "Europe/Athens",
            "Europe/Belgrade",
            "Europe/Berlin",
            "Europe/Bratislava",
            "Europe/Brussels",
            "Europe/Bucharest",
            "Europe/Budapest",
            "Europe/Copenhagen",
            "Europe/Dublin",
            "Europe/Dublin",
            "Europe/Gibraltar",
            "Europe/Helsinki",
            "Europe/Istanbul",
            "Europe/Ljubljana",
            "Europe/London",
            "Europe/London",
            "Europe/Luxembourg",
            "Europe/Madrid",
            "Europe/Malta",
            "Europe/Mariehamn",
            "Europe/Monaco",
            "Europe/Moscow",
            "Europe/Oslo",
            "Europe/Paris",
            "Europe/Podgorica",
            "Europe/Prague",
            "Europe/Rome",
            "Europe/Samara",
            "Europe/Samara",
            "Europe/San_Marino",
            "Europe/Sarajevo",
            "Europe/Skopje",
            "Europe/Sofia",
            "Europe/Stockholm",
            "Europe/Tirane",
            "Europe/Vaduz",
            "Europe/Vatican",
            "Europe/Vienna",
            "Europe/Volgograd",
            "Europe/Warsaw",
            "Europe/Zagreb",
            "Europe/Zurich",
            "Indian/Antananarivo",
            "Indian/Chagos",
            "Indian/Christmas",
            "Indian/Cocos",
            "Indian/Comoro",
            "Indian/Kerguelen",
            "Indian/Mahe",
            "Indian/Maldives",
            "Indian/Mauritius",
            "Indian/Mayotte",
            "Indian/Reunion",
            "Pacific/Apia",
            "Pacific/Auckland",
            "Pacific/Chatham",
            "Pacific/Easter",
            "Pacific/Efate",
            "Pacific/Enderbury",
            "Pacific/Fakaofo",
            "Pacific/Fiji",
            "Pacific/Funafuti",
            "Pacific/Galapagos",
            "Pacific/Gambier",
            "Pacific/Guadalcanal",
            "Pacific/Guam",
            "Pacific/Guam",
            "Pacific/Honolulu",
            "Pacific/Kiritimati",
            "Pacific/Kosrae",
            "Pacific/Kwajalein",
            "Pacific/Majuro",
            "Pacific/Marquesas",
            "Pacific/Nauru",
            "Pacific/Niue",
            "Pacific/Norfolk",
            "Pacific/Noumea",
            "Pacific/Palau",
            "Pacific/Pitcairn",
            "Pacific/Ponape",
            "Pacific/Port_Moresby",
            "Pacific/Rarotonga",
            "Pacific/Saipan",
            "Pacific/Saipan",
            "Pacific/Tahiti",
            "Pacific/Tarawa",
            "Pacific/Tongatapu",
            "Pacific/Truk",
            "Pacific/Wake",
            "Pacific/Wallis"});
            this.appointmentTimezonesComboBox.Location = new System.Drawing.Point(393, 119);
            this.appointmentTimezonesComboBox.Name = "appointmentTimezonesComboBox";
            this.appointmentTimezonesComboBox.Size = new System.Drawing.Size(121, 21);
            this.appointmentTimezonesComboBox.Sorted = true;
            this.appointmentTimezonesComboBox.TabIndex = 13;
            this.toolTip.SetToolTip(this.appointmentTimezonesComboBox, "Select or enter Timezone (default is UTC), only for Recurrences!!!");
            this.appointmentTimezonesComboBox.TextChanged += new System.EventHandler(this.appointmentTimezonesComboBox_TextChanged);
            // 
            // futureMonthTextBox
            // 
            this.futureMonthTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.futureMonthTextBox.Location = new System.Drawing.Point(556, 120);
            this.futureMonthTextBox.Name = "futureMonthTextBox";
            this.futureMonthTextBox.Size = new System.Drawing.Size(33, 21);
            this.futureMonthTextBox.TabIndex = 12;
            this.futureMonthTextBox.Text = "0";
            this.toolTip.SetToolTip(this.futureMonthTextBox, "How many months into the future (0 if all)");
            this.futureMonthTextBox.Validating += new System.ComponentModel.CancelEventHandler(this.futureMonthTextBox_Validating);
            // 
            // pastMonthTextBox
            // 
            this.pastMonthTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.pastMonthTextBox.Location = new System.Drawing.Point(520, 120);
            this.pastMonthTextBox.Name = "pastMonthTextBox";
            this.pastMonthTextBox.Size = new System.Drawing.Size(33, 21);
            this.pastMonthTextBox.TabIndex = 11;
            this.pastMonthTextBox.Text = "1";
            this.toolTip.SetToolTip(this.pastMonthTextBox, "How many months into the past (0 if all)");
            this.pastMonthTextBox.Validating += new System.ComponentModel.CancelEventHandler(this.pastMonthTextBox_Validating);
            // 
            // btSyncAppointments
            // 
            this.btSyncAppointments.AutoSize = true;
            this.btSyncAppointments.Location = new System.Drawing.Point(453, 46);
            this.btSyncAppointments.Name = "btSyncAppointments";
            this.btSyncAppointments.Size = new System.Drawing.Size(136, 17);
            this.btSyncAppointments.TabIndex = 10;
            this.btSyncAppointments.Text = "Sync &Appointments";
            this.toolTip.SetToolTip(this.btSyncAppointments, "This specifies whether notes are synchronized.");
            this.btSyncAppointments.UseVisualStyleBackColor = true;
            this.btSyncAppointments.CheckedChanged += new System.EventHandler(this.btSyncAppointments_CheckedChanged);
            // 
            // appointmentFoldersComboBox
            // 
            this.appointmentFoldersComboBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.appointmentFoldersComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.appointmentFoldersComboBox.FormattingEnabled = true;
            this.appointmentFoldersComboBox.Location = new System.Drawing.Point(6, 119);
            this.appointmentFoldersComboBox.Name = "appointmentFoldersComboBox";
            this.appointmentFoldersComboBox.Size = new System.Drawing.Size(381, 21);
            this.appointmentFoldersComboBox.TabIndex = 9;
            this.toolTip.SetToolTip(this.appointmentFoldersComboBox, "Select the Outlook Appointments folder you want to sync");
            this.appointmentFoldersComboBox.SelectedIndexChanged += new System.EventHandler(this.appointmentFoldersComboBox_SelectedIndexChanged);
            // 
            // btSyncNotes
            // 
            this.btSyncNotes.AutoSize = true;
            this.btSyncNotes.Location = new System.Drawing.Point(357, 46);
            this.btSyncNotes.Name = "btSyncNotes";
            this.btSyncNotes.Size = new System.Drawing.Size(90, 17);
            this.btSyncNotes.TabIndex = 5;
            this.btSyncNotes.Text = "Sync &Notes";
            this.toolTip.SetToolTip(this.btSyncNotes, "This specifies whether notes are synchronized.");
            this.btSyncNotes.UseVisualStyleBackColor = true;
            this.btSyncNotes.CheckedChanged += new System.EventHandler(this.btSyncNotes_CheckedChanged);
            // 
            // btSyncContacts
            // 
            this.btSyncContacts.AutoSize = true;
            this.btSyncContacts.Checked = true;
            this.btSyncContacts.CheckState = System.Windows.Forms.CheckState.Checked;
            this.btSyncContacts.Location = new System.Drawing.Point(243, 46);
            this.btSyncContacts.Name = "btSyncContacts";
            this.btSyncContacts.Size = new System.Drawing.Size(108, 17);
            this.btSyncContacts.TabIndex = 6;
            this.btSyncContacts.Text = "Sync &Contacts";
            this.toolTip.SetToolTip(this.btSyncContacts, "This specifies whether contacts are synchronized.");
            this.btSyncContacts.UseVisualStyleBackColor = true;
            this.btSyncContacts.CheckedChanged += new System.EventHandler(this.btSyncContacts_CheckedChanged);
            // 
            // btPromptDelete
            // 
            this.btPromptDelete.AutoSize = true;
            this.btPromptDelete.Checked = true;
            this.btPromptDelete.CheckState = System.Windows.Forms.CheckState.Checked;
            this.btPromptDelete.Location = new System.Drawing.Point(119, 46);
            this.btPromptDelete.Name = "btPromptDelete";
            this.btPromptDelete.Size = new System.Drawing.Size(118, 17);
            this.btPromptDelete.TabIndex = 8;
            this.btPromptDelete.Text = "Prompt De&letion";
            this.toolTip.SetToolTip(this.btPromptDelete, resources.GetString("btPromptDelete.ToolTip"));
            this.btPromptDelete.UseVisualStyleBackColor = true;
            // 
            // noteFoldersComboBox
            // 
            this.noteFoldersComboBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.noteFoldersComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.noteFoldersComboBox.FormattingEnabled = true;
            this.noteFoldersComboBox.Location = new System.Drawing.Point(6, 92);
            this.noteFoldersComboBox.Name = "noteFoldersComboBox";
            this.noteFoldersComboBox.Size = new System.Drawing.Size(583, 21);
            this.noteFoldersComboBox.TabIndex = 7;
            this.toolTip.SetToolTip(this.noteFoldersComboBox, "Select the Outlook Notes folder you want to sync");
            this.noteFoldersComboBox.SelectedIndexChanged += new System.EventHandler(this.noteFoldersComboBox_SelectedIndexChanged);
            // 
            // panel1
            // 
            this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Location = new System.Drawing.Point(6, 155);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(580, 1);
            this.panel1.TabIndex = 3;
            // 
            // btSyncDelete
            // 
            this.btSyncDelete.AutoSize = true;
            this.btSyncDelete.Checked = true;
            this.btSyncDelete.CheckState = System.Windows.Forms.CheckState.Checked;
            this.btSyncDelete.Location = new System.Drawing.Point(8, 46);
            this.btSyncDelete.Name = "btSyncDelete";
            this.btSyncDelete.Size = new System.Drawing.Size(105, 17);
            this.btSyncDelete.TabIndex = 2;
            this.btSyncDelete.Text = "Sync &Deletion";
            this.toolTip.SetToolTip(this.btSyncDelete, "This specifies whether deletions are\r\nsynchronized. Enabling this option\r\nmeans i" +
                    "f you delete a contact from\r\nGoogle, then it will be deleted from\r\nOutlook and v" +
                    "ice versa.");
            this.btSyncDelete.UseVisualStyleBackColor = true;
            this.btSyncDelete.CheckedChanged += new System.EventHandler(this.btSyncDelete_CheckedChanged);
            // 
            // cmbSyncProfile
            // 
            this.cmbSyncProfile.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbSyncProfile.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSyncProfile.FormattingEnabled = true;
            this.cmbSyncProfile.Location = new System.Drawing.Point(100, 19);
            this.cmbSyncProfile.Name = "cmbSyncProfile";
            this.cmbSyncProfile.Size = new System.Drawing.Size(489, 21);
            this.cmbSyncProfile.TabIndex = 1;
            this.toolTip.SetToolTip(this.cmbSyncProfile, "This is a profile name of your choice.\r\nIt must be unique in each computer\r\nand a" +
                    "ccount you intend to sync with\r\nyour Google Mail account.");
            this.cmbSyncProfile.SelectedIndexChanged += new System.EventHandler(this.cmbSyncProfile_SelectedIndexChanged);
            // 
            // contactFoldersComboBox
            // 
            this.contactFoldersComboBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.contactFoldersComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.contactFoldersComboBox.FormattingEnabled = true;
            this.contactFoldersComboBox.Location = new System.Drawing.Point(6, 65);
            this.contactFoldersComboBox.Name = "contactFoldersComboBox";
            this.contactFoldersComboBox.Size = new System.Drawing.Size(583, 21);
            this.contactFoldersComboBox.TabIndex = 6;
            this.toolTip.SetToolTip(this.contactFoldersComboBox, "Select the Outlook Contacts folder you want to sync");
            this.contactFoldersComboBox.SelectedIndexChanged += new System.EventHandler(this.contacFoldersComboBox_SelectedIndexChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(7, 22);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(80, 13);
            this.label5.TabIndex = 0;
            this.label5.Text = "&Sync Profile:";
            // 
            // lastSyncLabel
            // 
            this.lastSyncLabel.AutoSize = true;
            this.lastSyncLabel.Location = new System.Drawing.Point(7, 16);
            this.lastSyncLabel.Name = "lastSyncLabel";
            this.lastSyncLabel.Size = new System.Drawing.Size(80, 13);
            this.lastSyncLabel.TabIndex = 0;
            this.lastSyncLabel.Text = "Last Sync on";
            // 
            // logGroupBox
            // 
            this.logGroupBox.Controls.Add(this.syncConsole);
            this.logGroupBox.Controls.Add(this.lastSyncLabel);
            this.logGroupBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.logGroupBox.Location = new System.Drawing.Point(614, 0);
            this.logGroupBox.Name = "logGroupBox";
            this.logGroupBox.Size = new System.Drawing.Size(523, 538);
            this.logGroupBox.TabIndex = 2;
            this.logGroupBox.TabStop = false;
            this.logGroupBox.Text = "Sync Details && Log";
            // 
            // syncConsole
            // 
            this.syncConsole.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.syncConsole.BackColor = System.Drawing.SystemColors.Info;
            this.syncConsole.Location = new System.Drawing.Point(6, 33);
            this.syncConsole.Multiline = true;
            this.syncConsole.Name = "syncConsole";
            this.syncConsole.ReadOnly = true;
            this.syncConsole.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.syncConsole.Size = new System.Drawing.Size(514, 499);
            this.syncConsole.TabIndex = 1;
            this.toolTip.SetToolTip(this.syncConsole, "This window shows information\r\n from the last sync.");
            // 
            // groupBox4
            // 
            this.groupBox4.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox4.Controls.Add(this.chkUseFileAs);
            this.groupBox4.Controls.Add(this.proxySettingsLinkLabel);
            this.groupBox4.Controls.Add(this.label2);
            this.groupBox4.Controls.Add(this.label3);
            this.groupBox4.Controls.Add(this.UserName);
            this.groupBox4.Controls.Add(this.Password);
            this.groupBox4.Location = new System.Drawing.Point(6, 20);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(597, 92);
            this.groupBox4.TabIndex = 0;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Google Account";
            // 
            // chkUseFileAs
            // 
            this.chkUseFileAs.AutoSize = true;
            this.chkUseFileAs.Checked = true;
            this.chkUseFileAs.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkUseFileAs.Location = new System.Drawing.Point(161, 69);
            this.chkUseFileAs.Name = "chkUseFileAs";
            this.chkUseFileAs.Size = new System.Drawing.Size(281, 17);
            this.chkUseFileAs.TabIndex = 7;
            this.chkUseFileAs.Text = "Use Outlook Contact\'s FileAs for Google Title";
            this.chkUseFileAs.UseVisualStyleBackColor = true;
            // 
            // proxySettingsLinkLabel
            // 
            this.proxySettingsLinkLabel.AutoSize = true;
            this.proxySettingsLinkLabel.Location = new System.Drawing.Point(7, 69);
            this.proxySettingsLinkLabel.Name = "proxySettingsLinkLabel";
            this.proxySettingsLinkLabel.Size = new System.Drawing.Size(90, 13);
            this.proxySettingsLinkLabel.TabIndex = 4;
            this.proxySettingsLinkLabel.TabStop = true;
            this.proxySettingsLinkLabel.Text = "Proxy Settings";
            this.toolTip.SetToolTip(this.proxySettingsLinkLabel, resources.GetString("proxySettingsLinkLabel.ToolTip"));
            this.proxySettingsLinkLabel.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.proxySettingsLinkLabel_LinkClicked);
            // 
            // resetMatchesLinkLabel
            // 
            this.resetMatchesLinkLabel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.resetMatchesLinkLabel.AutoSize = true;
            this.resetMatchesLinkLabel.Location = new System.Drawing.Point(52, 556);
            this.resetMatchesLinkLabel.Name = "resetMatchesLinkLabel";
            this.resetMatchesLinkLabel.Size = new System.Drawing.Size(89, 13);
            this.resetMatchesLinkLabel.TabIndex = 2;
            this.resetMatchesLinkLabel.TabStop = true;
            this.resetMatchesLinkLabel.Text = "&Reset Matches";
            this.toolTip.SetToolTip(this.resetMatchesLinkLabel, "This unlinks Outlook contacts with their\r\ncorresponding Google contatcs. If you\r\n" +
                    "accidentaly delete a contact and you\r\ndont want the deletion to be synchronised," +
                    "\r\nclick  this link.");
            this.resetMatchesLinkLabel.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.resetMatchesLinkLabel_LinkClicked);
            // 
            // Donate
            // 
            this.Donate.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.Donate.BackColor = System.Drawing.Color.Transparent;
            this.Donate.Image = ((System.Drawing.Image)(resources.GetObject("Donate.Image")));
            this.Donate.Location = new System.Drawing.Point(12, 556);
            this.Donate.Name = "Donate";
            this.Donate.Size = new System.Drawing.Size(34, 34);
            this.Donate.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.Donate.TabIndex = 4;
            this.Donate.TabStop = false;
            this.toolTip.SetToolTip(this.Donate, resources.GetString("Donate.ToolTip"));
            this.Donate.Click += new System.EventHandler(this.Donate_Click);
            this.Donate.MouseEnter += new System.EventHandler(this.Donate_MouseEnter);
            this.Donate.MouseLeave += new System.EventHandler(this.Donate_MouseLeave);
            // 
            // pictureBoxExit
            // 
            this.pictureBoxExit.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.pictureBoxExit.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("pictureBoxExit.BackgroundImage")));
            this.pictureBoxExit.Image = ((System.Drawing.Image)(resources.GetObject("pictureBoxExit.Image")));
            this.pictureBoxExit.Location = new System.Drawing.Point(1124, 559);
            this.pictureBoxExit.Margin = new System.Windows.Forms.Padding(0);
            this.pictureBoxExit.Name = "pictureBoxExit";
            this.pictureBoxExit.Size = new System.Drawing.Size(24, 25);
            this.pictureBoxExit.TabIndex = 5;
            this.pictureBoxExit.TabStop = false;
            this.toolTip.SetToolTip(this.pictureBoxExit, "Exit Go Contact Sync Mod");
            this.pictureBoxExit.Click += new System.EventHandler(this.pictureBoxExit_Click);
            // 
            // settingsGroupBox
            // 
            this.settingsGroupBox.Controls.Add(this.groupBox1);
            this.settingsGroupBox.Controls.Add(this.groupBox4);
            this.settingsGroupBox.Controls.Add(this.groupBox2);
            this.settingsGroupBox.Dock = System.Windows.Forms.DockStyle.Left;
            this.settingsGroupBox.Location = new System.Drawing.Point(0, 0);
            this.settingsGroupBox.Name = "settingsGroupBox";
            this.settingsGroupBox.Size = new System.Drawing.Size(609, 538);
            this.settingsGroupBox.TabIndex = 0;
            this.settingsGroupBox.TabStop = false;
            this.settingsGroupBox.Text = "Program Settings";
            // 
            // actionsTableLayout
            // 
            this.actionsTableLayout.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.actionsTableLayout.ColumnCount = 3;
            this.actionsTableLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.actionsTableLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.actionsTableLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 72F));
            this.actionsTableLayout.Controls.Add(this.cancelButton, 0, 0);
            this.actionsTableLayout.Controls.Add(this.syncButton, 0, 0);
            this.actionsTableLayout.Controls.Add(this.hideButton, 2, 0);
            this.actionsTableLayout.Location = new System.Drawing.Point(879, 556);
            this.actionsTableLayout.Name = "actionsTableLayout";
            this.actionsTableLayout.RowCount = 1;
            this.actionsTableLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.actionsTableLayout.Size = new System.Drawing.Size(242, 31);
            this.actionsTableLayout.TabIndex = 1;
            // 
            // cancelButton
            // 
            this.cancelButton.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.cancelButton.Enabled = false;
            this.cancelButton.Location = new System.Drawing.Point(88, 3);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(79, 25);
            this.cancelButton.TabIndex = 2;
            this.cancelButton.Text = "&Cancel";
            this.cancelButton.UseVisualStyleBackColor = true;
            this.cancelButton.Click += new System.EventHandler(this.CancelButton_Click);
            // 
            // hideButton
            // 
            this.hideButton.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.hideButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.hideButton.Location = new System.Drawing.Point(173, 3);
            this.hideButton.Name = "hideButton";
            this.hideButton.Size = new System.Drawing.Size(66, 25);
            this.hideButton.TabIndex = 3;
            this.hideButton.Text = "&Hide";
            this.hideButton.UseVisualStyleBackColor = true;
            this.hideButton.Click += new System.EventHandler(this.hideButton_Click);
            // 
            // MainPanel
            // 
            this.MainPanel.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.MainPanel.Controls.Add(this.logGroupBox);
            this.MainPanel.Controls.Add(this.MainSplitter);
            this.MainPanel.Controls.Add(this.settingsGroupBox);
            this.MainPanel.Location = new System.Drawing.Point(12, 12);
            this.MainPanel.Name = "MainPanel";
            this.MainPanel.Size = new System.Drawing.Size(1137, 538);
            this.MainPanel.TabIndex = 0;
            // 
            // MainSplitter
            // 
            this.MainSplitter.Location = new System.Drawing.Point(609, 0);
            this.MainSplitter.Name = "MainSplitter";
            this.MainSplitter.Size = new System.Drawing.Size(5, 538);
            this.MainSplitter.TabIndex = 5;
            this.MainSplitter.TabStop = false;
            // 
            // iconTimer
            // 
            this.iconTimer.Interval = 150;
            this.iconTimer.Tick += new System.EventHandler(this.iconTimer_Tick);
            // 
            // SettingsForm
            // 
            this.AcceptButton = this.syncButton;
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1161, 599);
            this.Controls.Add(this.Donate);
            this.Controls.Add(this.MainPanel);
            this.Controls.Add(this.pictureBoxExit);
            this.Controls.Add(this.resetMatchesLinkLabel);
            this.Controls.Add(this.actionsTableLayout);
            this.Font = new System.Drawing.Font("Verdana", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.HelpButton = true;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(725, 528);
            this.Name = "SettingsForm";
            this.Text = "GO Contact Sync Mod - Settings";
            this.WindowState = System.Windows.Forms.FormWindowState.Minimized;
            this.HelpButtonClicked += new System.ComponentModel.CancelEventHandler(this.SettingsForm_HelpButtonClicked);
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.SettingsForm_FormClosing);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.SettingsForm_FormClosed);
            this.Load += new System.EventHandler(this.SettingsForm_Load);
            this.HelpRequested += new System.Windows.Forms.HelpEventHandler(this.SettingsForm_HelpRequested);
            this.Resize += new System.EventHandler(this.SettingsForm_Resize);
            this.systemTrayMenu.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.autoSyncInterval)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.logGroupBox.ResumeLayout(false);
            this.logGroupBox.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.Donate)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxExit)).EndInit();
            this.settingsGroupBox.ResumeLayout(false);
            this.actionsTableLayout.ResumeLayout(false);
            this.MainPanel.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox Password;
        private System.Windows.Forms.TextBox UserName;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button syncButton;
        private System.Windows.Forms.CheckedListBox syncOptionBox;
        internal System.Windows.Forms.NotifyIcon notifyIcon;
        private System.Windows.Forms.NumericUpDown autoSyncInterval;
        private System.Windows.Forms.CheckBox autoSyncCheckBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Timer syncTimer;
        private System.Windows.Forms.Label nextSyncLabel;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label lastSyncLabel;
        private System.Windows.Forms.GroupBox logGroupBox;
        private System.Windows.Forms.TextBox syncConsole;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.ContextMenuStrip systemTrayMenu;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem2;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem3;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem4;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.CheckBox runAtStartupCheckBox;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem5;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ToolTip toolTip;
        private System.Windows.Forms.CheckBox btSyncDelete;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.GroupBox settingsGroupBox;
        private System.Windows.Forms.TableLayoutPanel actionsTableLayout;
        private System.Windows.Forms.Panel MainPanel;
        private System.Windows.Forms.Splitter MainSplitter;
        private System.Windows.Forms.LinkLabel resetMatchesLinkLabel;
        internal System.Windows.Forms.PictureBox Donate;
        private System.Windows.Forms.Button hideButton;
        private System.Windows.Forms.LinkLabel proxySettingsLinkLabel;
        private System.Windows.Forms.CheckBox reportSyncResultCheckBox;
        private System.Windows.Forms.CheckBox btSyncNotes;
        private System.Windows.Forms.CheckBox btSyncContacts;
        private System.Windows.Forms.ComboBox contactFoldersComboBox;
        private System.Windows.Forms.ComboBox cmbSyncProfile;
        private System.Windows.Forms.ComboBox noteFoldersComboBox;
        private System.Windows.Forms.CheckBox btPromptDelete;
        private System.Windows.Forms.CheckBox chkUseFileAs;
        private System.Windows.Forms.PictureBox pictureBoxExit;
        private System.Windows.Forms.Button cancelButton;
        private System.Windows.Forms.Timer iconTimer;
        private System.Windows.Forms.ComboBox appointmentFoldersComboBox;
        private System.Windows.Forms.CheckBox btSyncAppointments;
        private System.Windows.Forms.TextBox futureMonthTextBox;
        private System.Windows.Forms.TextBox pastMonthTextBox;
        private System.Windows.Forms.ComboBox appointmentTimezonesComboBox;

    }
}

