﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Forms;

namespace GoContactSyncMod
{
    public partial class ErrorDialog : Form
    {
        public ErrorDialog()
        {
            InitializeComponent();
        }

        public void setErrorText(Exception ex)
        {
            if (VersionInformation.isNewVersionAvailable())
            {
                richTextBoxError.AppendText(Environment.NewLine);
                richTextBoxError.AppendText("NEW VERSION AVAILABLE - ");
                LinkLabel downloadLink = new LinkLabel();
                downloadLink.Text = "DOWNLOAD NOW";
                downloadLink.AutoSize = true;
                downloadLink.LinkColor = Color.FromArgb(0, 102, 204);
                downloadLink.Location = richTextBoxError.GetPositionFromCharIndex(richTextBoxError.TextLength);
                downloadLink.LinkClicked += (openDowloadUrl);
                richTextBoxError.Controls.Add(downloadLink);
                richTextBoxError.AppendText(downloadLink.Text);
                richTextBoxError.AppendText(Environment.NewLine);
                richTextBoxError.AppendText(Environment.NewLine);
                AppendTextWithColor("PLEASE UPDATE TO THE LATEST VERSION!" + Environment.NewLine, Color.Firebrick);
            }

            AppendTextWithColor("FIRST CHECK IF THIS ERROR HAS ALREADY BEEN REPORTED!", Color.Firebrick);
            AppendTextWithColor(Environment.NewLine + "IF THE PROBLEM STILL EXISTS WRITE AN ERROR REPORT ", Color.Firebrick);
            LinkLabel bugsLink = new LinkLabel();
            bugsLink.Text = "HERE!";
            bugsLink.AutoSize = true;
            bugsLink.LinkColor = Color.FromArgb(0, 102, 204);
            bugsLink.Location = richTextBoxError.GetPositionFromCharIndex(richTextBoxError.TextLength);
            bugsLink.LinkClicked += (openBugsUrl);
            richTextBoxError.Controls.Add(bugsLink);

            richTextBoxError.AppendText(Environment.NewLine);
            richTextBoxError.AppendText(Environment.NewLine);

            richTextBoxError.AppendText("GCSM VERSION:    " + VersionInformation.getGCSMVersion().ToString());
            richTextBoxError.AppendText(Environment.NewLine);
            richTextBoxError.AppendText("OUTLOOK VERSION: " + VersionInformation.GetOutlookVersion(Synchronizer.OutlookApplication).ToString() + Environment.NewLine);
            richTextBoxError.AppendText("OS VERSION:      " + VersionInformation.GetWindowsVersionName() + Environment.NewLine);
            richTextBoxError.AppendText(Environment.NewLine);
            richTextBoxError.AppendText("ERROR MESAGE:" + Environment.NewLine + Environment.NewLine);
            AppendTextWithColor(ex.Message + Environment.NewLine,Color.Firebrick);
            richTextBoxError.AppendText(Environment.NewLine);
            richTextBoxError.AppendText("ERROR MESAGE STACK TRACE:" + Environment.NewLine + Environment.NewLine);
            if (ex.StackTrace != null)
                AppendTextWithColor(ex.StackTrace, Color.Firebrick);
            else
                AppendTextWithColor("NO STACK TRACE AVAILABLE",Color.Firebrick);

            string message = richTextBoxError.Text.Replace("\n", "\r\n");
            //copy to clipboard
            try
            {
                System.Windows.Clipboard.SetText(message);
            }
            catch (Exception e)
            {
                Logger.Log("Message couldn't be copied to clipboard: " + e.Message, EventType.Debug);
            }
        }

        public string getErrorText()
        {
            return richTextBoxError.Text;
        }

        private void AppendTextWithColor(string text, Color color)
        {
            int start = richTextBoxError.TextLength;
            richTextBoxError.AppendText(text);
            int end = richTextBoxError.TextLength;

            // Textbox may transform chars, so (end-start) != text.Length
            richTextBoxError.Select(start, end - start);
            {
                richTextBoxError.SelectionColor = color;
                // could set box.SelectionBackColor, box.SelectionFont too.
            }
            richTextBoxError.SelectionLength = 0; // clear
        }

        private void openDowloadUrl(object sender, EventArgs e)
        {
            Process.Start("https://sourceforge.net/projects/googlesyncmod/files/latest/download");
        }

        private void openBugsUrl(object sender, EventArgs e)
        {
            Process.Start("http://sourceforge.net/p/googlesyncmod/bugs/?source=navbar");
        }

        private void ErrorDialog_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.Visible = false;
        }
    }
}
