using System;
using System.Collections.Generic;
using System.Text;
using NUnit.Framework;
using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Interop.Outlook;
using System.Threading;
using System.Web;
using System.Net;
using System.IO;
using System.Drawing;
using System.Configuration;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;

namespace GoContactSyncMod.UnitTests
{
    [TestFixture]
    public class SyncAppointmentsTests
    {
        Syncronizer sync;

        static Logger.LogUpdatedHandler _logUpdateHandler = null;

        //Constants for test appointment
        const string name = "AN_OUTLOOK_TEST_APPOINTMENT";
        //readonly When whenDay = new When(DateTime.Now, DateTime.Now, true);
        //readonly When whenTime = new When(DateTime.Now, DateTime.Now.AddHours(1), false);
        //ToDo:const string groupName = "A TEST GROUP";


        [TestFixtureSetUp]
        public void Init()
        {
            //string timestamp = DateTime.Now.Ticks.ToString();            
            if (_logUpdateHandler == null)
            {
                _logUpdateHandler = new Logger.LogUpdatedHandler(Logger_LogUpdated);
                Logger.LogUpdated += _logUpdateHandler;
            }

            string gmailUsername;
            string gmailPassword;
            string syncProfile;
            string syncContactsFolder;
            string syncNotesFolder;
            string syncAppointmentsFolder;

            GoogleAPITests.LoadSettings(out gmailUsername, out gmailPassword, out syncProfile, out syncContactsFolder, out syncNotesFolder, out syncAppointmentsFolder);

            sync = new Syncronizer();
            sync.SyncAppointments = true;
            sync.SyncNotes = false;
            sync.SyncContacts = false;
            sync.SyncProfile = syncProfile;
            Assert.IsNotNull(syncAppointmentsFolder);
            Syncronizer.SyncAppointmentsFolder = syncAppointmentsFolder;
            Syncronizer.MonthsInPast = 1;
            Syncronizer.MonthsInFuture = 1;

            sync.LoginToGoogle(gmailUsername, gmailPassword);
            sync.LoginToOutlook();

        }

        [SetUp]
        public void SetUp()
        {
            // delete previously failed test appointments
            DeleteTestAppointments();

        }

        private void DeleteTestAppointments()
        {
            sync.LoadAppointments();

            Outlook.AppointmentItem outlookAppointment = sync.OutlookAppointments.Find("[Subject] = '" + name + "'") as Outlook.AppointmentItem;
            while (outlookAppointment != null)
            {
                DeleteTestAppointment(outlookAppointment);
                outlookAppointment = sync.OutlookAppointments.Find("[Subject] = '" + name + "'") as Outlook.AppointmentItem;
            }

            foreach (Event googleAppointment in sync.GoogleAppointments)
            {
                if (googleAppointment != null &&
                    googleAppointment.Summary != null && 
                    googleAppointment.Summary == name)
                {
                    DeleteTestAppointment(googleAppointment);
                }
            }
        }

        void Logger_LogUpdated(string message)
        {
            Console.WriteLine(message);
        }

        [TestFixtureTearDown]
        public void TearDown()
        {
            sync.LogoffOutlook();
            sync.LogoffGoogle();
        }
        

        [Test]
        public void TestSync_Time()
        {
            sync.SyncOption = SyncOption.MergeOutlookWins;

            // create new appointment to sync
            Outlook.AppointmentItem outlookAppointment = Syncronizer.CreateOutlookAppointmentItem(Syncronizer.SyncAppointmentsFolder);
            outlookAppointment.Subject = name;
            outlookAppointment.Start = DateTime.Now;
            outlookAppointment.End = DateTime.Now.AddHours(1);
            outlookAppointment.AllDayEvent = false;

            outlookAppointment.Save();


            sync.SyncOption = SyncOption.OutlookToGoogleOnly;

            var googleAppointment = new Event();
            sync.UpdateAppointment(outlookAppointment, ref googleAppointment);

            googleAppointment = null;

            sync.SyncOption = SyncOption.GoogleToOutlookOnly;
            //load the same appointment from google.
            MatchAppointments(sync);
            AppointmentMatch match = FindMatch(outlookAppointment);

            Assert.IsNotNull(match);
            Assert.IsNotNull(match.GoogleAppointment);
            Assert.IsNotNull(match.OutlookAppointment);

            Outlook.AppointmentItem recreatedOutlookAppointment = Syncronizer.CreateOutlookAppointmentItem(Syncronizer.SyncAppointmentsFolder);
            sync.UpdateAppointment(ref match.GoogleAppointment, recreatedOutlookAppointment, match.GoogleAppointmentExceptions);
            Assert.IsNotNull(outlookAppointment);
            Assert.IsNotNull(recreatedOutlookAppointment);
            // match recreatedOutlookAppointment with outlookAppointment

            Assert.AreEqual(outlookAppointment.Subject, recreatedOutlookAppointment.Subject);

            Assert.AreEqual(outlookAppointment.Start, recreatedOutlookAppointment.Start);
            Assert.AreEqual(outlookAppointment.End, recreatedOutlookAppointment.End);
            Assert.AreEqual(outlookAppointment.AllDayEvent, recreatedOutlookAppointment.AllDayEvent);
            //ToDo: Check other properties

            DeleteTestAppointments(match);
            recreatedOutlookAppointment.Delete();
        }

        [Test]
        public void TestSync_Day()
        {
            sync.SyncOption = SyncOption.MergeOutlookWins;

            // create new appointment to sync
            Outlook.AppointmentItem outlookAppointment = Syncronizer.CreateOutlookAppointmentItem(Syncronizer.SyncAppointmentsFolder);
            outlookAppointment.Subject = name;
            outlookAppointment.Start = DateTime.Now;
            outlookAppointment.End = DateTime.Now;
            outlookAppointment.AllDayEvent = true;

            outlookAppointment.Save();


            sync.SyncOption = SyncOption.OutlookToGoogleOnly;

            var googleAppointment = new Event();
            sync.UpdateAppointment(outlookAppointment, ref googleAppointment);
           
            googleAppointment = null;

            sync.SyncOption = SyncOption.GoogleToOutlookOnly;
            //load the same appointment from google.
            MatchAppointments(sync);
            AppointmentMatch match = FindMatch(outlookAppointment);

            Assert.IsNotNull(match);
            Assert.IsNotNull(match.GoogleAppointment);
            Assert.IsNotNull(match.OutlookAppointment);

            Outlook.AppointmentItem recreatedOutlookAppointment = Syncronizer.CreateOutlookAppointmentItem(Syncronizer.SyncAppointmentsFolder);
            sync.UpdateAppointment(ref match.GoogleAppointment, recreatedOutlookAppointment, match.GoogleAppointmentExceptions);            
            Assert.IsNotNull(outlookAppointment);
            Assert.IsNotNull(recreatedOutlookAppointment);
            // match recreatedOutlookAppointment with outlookAppointment
            Assert.AreEqual(outlookAppointment.Subject, recreatedOutlookAppointment.Subject);

            Assert.AreEqual(outlookAppointment.Start, recreatedOutlookAppointment.Start);
            Assert.AreEqual(outlookAppointment.End, recreatedOutlookAppointment.End);
            Assert.AreEqual(outlookAppointment.AllDayEvent, recreatedOutlookAppointment.AllDayEvent);
            //ToDo: Check other properties

            DeleteTestAppointments(match);
            recreatedOutlookAppointment.Delete();
        }        

        [Test]
        public void TestExtendedProps()
        {
            sync.SyncOption = SyncOption.MergeOutlookWins;
        

            // create new appointment to sync
            Outlook.AppointmentItem outlookAppointment = Syncronizer.CreateOutlookAppointmentItem(Syncronizer.SyncAppointmentsFolder);
            outlookAppointment.Subject = name;
            outlookAppointment.Start = DateTime.Now;
            outlookAppointment.Start = DateTime.Now;
            outlookAppointment.AllDayEvent = true;

            outlookAppointment.Save();

            var googleAppointment = new Event();
            sync.UpdateAppointment(outlookAppointment, ref googleAppointment);
                      
            Assert.AreEqual(name, googleAppointment.Summary);

            // read appointment from google
            googleAppointment = null;
            MatchAppointments(sync);
            AppointmentsMatcher.SyncAppointments(sync);

            AppointmentMatch match = FindMatch(outlookAppointment);
            
            Assert.IsNotNull(match);
            Assert.IsNotNull(match.GoogleAppointment);

            // get extended prop
            Assert.AreEqual(AppointmentPropertiesUtils.GetOutlookId(outlookAppointment), AppointmentPropertiesUtils.GetGoogleOutlookAppointmentId(sync.SyncProfile, match.GoogleAppointment));

            DeleteTestAppointments(match);
        }

        //ToDo:
        //[Test]
        //public void TestSyncDeletedOulook()
        //{
        //    //ToDo: Check for eache SyncOption and SyncDelete combination
        //    sync.SyncOption = SyncOption.MergeOutlookWins;
        //    sync.SyncDelete = true;

        //    // create new appointment to sync
        //    Outlook.AppointmentItem outlookAppointment = Syncronizer.CreateOutlookAppointmentItem(Syncronizer.SyncAppointmentsFolder);
        //    outlookAppointment.Subject = name;
        //    outlookAppointment.Start = DateTime.Now;
        //    outlookAppointment.Start = DateTime.Now.AddHours(1); 
        //    outlookAppointment.Save();

        //    var googleAppointment = new Event();
        //    sync.UpdateAppointment(outlookAppointment, googleAppointment);           

        //    // delete outlook appointment
        //    outlookAppointment.Delete();

        //    // sync
        //    MatchAppointments(sync);
        //    AppointmentsMatcher.SyncAppointments(sync);
        //    // find match
        //    AppointmentMatch match = FindMatch(outlookAppointment);

        //    Assert.IsNotNull(match);

        //    // delete
        //    sync.UpdateAppointment(match);

        //    // sync
        //    MatchAppointments(sync);
        //    AppointmentsMatcher.SyncAppointments(sync);

        //    // check if google appointment still exists
        //    googleAppointment = null;
        //    match = sync.AppointmentByProperty(name, email);
        //    //foreach (AppointmentMatch m in sync.Appointments)
        //    //{
        //    //    if (m.GoogleAppointment.Title == name)
        //    //    {
        //    //        googleAppointment = m.GoogleAppointment;
        //    //        break;
        //    //    }
        //    //}
        //    Assert.IsNull(match);
        //}

        //[Test]
        //public void TestSyncDeletedGoogle()
        //{
        //    //ToDo: Check for eache SyncOption and SyncDelete combination
        //    sync.SyncOption = SyncOption.MergeOutlookWins;
        //    sync.SyncDelete = true;

        //    // create new appointment to sync
        //    Outlook.AppointmentItem outlookAppointment = Syncronizer.CreateOutlookAppointmentItem(Syncronizer.SyncAppointmentsFolder);
        //    outlookAppointment.FullName = name;
        //    outlookAppointment.FileAs = name;
        //    outlookAppointment.Email1Address = email;
        //    outlookAppointment.Email2Address = email.Replace("00", "01");
        //    outlookAppointment.Email3Address = email.Replace("00", "02");
        //    outlookAppointment.HomeAddress = "10 Parades";
        //    outlookAppointment.PrimaryTelephoneNumber = "123";
        //    outlookAppointment.Save();

        //    Appointment googleAppointment = new Appointment();
        //    sync.UpdateAppointment(outlookAppointment, googleAppointment);
        //    AppointmentMatch match = new AppointmentMatch(outlookAppointment, sync), googleAppointment);

        //    //save appointments
        //    sync.UpdateAppointment(match);

        //    // delete google appointment
        //    GoogleAppointment.Delete();

        //    // sync
        //    MatchAppointments(sync);
        //    match = sync.AppointmentByProperty(name, email);
        //    AppointmentsMatcher.SyncAppointment(match, sync);

        //    // delete
        //    sync.UpdateAppointment(match);

        //    // sync
        //    MatchAppointments(sync);
        //    AppointmentsMatcher.SyncAppointments(sync);
        //    match = sync.AppointmentByProperty(name, email);

        //    // check if outlook appointment still exists
        //    Assert.IsNull(match);

        //    DeleteTestAppointments(match);
        //}

        
        //[Test]
        //public void TestSyncGroups()
        //{
        //    sync.SyncOption = SyncOption.MergeOutlookWins;

        //    // create new appointment to sync
        //    Outlook.AppointmentItem outlookAppointment = Syncronizer.CreateOutlookAppointmentItem(Syncronizer.SyncAppointmentsFolder);
        //    outlookAppointment.FullName = name;
        //    outlookAppointment.FileAs = name;
        //    outlookAppointment.Email1Address = email;
        //    outlookAppointment.Email2Address = email.Replace("00", "01");
        //    outlookAppointment.Email3Address = email.Replace("00", "02");
        //    outlookAppointment.HomeAddress = "10 Parades";
        //    outlookAppointment.PrimaryTelephoneNumber = "123";
        //    outlookAppointment.Categories = groupName;
        //    outlookAppointment.Save();

        //    //Outlook appointment should now have a group
        //    Assert.AreEqual(groupName, outlookAppointment.Categories);

        //    //Sync Groups first
        //    MatchAppointments(sync);
        //    AppointmentsMatcher.SyncGroups(sync);

        //    Appointment googleAppointment = new Appointment();
        //    sync.UpdateAppointment(outlookAppointment, googleAppointment);
        //    AppointmentMatch match = new AppointmentMatch(outlookAppointment, sync), googleAppointment);

        //    //sync and save appointment to google.
        //    AppointmentsMatcher.SyncAppointment(match, sync);
        //    sync.UpdateAppointment(match);

        //    //load the same appointment from google.
        //    MatchAppointments(sync);
        //    match = sync.AppointmentByProperty(name, email);
        //    AppointmentsMatcher.SyncAppointment(match, sync);

        //    // google appointment should now have the same group
        //    System.Collections.ObjectModel.Collection<Group> googleGroups = Utilities.GetGoogleGroups(sync, match.GoogleAppointment);
        //    Assert.AreEqual(2, googleGroups.Count);
        //    Assert.Contains(sync.GetGoogleGroupByName(groupName), googleGroups);
        //    Assert.Contains(sync.GetGoogleGroupByName(Syncronizer.myAppointmentsGroup), googleGroups);

        //    // delete outlook appointment
        //    outlookAppointment.Delete();
        //    outlookAppointment = Syncronizer.CreateOutlookAppointmentItem(Syncronizer.SyncAppointmentsFolder);
        //    sync.UpdateAppointment(match.GoogleAppointment, outlookAppointment);
        //    match = new AppointmentMatch(outlookAppointment, sync), match.GoogleAppointment);
        //    outlookAppointment.Save();

        //    sync.SyncOption = SyncOption.MergeGoogleWins;

        //    //sync and save appointment to outlook
        //    AppointmentsMatcher.SyncAppointment(match, sync);
        //    sync.UpdateAppointment(match);

        //    //load the same appointment from outlook
        //    MatchAppointments(sync);
        //    match = sync.AppointmentByProperty(name, email);
        //    AppointmentsMatcher.SyncAppointment(match, sync);


        //    Assert.AreEqual(groupName, outlookAppointment.Categories);

        //    DeleteTestAppointments(match);

        //    // delete test group
        //    Group group = sync.GetGoogleGroupByName(groupName);
        //    if (group != null)
        //        sync.CalendarService.Delete(group);
        //}

        //[Test]
        //public void TestSyncDeletedGoogleGroup()
        //{
        //    //ToDo: Check for eache SyncOption and SyncDelete combination
        //    sync.SyncOption = SyncOption.MergeOutlookWins;
        //    sync.SyncDelete = true;

        //    // create new appointment to sync
        //    Outlook.AppointmentItem outlookAppointment = Syncronizer.CreateOutlookAppointmentItem(Syncronizer.SyncAppointmentsFolder);
        //    outlookAppointment.FullName = name;
        //    outlookAppointment.FileAs = name;
        //    outlookAppointment.Email1Address = email;
        //    outlookAppointment.Email2Address = email.Replace("00", "01");
        //    outlookAppointment.Email3Address = email.Replace("00", "02");
        //    outlookAppointment.HomeAddress = "10 Parades";
        //    outlookAppointment.PrimaryTelephoneNumber = "123";
        //    outlookAppointment.Categories = groupName;
        //    outlookAppointment.Save();

        //    //Outlook appointment should now have a group
        //    Assert.AreEqual(groupName, outlookAppointment.Categories);

        //    //Sync Groups first
        //    MatchAppointments(sync);
        //    AppointmentsMatcher.SyncGroups(sync);

        //    //Create now Google Appointment and assing new Group
        //    Appointment googleAppointment = new Appointment();
        //    sync.UpdateAppointment(outlookAppointment, googleAppointment);
        //    AppointmentMatch match = new AppointmentMatch(outlookAppointment, sync), googleAppointment);

        //    //save appointment to google.            
        //    sync.UpdateAppointment(match);

        //    //load the same appointment from google.
        //    MatchAppointments(sync);
        //    match = sync.AppointmentByProperty(name, email);
        //    AppointmentsMatcher.SyncAppointment(match, sync);

        //    // google appointment should now have the same group
        //    System.Collections.ObjectModel.Collection<Group> googleGroups = Utilities.GetGoogleGroups(sync, match.GoogleAppointment);
        //    Group group = sync.GetGoogleGroupByName(groupName);

        //    Assert.AreEqual(2, googleGroups.Count);
        //    Assert.Contains(group, googleGroups);
        //    Assert.Contains(sync.GetGoogleGroupByName(Syncronizer.myAppointmentsGroup), googleGroups);

        //    // delete group from google
        //    Utilities.RemoveGoogleGroup(match.GoogleAppointment, group);

        //    googleGroups = Utilities.GetGoogleGroups(sync, match.GoogleAppointment);
        //    Assert.AreEqual(1, googleGroups.Count);
        //    Assert.Contains(sync.GetGoogleGroupByName(Syncronizer.myAppointmentsGroup), googleGroups);

        //    //save appointment to google.
        //    sync.SaveGoogleAppointment(match.GoogleAppointment);

        //    sync.SyncOption = SyncOption.GoogleToOutlookOnly;

        //    //Sync Groups first
        //    MatchAppointments(sync);
        //    AppointmentsMatcher.SyncGroups(sync);

        //    //sync and save appointment to outlook.
        //    match = sync.AppointmentByProperty(name, email);
        //    sync.UpdateAppointment(match.GoogleAppointment, outlookAppointment);
        //    sync.UpdateAppointment(match);

        //    // google and outlook should now have no category except for the System Group: My Appointments
        //    googleGroups = Utilities.GetGoogleGroups(sync, match.GoogleAppointment);
        //    Assert.AreEqual(1, googleGroups.Count);
        //    Assert.AreEqual(null, outlookAppointment.Categories);
        //    Assert.Contains(sync.GetGoogleGroupByName(Syncronizer.myAppointmentsGroup), googleGroups);

        //    DeleteTestAppointments(match);

        //    // delete test group
        //    if (group != null)
        //        sync.CalendarService.Delete(group);
        //}

        //[Test]
        //public void TestSyncDeletedOutlookGroup()
        //{
        //    //ToDo: Check for eache SyncOption and SyncDelete combination
        //    sync.SyncOption = SyncOption.MergeOutlookWins;
        //    sync.SyncDelete = true;

        //    // create new appointment to sync
        //    Outlook.AppointmentItem outlookAppointment = Syncronizer.CreateOutlookAppointmentItem(Syncronizer.SyncAppointmentsFolder);
        //    outlookAppointment.FullName = name;
        //    outlookAppointment.FileAs = name;
        //    outlookAppointment.Email1Address = email;
        //    outlookAppointment.Email2Address = email.Replace("00", "01");
        //    outlookAppointment.Email3Address = email.Replace("00", "02");
        //    outlookAppointment.HomeAddress = "10 Parades";
        //    outlookAppointment.PrimaryTelephoneNumber = "123";
        //    outlookAppointment.Categories = groupName;
        //    outlookAppointment.Save();

        //    //Outlook appointment should now have a group
        //    Assert.AreEqual(groupName, outlookAppointment.Categories);

        //    //Now sync Groups
        //    MatchAppointments(sync);
        //    AppointmentsMatcher.SyncGroups(sync);

        //    Appointment googleAppointment = new Appointment();
        //    sync.UpdateAppointment(outlookAppointment, googleAppointment);
        //    AppointmentMatch match = new AppointmentMatch(outlookAppointment, sync), googleAppointment);

        //    //save appointment to google.
        //    sync.UpdateAppointment(match);

        //    //load the same appointment from google.
        //    MatchAppointments(sync);
        //    match = sync.AppointmentByProperty(name, email);
        //    AppointmentsMatcher.SyncAppointment(match, sync);

        //    // google appointment should now have the same group
        //    System.Collections.ObjectModel.Collection<Group> googleGroups = Utilities.GetGoogleGroups(sync, match.GoogleAppointment);
        //    Group group = sync.GetGoogleGroupByName(groupName);
        //    Assert.AreEqual(2, googleGroups.Count);
        //    Assert.Contains(sync.GetGoogleGroupByName(Syncronizer.myAppointmentsGroup), googleGroups);
        //    Assert.Contains(group, googleGroups);

        //    // delete group from outlook
        //    Utilities.RemoveOutlookGroup(outlookAppointment, groupName);

        //    //save appointment to google.
        //    sync.UpdateAppointment(match);

        //    //load the same appointment from google.
        //    MatchAppointments(sync);
        //    match = sync.AppointmentByProperty(name, email);
        //    sync.UpdateAppointment(outlookAppointment, match.GoogleAppointment);

        //    // google and outlook should now have no category
        //    googleGroups = Utilities.GetGoogleGroups(sync, match.GoogleAppointment);
        //    Assert.AreEqual(null, outlookAppointment.Categories);
        //    Assert.AreEqual(1, googleGroups.Count);
        //    Assert.Contains(sync.GetGoogleGroupByName(Syncronizer.myAppointmentsGroup), googleGroups);

        //    DeleteTestAppointments(match);

        //    // delete test group
        //    if (group != null)
        //        sync.CalendarService.Delete(group);
        //}

        //[Test]
        //public void TestResetMatches()
        //{
        //    sync.SyncOption = SyncOption.MergeOutlookWins;

        //    // create new appointment to sync
        //    Outlook.AppointmentItem outlookAppointment = Syncronizer.CreateOutlookAppointmentItem(Syncronizer.SyncAppointmentsFolder);
        //    outlookAppointment.FullName = name;
        //    outlookAppointment.FileAs = name;
        //    outlookAppointment.Email1Address = email;
        //    outlookAppointment.Email2Address = email.Replace("00", "01");
        //    outlookAppointment.Email3Address = email.Replace("00", "02");
        //    outlookAppointment.HomeAddress = "10 Parades";
        //    outlookAppointment.PrimaryTelephoneNumber = "123";
        //    //outlookAppointment.Categories = groupName; //Group is not relevant here
        //    outlookAppointment.Save();

        //    Appointment googleAppointment = new Appointment();
        //    sync.UpdateAppointment(outlookAppointment, googleAppointment);
        //    AppointmentMatch match = new AppointmentMatch(outlookAppointment, sync), googleAppointment);

        //    //save appointment to google.
        //    sync.UpdateAppointment(match);

        //    //load the same appointment from google.
        //    MatchAppointments(sync);
        //    match = sync.AppointmentByProperty(name, email);
        //    AppointmentsMatcher.SyncAppointment(match, sync);

        //    // delete outlook appointment
        //    outlookAppointment.Delete();
        //    match.OutlookAppointment = null;

        //    //load the same appointment from google
        //    MatchAppointments(sync);
        //    match = sync.AppointmentByProperty(name, email);
        //    AppointmentsMatcher.SyncAppointment(match, sync);

        //    Assert.IsNull(match.OutlookAppointment);

        //    // reset matches
        //    sync.ResetMatch(match.GoogleAppointment);
        //    //Not, because NULL: sync.ResetMatch(match.OutlookAppointment.GetOriginalItemFromOutlook(sync));

        //    // load same appointment match
        //    MatchAppointments(sync);
        //    match = sync.AppointmentByProperty(name, email);
        //    AppointmentsMatcher.SyncAppointment(match, sync);

        //    // google appointment should still be present and OutlookAppointment should be filled
        //    Assert.IsNotNull(match.GoogleAppointment);
        //    Assert.IsNotNull(match.OutlookAppointment);

        //    DeleteTestAppointments();

        //    // create new appointment to sync
        //    outlookAppointment = Syncronizer.CreateOutlookAppointmentItem(Syncronizer.SyncAppointmentsFolder);
        //    outlookAppointment.FullName = name;
        //    outlookAppointment.FileAs = name;
        //    outlookAppointment.Email1Address = email;
        //    outlookAppointment.Email2Address = email.Replace("00", "01");
        //    outlookAppointment.Email3Address = email.Replace("00", "02");
        //    outlookAppointment.HomeAddress = "10 Parades";
        //    outlookAppointment.PrimaryTelephoneNumber = "123";
        //    outlookAppointment.Save();

        //    // same test for delete google appointment...
        //    googleAppointment = new Appointment();
        //    sync.UpdateAppointment(outlookAppointment, googleAppointment);
        //    match = new AppointmentMatch(outlookAppointment, sync), googleAppointment);

        //    //save appointment to google.
        //    sync.UpdateAppointment(match);

        //    //load the same appointment from google.
        //    MatchAppointments(sync);
        //    match = sync.AppointmentByProperty(name, email);
        //    AppointmentsMatcher.SyncAppointment(match, sync);

        //    // delete google appointment           
        //    match.GoogleAppointment.Delete();
        //    match.GoogleAppointment = null;

        //    //load the same appointment from google.
        //    MatchAppointments(sync);
        //    match = sync.AppointmentByProperty(name, email);
        //    AppointmentsMatcher.SyncAppointment(match, sync);

        //    Assert.IsNull(match.GoogleAppointment);

        //    // reset matches
        //    //Not, because null: sync.ResetMatch(match.GoogleAppointment);
        //    sync.ResetMatch(match.OutlookAppointment.GetOriginalItemFromOutlook());

        //    // load same appointment match
        //    MatchAppointments(sync);
        //    match = sync.AppointmentByProperty(name, email);
        //    AppointmentsMatcher.SyncAppointment(match, sync);

        //    // Outlook appointment should still be present and GoogleAppointment should be filled
        //    Assert.IsNotNull(match.OutlookAppointment);
        //    Assert.IsNotNull(match.GoogleAppointment);

        //    outlookAppointment.Delete();
        //}

        private void DeleteTestAppointments(AppointmentMatch match)
        {
            if (match != null)
            {
                DeleteTestAppointment(match.GoogleAppointment);
                DeleteTestAppointment(match.OutlookAppointment);
            }
        }

        private void DeleteTestAppointment(Outlook.AppointmentItem outlookAppointment)
        {
            if (outlookAppointment != null)
            {
                try
                {
                    string name = outlookAppointment.Subject;
                    outlookAppointment.Delete();
                    Logger.Log("Deleted Outlook test appointment: " + name, EventType.Information);
                }
                finally
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(outlookAppointment);
                    outlookAppointment = null;
                }

            }
        }


       

        private void DeleteTestAppointment(Event googleAppointment)
        {
            if (googleAppointment != null && !googleAppointment.Status.Equals("cancelled"))
            {
                sync.EventRequest.Delete(sync.PrimaryCalendar.Id, googleAppointment.Id);
                Logger.Log("Deleted Google test appointment: " + googleAppointment.Summary, EventType.Information);
                Thread.Sleep(2000);
            }
        }
        

        internal AppointmentMatch FindMatch(Outlook.AppointmentItem outlookAppointment)
        {
            foreach (AppointmentMatch match in sync.Appointments)
            {
                if (match.OutlookAppointment != null && match.OutlookAppointment.EntryID == outlookAppointment.EntryID)
                    return match;
            }
            return null;
        }

        private void MatchAppointments(Syncronizer sync)
        {
            Thread.Sleep(5000); //Wait, until Appointment is really saved and available to retrieve again
            sync.MatchAppointments();
        }

        internal AppointmentMatch FindMatch(Event googleAppointment)
        {
            if (googleAppointment != null)
            {
                foreach (AppointmentMatch match in sync.Appointments)
                {
                    if (match.GoogleAppointment != null && match.GoogleAppointment.Id == googleAppointment.Id)
                        return match;
                }
            }
            return null;
        }

    }
}
