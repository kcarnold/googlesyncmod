using System;
using System.Collections.Generic;
using System.Text;
using NUnit.Framework;
using Google.GData.Contacts;
using Google.GData.Client;
using Google.GData.Extensions;
using System.Configuration;
using Google.Contacts;
using Google.Documents;
using Google.GData.Client.ResumableUpload;
using Google.GData.Documents;
using System.Collections;
using Google.Apis.Calendar.v3;
using Google.Apis.Util.Store;
using Google.Apis.Auth.OAuth2;
using System.Threading;
using System.IO;
using Google.Apis.Calendar.v3.Data;

namespace GoContactSyncMod.UnitTests
{
    [TestFixture]
    public class GoogleAPITests
    {
        OAuth2Authenticator _authenticator;
        static Logger.LogUpdatedHandler _logUpdateHandler = null;
        void Logger_LogUpdated(string message)
        {
            Console.WriteLine(message);
        }

        [TestFixtureSetUp]
        public void Init()
        {
            //string timestamp = DateTime.Now.Ticks.ToString();            
            if (_logUpdateHandler == null)
            {
                _logUpdateHandler = new Logger.LogUpdatedHandler(Logger_LogUpdated);
                Logger.LogUpdated += _logUpdateHandler;
            }
        }    

        [Test]
        public void CreateNewContact()
        {
            string gmailUsername;
            string syncProfile;
            GoogleAPITests.LoadSettings(out gmailUsername, out syncProfile);

            ContactsRequest service;

            var scopes = new List<string>();
            //Contacts-Scope
            scopes.Add("https://www.google.com/m8/feeds");
            //Notes-Scope
            scopes.Add("https://docs.google.com/feeds/");
            //scopes.Add("https://docs.googleusercontent.com/");
            //scopes.Add("https://spreadsheets.google.com/feeds/");
            //Calendar-Scope
            //scopes.Add("https://www.googleapis.com/auth/calendar");
            scopes.Add(CalendarService.Scope.Calendar);

            UserCredential credential;
            byte[] jsonSecrets = Properties.Resources.client_secrets;
            
            using (var stream = new MemoryStream(jsonSecrets))
            {
                FileDataStore fDS = new FileDataStore(Logger.AuthFolder, true);

                GoogleClientSecrets clientSecrets = GoogleClientSecrets.Load(stream);

                credential = GCSMOAuth2WebAuthorizationBroker.AuthorizeAsync(
                                clientSecrets.Secrets,
                                scopes.ToArray(),
                                gmailUsername,
                                CancellationToken.None,
                                fDS).
                                Result;
                
                OAuth2Parameters parameters = new OAuth2Parameters
                {
                    ClientId = clientSecrets.Secrets.ClientId,
                    ClientSecret = clientSecrets.Secrets.ClientSecret,

                    // Note: AccessToken is valid only for 60 minutes
                    AccessToken = credential.Token.AccessToken,
                    RefreshToken = credential.Token.RefreshToken
                };
               
                RequestSettings settings = new RequestSettings("GoContactSyncMod", parameters);

                service = new ContactsRequest(settings);
            }

            


            #region Delete previously created test contact.
            ContactsQuery query = new ContactsQuery(ContactsQuery.CreateContactsUri("default"));
            query.NumberToRetrieve = 500;

            Feed<Contact> feed = service.Get<Contact>(query);

            Logger.Log("Loaded Google contacts", EventType.Information);

            foreach (Contact entry in feed.Entries)
            {
                if (entry.PrimaryEmail != null && entry.PrimaryEmail.Address == "johndoe@example.com")
                {
                    service.Delete(entry);
                    Logger.Log("Deleted Google contact", EventType.Information);
                    //break;
                }
            }
            #endregion

            Contact newEntry = new Contact();
            newEntry.Title = "John Doe";

            EMail primaryEmail = new EMail("johndoe@example.com");
            primaryEmail.Primary = true;
            primaryEmail.Rel = ContactsRelationships.IsWork;
            newEntry.Emails.Add(primaryEmail);

            PhoneNumber phoneNumber = new PhoneNumber("555-555-5551");
            phoneNumber.Primary = true;
            phoneNumber.Rel = ContactsRelationships.IsMobile;
            newEntry.Phonenumbers.Add(phoneNumber);

            StructuredPostalAddress postalAddress = new StructuredPostalAddress();
            postalAddress.Street = "123 somewhere lane";
            postalAddress.Primary = true;
            postalAddress.Rel = ContactsRelationships.IsHome;
            newEntry.PostalAddresses.Add(postalAddress);

            newEntry.Content = "Who is this guy?";

            Uri feedUri = new Uri(ContactsQuery.CreateContactsUri("default"));

            Contact createdEntry = service.Insert(feedUri, newEntry);

            Logger.Log("Created Google contact", EventType.Information);

            Assert.IsNotNull(createdEntry.ContactEntry.Id.Uri);

            Contact updatedEntry = service.Update(createdEntry);

            Logger.Log("Updated Google contact", EventType.Information);

            //delete test contacts
            service.Delete(createdEntry);

            Logger.Log("Deleted Google contact", EventType.Information);
        }

        [Test]
        public void CreateNewAppointment()
        {
            string gmailUsername;
            string syncProfile;
            GoogleAPITests.LoadSettings(out gmailUsername, out syncProfile);

            EventsResource service;
            CalendarListEntry primaryCalendar = null;
            var scopes = new List<string>();
            //Contacts-Scope
            scopes.Add("https://www.google.com/m8/feeds");
            //Notes-Scope
            scopes.Add("https://docs.google.com/feeds/");
            //scopes.Add("https://docs.googleusercontent.com/");
            //scopes.Add("https://spreadsheets.google.com/feeds/");
            //Calendar-Scope
            //scopes.Add("https://www.googleapis.com/auth/calendar");
            scopes.Add(CalendarService.Scope.Calendar);

            UserCredential credential;
            byte[] jsonSecrets = Properties.Resources.client_secrets;
            
            //using (var stream = new FileStream(Path.GetDirectoryName(System.Reflection.Assembly.GetAssembly(this.GetType()).Location) + "\\client_secrets.json", FileMode.Open, FileAccess.Read))
            //using (var stream = new FileStream("client_secrets.json", FileMode.Open, FileAccess.Read))
                    //using (var stream = new FileStream(Application.StartupPath + "\\client_secrets.json", FileMode.Open, FileAccess.Read))
            using (var stream = new MemoryStream(jsonSecrets))            
            {
                FileDataStore fDS = new FileDataStore(Logger.AuthFolder, true);
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                GoogleClientSecrets.Load(stream).Secrets, scopes, gmailUsername, CancellationToken.None,
                fDS).Result;

                var initializer = new Google.Apis.Services.BaseClientService.Initializer();
                initializer.HttpClientInitializer = credential;
                var CalendarRequest = new CalendarService(initializer);
                //CalendarRequest.setUserCredentials(username, password);

                var list = CalendarRequest.CalendarList.List().Execute().Items;
                foreach (var calendar in list)
                {
                    if (calendar.Primary != null && calendar.Primary.Value)
                    {
                        primaryCalendar = calendar;
                        break;
                    }
                }

                if (primaryCalendar == null)
                    throw new Exception("Primary Calendar not found");


                //EventQuery query = new EventQuery("https://www.google.com/calendar/feeds/default/private/full");
                //ToDo: Upgrade to v3, EventQuery query = new EventQuery("https://www.googleapis.com/calendar/v3/calendars/default/events");
                service = CalendarRequest.Events;
            }

            #region Delete previously created test contact.
            var query = service.List(primaryCalendar.Id);
            query.MaxResults = 500;
            query.TimeMin = DateTime.Now.AddDays(-10);
            query.TimeMax = DateTime.Now.AddDays(10);
            //query.Q = "GCSM Test Appointment";

            var feed = query.Execute();
            Logger.Log("Loaded Google appointments", EventType.Information);
            foreach (Google.Apis.Calendar.v3.Data.Event entry in feed.Items)
            {
                if (entry.Summary != null && entry.Summary.Contains("GCSM Test Appointment") && !entry.Status.Equals("cancelled"))
                {
                    Logger.Log("Deleting Google appointment:" +entry.Summary + " - " + entry.Start.DateTime.ToString(), EventType.Information);
                    service.Delete(primaryCalendar.Id, entry.Id);
                    Logger.Log("Deleted Google appointment", EventType.Information);
                    //break;
                }
            }

                   
            
            #endregion

            var newEntry = Factory.NewEvent();
            newEntry.Summary = "GCSM Test Appointment";
            newEntry.Start.DateTime = DateTime.Now;
            newEntry.End.DateTime = DateTime.Now;                                    

            var createdEntry = service.Insert(newEntry,primaryCalendar.Id).Execute();

            Logger.Log("Created Google appointment", EventType.Information);

            Assert.IsNotNull(createdEntry.Id);

            var updatedEntry = service.Update(createdEntry, primaryCalendar.Id,createdEntry.Id).Execute();

            Logger.Log("Updated Google appointment", EventType.Information);

            //delete test contacts
            service.Delete(primaryCalendar.Id, updatedEntry.Id).Execute();

            Logger.Log("Deleted Google appointment", EventType.Information);
        }


        [Test]
        public void CreateNewNote()
        {
            string gmailUsername;
            string syncProfile;
            GoogleAPITests.LoadSettings(out gmailUsername, out syncProfile);

            DocumentsRequest service;

            var scopes = new List<string>();
            //Contacts-Scope
            scopes.Add("https://www.google.com/m8/feeds");
            //Notes-Scope
            scopes.Add("https://docs.google.com/feeds/");
            //scopes.Add("https://docs.googleusercontent.com/");
            //scopes.Add("https://spreadsheets.google.com/feeds/");
            //Calendar-Scope
            //scopes.Add("https://www.googleapis.com/auth/calendar");
            scopes.Add(CalendarService.Scope.Calendar);

            UserCredential credential;
            byte[] jsonSecrets = Properties.Resources.client_secrets;

            using (var stream = new MemoryStream(jsonSecrets))
            {
                FileDataStore fDS = new FileDataStore(Logger.AuthFolder, true);

                GoogleClientSecrets clientSecrets = GoogleClientSecrets.Load(stream);

                credential = GCSMOAuth2WebAuthorizationBroker.AuthorizeAsync(
                                clientSecrets.Secrets,
                                scopes.ToArray(),
                                gmailUsername,
                                CancellationToken.None,
                                fDS).
                                Result;

                OAuth2Parameters parameters = new OAuth2Parameters
                {
                    ClientId = clientSecrets.Secrets.ClientId,
                    ClientSecret = clientSecrets.Secrets.ClientSecret,

                    // Note: AccessToken is valid only for 60 minutes
                    AccessToken = credential.Token.AccessToken,
                    RefreshToken = credential.Token.RefreshToken
                };

                RequestSettings settings = new RequestSettings("GoContactSyncMod", parameters);

                service = new DocumentsRequest(settings);

                //Instantiate an Authenticator object according to your authentication, to use ResumableUploader
                _authenticator = new OAuth2Authenticator("GCSM Unit Tests", parameters);
            }

            

            //Delete previously created test note.            
            DeleteTestNote(service);

            Document newEntry = new Document();
            newEntry.Type = Document.DocumentType.Document;
            newEntry.Title = "AN_OUTLOOK_TEST_NOTE";

            string file = NotePropertiesUtils.CreateNoteFile("AN_OUTLOOK_TEST_NOTE", "This is just a test note to test GoContactSyncMod", null);
            newEntry.MediaSource = new MediaFileSource(file, MediaFileSource.GetContentTypeForFileName(file));

            #region normal flow, only working to create documents without content (metadata only), e.g. for Notes folder

            Document createdEntry = Synchronizer.SaveGoogleNote(null, newEntry, service);

            Assert.IsNotNull(createdEntry.DocumentEntry.Id.Uri);

            Logger.Log("Created Google note", EventType.Information);

            //Wait 5 seconds to give the testcase the chance to finish
            System.Threading.Thread.Sleep(5000);

            //delete test note            
            DeleteTestNote(service);
            #endregion

            #region workaround flow to use UploadDocument, not needed anymore because of new approach to use ResumableUploader
            //Google.GData.Documents.DocumentEntry createdEntry2 = service.Service.UploadDocument(file, newEntry.Title);
            
            //Assert.IsNotNull(createdEntry2.Id.Uri);

            ////delete test note            
            //DeleteTestNote(service);
            #endregion

            #region New approach how to update an existing document: https://developers.google.com/google-apps/documents-list/#updatingchanging_documents_and_files            
            //Instantiate the ResumableUploader component.      
            ResumableUploader uploader = new ResumableUploader();            
            uploader.AsyncOperationCompleted += new AsyncOperationCompletedEventHandler(OnGoogleNoteCreated);
            Synchronizer.CreateGoogleNote(newEntry, file, service, uploader, _authenticator);            
            #endregion            

            //Wait 5 seconds to give the testcase the chance to finish the Async events
            System.Threading.Thread.Sleep(5000);

            DeleteTestNote(service);
        }

        private void OnGoogleNoteCreated(object sender, AsyncOperationCompletedEventArgs e)
        {
            DocumentEntry entry = e.Entry as DocumentEntry;

            Assert.IsNotNull(entry);

            Logger.Log("Created Google note", EventType.Information);

            //Now update the same entry
            //Instantiate the ResumableUploader component.      
            ResumableUploader uploader = new ResumableUploader();
            uploader.AsyncOperationCompleted += new AsyncOperationCompletedEventHandler(OnGoogleNoteUpdated);
            uploader.UpdateAsync(_authenticator, entry, e.UserState);


        }

        private void OnGoogleNoteUpdated(object sender, AsyncOperationCompletedEventArgs e)
        {
            DocumentEntry entry = e.Entry as DocumentEntry;
            
            Assert.IsNotNull(entry);

            Logger.Log("Updated Google note", EventType.Information);
            
            System.IO.File.Delete(e.UserState as string);
        }

        private static void DeleteTestNote(DocumentsRequest service)
        {
            //ToDo: Doesn'T work always, frequently throwing 401, Precondition failed, maybe Google API bug
            //service.Delete(createdEntry);

            //Todo: Workaround to load document again
            DocumentQuery query = new DocumentQuery(service.BaseUri);
            query.NumberToRetrieve = 500;

            Feed<Document> feed = service.Get<Document>(query);

            Logger.Log("Loaded Google notes", EventType.Information);

            foreach (Document entry in feed.Entries)
            {
                if (entry.Title == "AN_OUTLOOK_TEST_NOTE")
                {
                    //service.Delete(entry);
                    service.Delete(new Uri(Google.GData.Documents.DocumentsListQuery.documentsBaseUri + "/" + entry.ResourceId), entry.ETag);
                    Logger.Log("Deleted Google note", EventType.Information);
                    //break;
                }
            }
        }

        internal static void LoadSettings(out string gmailUsername, out string syncProfile, out string syncContactsFolder, out string syncNotesFolder, out string syncAppointmentsFolder)
        {
            Microsoft.Win32.RegistryKey regKeyAppRoot = LoadSettings(out gmailUsername, out syncProfile);

            syncContactsFolder = "";
            syncNotesFolder = "";
            syncAppointmentsFolder = "";
            Synchronizer.SyncAppointmentsGoogleFolder = "";

            //First, check if there is a folder called GCSMTestContacts and GCSMTestNotes available, if yes, use them
            ArrayList outlookContactFolders = new ArrayList();
            ArrayList outlookNoteFolders = new ArrayList();
            ArrayList outlookAppointmentFolders = new ArrayList();
            Microsoft.Office.Interop.Outlook.Folders folders = Synchronizer.OutlookNameSpace.Folders;
            foreach (Microsoft.Office.Interop.Outlook.Folder folder in folders)
            {
                try
                {
                    SettingsForm.GetOutlookMAPIFolders(outlookContactFolders, outlookNoteFolders, outlookAppointmentFolders, folder);
                }
                catch (Exception e)
                {
                    Logger.Log("Error getting available Outlook folders: " + e.Message, EventType.Warning);
                }
            }

            foreach (OutlookFolder folder in outlookContactFolders)
            {
                if (folder.FolderName.ToUpper().Contains("GCSMTestContacts".ToUpper()))
                {
                    Logger.Log("Uses Test folder: " + folder.DisplayName, EventType.Information);
                    syncContactsFolder = folder.FolderID;
                    break;
                }
            }
            foreach (OutlookFolder folder in outlookNoteFolders)
            {
                if (folder.FolderName.ToUpper().Contains("GCSMTestNotes".ToUpper()))
                {
                    Logger.Log("Uses Test folder: " + folder.DisplayName, EventType.Information);
                    syncNotesFolder = folder.FolderID;
                    break;
                }
            }

            foreach (OutlookFolder folder in outlookAppointmentFolders)
            {
                if (folder.FolderName.ToUpper().Contains("GCSMTestAppointments".ToUpper()))
                {
                    Logger.Log("Uses Test folder: " + folder.DisplayName, EventType.Information);
                    syncAppointmentsFolder = folder.FolderID;
                    break;
                }
            }

            if (string.IsNullOrEmpty(syncContactsFolder))
                if (regKeyAppRoot.GetValue("SyncContactsFolder") != null)
                    syncContactsFolder = regKeyAppRoot.GetValue("SyncContactsFolder") as string;
            if (string.IsNullOrEmpty(syncNotesFolder))
                if (regKeyAppRoot.GetValue("SyncNotesFolder") != null)
                    syncNotesFolder = regKeyAppRoot.GetValue("SyncNotesFolder") as string;
            if (string.IsNullOrEmpty(syncAppointmentsFolder))
                if (regKeyAppRoot.GetValue("SyncAppointmentsFolder") != null)
                    syncAppointmentsFolder = regKeyAppRoot.GetValue("SyncAppointmentsFolder") as string;
            if (string.IsNullOrEmpty(Synchronizer.SyncAppointmentsGoogleFolder))
                if (regKeyAppRoot.GetValue("SyncAppointmentsGoogleFolder") != null)
                    Synchronizer.SyncAppointmentsGoogleFolder = regKeyAppRoot.GetValue("SyncAppointmentsGoogeFolder") as string;           
        }

        private static Microsoft.Win32.RegistryKey LoadSettings(out string gmailUsername, out string syncProfile)
        { 
            //sync.LoginToGoogle(ConfigurationManager.AppSettings["Gmail.Username"],  ConfigurationManager.AppSettings["Gmail.Password"]);
            //ToDo: Reading the username and config from the App.Config file doesn't work. If it works, consider special characters like & = &amp; in the XML file
            //ToDo: Maybe add a common Test account to the App.config and read it from there, encrypt the password
            //For now, read the userName from the Registry (same settings as for GoogleContactsSync Application
            gmailUsername = "";

            const string appRootKey = SettingsForm.AppRootKey;
            Microsoft.Win32.RegistryKey regKeyAppRoot = regKeyAppRoot = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(appRootKey);
            syncProfile = "Default Profile";
            if (regKeyAppRoot.GetValue("SyncProfile") != null)
                syncProfile = regKeyAppRoot.GetValue("SyncProfile") as string;

            regKeyAppRoot = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(appRootKey + (syncProfile != null ? ('\\' + syncProfile) : ""));

            if (regKeyAppRoot.GetValue("Username") != null)
            {
                gmailUsername = regKeyAppRoot.GetValue("Username") as string;
                
            }

            return regKeyAppRoot;
        }
    }

}
