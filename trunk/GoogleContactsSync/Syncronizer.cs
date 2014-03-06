using System;
using System.Collections.Generic;
using System.Diagnostics;
using Google.GData.Contacts;
using Google.GData.Client;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Drawing;
using System.IO;
using System.Collections.ObjectModel;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Google.Contacts;
using Google.Documents;
using Google.GData.Client.ResumableUpload;
using Google.GData.Documents;
using System.Threading;
using Google.GData.Calendar;

namespace GoContactSyncMod
{
	internal class Syncronizer
	{
		public const int OutlookUserPropertyMaxLength = 32;
		public const string OutlookUserPropertyTemplate = "g/con/{0}/";
        internal const string myContactsGroup = "System Group: My Contacts";
		private static object _syncRoot = new object();

        public int TotalCount { get; private set; }
		public int SyncedCount { get; private set; }
        public int DeletedCount { get; private set; }		
        public int ErrorCount { get; private set; }		
        public int SkippedCount { get; set; }		
        public int SkippedCountNotMatches { get; set; }		
        public ConflictResolution ConflictResolution { get; set; }
        public DeleteResolution DeleteGoogleResolution { get; set; }
        public DeleteResolution DeleteOutlookResolution { get; set; }


		public delegate void DuplicatesFoundHandler(string title, string message);
		public delegate void ErrorNotificationHandler(string title, Exception ex, EventType eventType);
		public event DuplicatesFoundHandler DuplicatesFound;
		public event ErrorNotificationHandler ErrorEncountered;

        public ContactsRequest ContactsRequest { get; private set; }

        private ClientLoginAuthenticator authenticator;
        public DocumentsRequest DocumentsRequest { get; private set; }
        public CalendarService CalendarService { get; private set; }

		private static Outlook.NameSpace _outlookNamespace;
        public static Outlook.NameSpace OutlookNameSpace
        {
            get {
                //Just create outlook instance again, in case the namespace is null
                CreateOutlookInstance();
                return _outlookNamespace; 
            }
        }

        public static Outlook.Application OutlookApplication { get; private set; }
        public Outlook.Items OutlookContacts { get; private set; }
        public Outlook.Items OutlookNotes { get; private set; }
        public Outlook.Items OutlookAppointments { get; private set; }
        public Collection<ContactMatch> OutlookContactDuplicates { get; set; }
        public Collection<ContactMatch> GoogleContactDuplicates { get; set; }
        public Collection<Contact> GoogleContacts { get; private set; }
        public Collection<Document> GoogleNotes { get; private set; }
        public Collection<EventEntry> GoogleAppointments { get; private set; }
        public Collection<Group> GoogleGroups { get; set; }
        internal Document googleNotesFolder;
        public string OutlookPropertyPrefix { get; private set; }

		public string OutlookPropertyNameId
		{
            get { return OutlookPropertyPrefix + "id"; }
		}

        /*public string OutlookPropertyNameUpdated
        {
            get { return OutlookPropertyPrefix + "up"; }
        }*/

        public string OutlookPropertyNameSynced
		{
			get { return OutlookPropertyPrefix + "up"; }
		}

		private SyncOption _syncOption = SyncOption.MergeOutlookWins;
		public SyncOption SyncOption
		{
			get { return _syncOption; }
			set { _syncOption = value; }
		}

        public string SyncProfile { get; set; }
        public static string SyncContactsFolder { get; set; }
        public static string SyncNotesFolder { get; set; }
        public static string SyncAppointmentsFolder { get; set; }
        public static ushort MonthsInPast { get; set; }
        public static ushort MonthsInFuture { get; set; }

		//private ConflictResolution? _conflictResolution;
		//public ConflictResolution? CResolution
		//{
		//    get { return _conflictResolution; }
		//    set { _conflictResolution = value; }
		//}

        public List<ContactMatch> Contacts { get; private set; }

        public List<NoteMatch> Notes { get; private set; }

        public List<AppointmentMatch> Appointments { get; private set; }

        //private string _authToken;
        //public string AuthToken
        //{
        //    get
        //    {
        //        return _authToken;
        //    }
        //}

		/// <summary>
		/// If true deletes contacts if synced before, but one is missing. Otherwise contacts will bever be automatically deleted
		/// </summary>
        public bool SyncDelete { get; set; }
        public bool PromptDelete { get; set; }
        /// <summary>
        /// If true sync also notes
        /// </summary>
        public bool SyncNotes { get; set; }

        /// <summary>
        /// If true sync also contacts
        /// </summary>
        public bool SyncContacts { get; set; }

        /// <summary>
        /// If true sync also appointments (calendar)
        /// </summary>
        public bool SyncAppointments { get; set; }

        /// <summary>
        /// if true, use Outlook's FileAs for Google Title/FullName. If false, use Outlook's Fullname
        /// </summary>
        public bool UseFileAs { get; set; }

		public void LoginToGoogle(string username, string password)
		{
			Logger.Log("Connecting to Google...", EventType.Information);
            if (ContactsRequest == null && SyncContacts || DocumentsRequest==null && SyncNotes || CalendarService==null & SyncAppointments)
            {
                RequestSettings rs = new RequestSettings("GoogleContactSyncMod", username, password); 
                if (SyncContacts)
                    ContactsRequest = new ContactsRequest(rs);
                if (SyncNotes)
                {
                    DocumentsRequest = new DocumentsRequest(rs);
                    //Instantiate an Authenticator object according to your authentication, to use ResumableUploader
                    authenticator = new ClientLoginAuthenticator(Application.ProductName, DocumentsRequest.Service.ServiceIdentifier, username, password);
                }
                if (SyncAppointments)
                {
                    CalendarService = new CalendarService("GoogleContactSyncMod");
                    CalendarService.setUserCredentials(username, password);
                }
            }

			int maxUserIdLength = Syncronizer.OutlookUserPropertyMaxLength - (Syncronizer.OutlookUserPropertyTemplate.Length - 3 + 2);//-3 = to remove {0}, +2 = to add length for "id" or "up"
			string userId = username;
			if (userId.Length > maxUserIdLength)
				userId = userId.GetHashCode().ToString("X"); //if a user id would overflow UserProperty name, then use that user id hash code as id.
            //Remove characters not allowed for Outlook user property names: []_#
            userId = userId.Replace("#", "").Replace("[", "").Replace("]", "").Replace("_", "");

			OutlookPropertyPrefix = string.Format(Syncronizer.OutlookUserPropertyTemplate, userId);
		}

		public void LoginToOutlook()
		{
			Logger.Log("Connecting to Outlook...", EventType.Information);

			try
			{
                CreateOutlookInstance();
			}
			catch (Exception e)
			{

                if (!(e is COMException) && !(e is System.InvalidCastException)) 
                    throw;

				try
				{
					// If outlook was closed/terminated inbetween, we will receive an Exception
					// System.Runtime.InteropServices.COMException (0x800706BA): The RPC server is unavailable. (Exception from HRESULT: 0x800706BA)
					// so recreate outlook instance
                    //And sometimes we we receive an Exception
                    // System.InvalidCastException 0x8001010E (RPC_E_WRONG_THREAD))
					Logger.Log("Cannot connect to Outlook, creating new instance....", EventType.Information);
					/*OutlookApplication = new Outlook.Application();
					_outlookNamespace = OutlookApplication.GetNamespace("mapi");
					_outlookNamespace.Logon();*/
                    OutlookApplication = null;
                    _outlookNamespace = null;
                    CreateOutlookInstance();

                    
				}
				catch (Exception ex)
				{
					string message = "Cannot connect to Outlook.\r\nPlease restart "+Application.ProductName+" and try again. If error persists, please inform developers on OutlookForge.";
					// Error again? We need full stacktrace, display it!
					throw new Exception(message, ex);
				}
			}

		}

        private static void CreateOutlookInstance()
        {
            if (OutlookApplication == null || _outlookNamespace == null)
            {

                //Try to create new Outlook application 3 times, because mostly it fails the first time, if not yet running
                for (int i = 0; i < 3; i++)
                {
                    try
                    {
                        // First try to get the running application in case Outlook is already started
                        try
                        {
                            OutlookApplication = Marshal.GetActiveObject("Outlook.Application") as Microsoft.Office.Interop.Outlook.Application;
                        }
                        catch (COMException)
                        {
                            // That failed - try to create a new application object, launching Outlook in the background
                            OutlookApplication = new Outlook.Application();
                        }
                        break;  //Exit the for loop, if creating outllok application was successful
                    }
                    catch (COMException ex)
                    {
                        if (i == 2)
                            throw new NotSupportedException("Could not create instance of 'Microsoft Outlook'. Make sure Outlook 2003 or above version is installed and retry.", ex);
                        else //wait ten seconds and try again
                            System.Threading.Thread.Sleep(1000 * 10);
                    }
                }
                      
                if (OutlookApplication == null)
                    throw new NotSupportedException("Could not create instance of 'Microsoft Outlook'. Make sure Outlook 2003 or above version is installed and retry.");


                //Try to create new Outlook namespace 3 times, because mostly it fails the first time, if not yet running
                for (int i = 0; i < 3; i++)
                {
                    try
                    {
                        _outlookNamespace = OutlookApplication.GetNamespace("mapi");
                        break;  //Exit the for loop, if creating outllok application was successful
                    }
                    catch (COMException ex)
                    {
                        if (i == 2)
                            throw new NotSupportedException("Could not connect to 'Microsoft Outlook'. Make sure Outlook 2003 or above version is installed and running.", ex);
                        else //wait ten seconds and try again
                            System.Threading.Thread.Sleep(1000 * 10);
                    }
                }                                   

                if (_outlookNamespace == null)
                    throw new NotSupportedException("Could not connect to 'Microsoft Outlook'. Make sure Outlook 2003 or above version is installed and retry.");
                else
                    Logger.Log("Connected to Outlook: " + VersionInformation.GetOutlookVersion(OutlookApplication), EventType.Debug);
            }

            /*
            // Get default profile name from registry, as this is not always "Outlook" and would popup a dialog to choose profile
            // no matter if default profile is set or not. So try to read the default profile, fallback is still "Outlook"
            string profileName = "Outlook";
            using (RegistryKey k = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\Outlook\SocialConnector", false))
            {
                if (k != null)
                    profileName = k.GetValue("PrimaryOscProfile", "Outlook").ToString();
            }
            _outlookNamespace.Logon(profileName, null, true, false);*/

            //Just try to access the outlookNamespace to check, if it is still accessible, throws COMException, if not reachable           
            if (string.IsNullOrEmpty(SyncContactsFolder))
            {
               _outlookNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
            }
            else
            {
                _outlookNamespace.GetFolderFromID(SyncContactsFolder);
            }
        }

		public void LogoffOutlook()
		{
            try
            {
                Logger.Log("Disconnecting from Outlook...", EventType.Debug);
                if (_outlookNamespace != null)
                {
                    _outlookNamespace.Logoff();
                }
            }
            catch (Exception)
            {
                // if outlook was closed inbetween, we get an System.InvalidCastException or similar exception, that indicates that outlook cannot be acced anymore
                // so as outlook is closed anyways, we just ignore the exception here
            }
            finally
            {
                if (_outlookNamespace != null)
                    Marshal.ReleaseComObject(_outlookNamespace);
                if (OutlookApplication != null)
                    Marshal.ReleaseComObject(OutlookApplication);
                _outlookNamespace = null;
                OutlookApplication = null;
                Logger.Log("Disconnected from Outlook", EventType.Debug);
            }
		}

        public void LogoffGoogle()
        {            
            ContactsRequest = null;            
        }

        private void LoadOutlookContacts()
        {
            Logger.Log("Loading Outlook contacts...", EventType.Information);
            OutlookContacts = GetOutlookItems(Outlook.OlDefaultFolders.olFolderContacts, SyncContactsFolder);
            Logger.Log("Outlook Contacts Found: " + OutlookContacts.Count, EventType.Debug);
        }


        private void LoadOutlookNotes()
        {
            Logger.Log("Loading Outlook Notes...", EventType.Information);
            OutlookNotes = GetOutlookItems(Outlook.OlDefaultFolders.olFolderNotes, SyncNotesFolder);
            Logger.Log("Outlook Notes Found: " + OutlookNotes.Count, EventType.Debug);
        }

        private void LoadOutlookAppointments()
        {
            Logger.Log("Loading Outlook appointments...", EventType.Information);
            OutlookAppointments = GetOutlookItems(Outlook.OlDefaultFolders.olFolderCalendar, SyncAppointmentsFolder);
            Logger.Log("Outlook Appointments Found: " + OutlookAppointments.Count, EventType.Debug);
        }

        private Outlook.Items GetOutlookItems(Outlook.OlDefaultFolders outlookDefaultFolder, string syncFolder)
        {
            Outlook.MAPIFolder mapiFolder = null;
            if (string.IsNullOrEmpty(syncFolder))
            {
                mapiFolder = OutlookNameSpace.GetDefaultFolder(outlookDefaultFolder);
                if (mapiFolder == null)
                    throw new Exception("Error getting Default OutlookFolder: " + outlookDefaultFolder);
            }
            else
            {
                mapiFolder = OutlookNameSpace.GetFolderFromID(syncFolder);
                if (mapiFolder == null)
                    throw new Exception("Error getting OutlookFolder: " + syncFolder);
                
                //Outlook.MAPIFolder Folder = OutlookNameSpace.GetFolderFromID(_syncFolder);
                //if (Folder != null)
                //{
                //    for (int i = 1; i <= Folder.Folders.Count; i++)
                //    {
                //        Outlook.Folder subFolder = Folder.Folders[i] as Outlook.Folder;
                //        if ((Outlook.OlDefaultFolders.olFolderContacts == outlookDefaultFolder && Outlook.OlItemType.olContactItem == subFolder.DefaultItemType) ||
                //                 (Outlook.OlDefaultFolders.olFolderNotes == outlookDefaultFolder && Outlook.OlItemType.olNoteItem == subFolder.DefaultItemType) 
                //                )
                //        {
                //            mapiFolder = subFolder as Outlook.MAPIFolder;
                //        }
                //    }
                //}
            }

            try
            {
                Outlook.Items items = mapiFolder.Items;
                if (items == null)
                    throw new Exception("Error getting Outlook items from OutlookFolder: " + mapiFolder.Name);
                else
                    return items;
            }
            finally
            {
                if (mapiFolder != null)
                    Marshal.ReleaseComObject(mapiFolder);
                mapiFolder = null;
            }
        }
        
        
        ///// <summary>
        ///// Moves duplicates from OutlookContacts to OutlookContactDuplicates
        ///// </summary>
        //private void FilterOutlookContactDuplicates()
        //{
        //    OutlookContactDuplicates = new Collection<Outlook.ContactItem>();
            
        //    if (OutlookContacts.Count < 2)
        //        return;

        //    Outlook.ContactItem main, other;
        //    bool found = true;
        //    int index = 0;

        //    while (found)
        //    {
        //        found = false;

        //        for (int i = index; i <= OutlookContacts.Count - 1; i++)
        //        {
        //            main = OutlookContacts[i] as Outlook.ContactItem;

        //            // only look forward
        //            for (int j = i + 1; j <= OutlookContacts.Count; j++)
        //            {
        //                other = OutlookContacts[j] as Outlook.ContactItem;

        //                if (other.FileAs == main.FileAs &&
        //                    other.Email1Address == main.Email1Address)
        //                {
        //                    OutlookContactDuplicates.Add(other);
        //                    OutlookContacts.Remove(j);
        //                    found = true;
        //                    index = i;
        //                    break;
        //                }
        //            }
        //            if (found)
        //                break;
        //        }
        //    }
        //}

        private void LoadGoogleContacts()
        {
            LoadGoogleContacts(null);
            Logger.Log("Google Contacts Found: " + GoogleContacts.Count, EventType.Debug);                
        }

		private Contact LoadGoogleContacts(AtomId id)
		{
            string message = "Error Loading Google Contacts. Cannot connect to Google.\r\nPlease ensure you are connected to the internet. If you are behind a proxy, change your proxy configuration!";

            Contact ret = null;
			try
			{
                if (id == null) // Only log, if not specific Google Contacts are searched                    
				    Logger.Log("Loading Google Contacts...", EventType.Information);
                
                GoogleContacts = new Collection<Contact>();

				ContactsQuery query = new ContactsQuery(ContactsQuery.CreateContactsUri("default"));
				query.NumberToRetrieve = 256;
				query.StartIndex = 0;

                //Only load Google Contacts in My Contacts group (to avoid syncing accounts added automatically to "Weitere Kontakte"/"Further Contacts")
                Group group = GetGoogleGroupByName(myContactsGroup);
                if (group != null)
                    query.Group = group.Id;

				//query.ShowDeleted = false;
				//query.OrderBy = "lastmodified";
                				
                Feed<Contact> feed=ContactsRequest.Get<Contact>(query);

                while (feed != null)
                {
                    foreach (Contact a in feed.Entries)
                    {
                        GoogleContacts.Add(a);
                        if (id != null && id.Equals(a.ContactEntry.Id))
                            ret = a;
                    }
                    query.StartIndex += query.NumberToRetrieve;
                    feed = ContactsRequest.Get<Contact>(feed, FeedRequestType.Next);
                    
                }                
	
			}
            catch (System.Net.WebException ex)
            {
                //Logger.Log(message, EventType.Error);
                throw new GDataRequestException(message, ex);
            }
            catch (System.NullReferenceException ex)
            {
                //Logger.Log(message, EventType.Error);
                throw new GDataRequestException(message, new System.Net.WebException("Error accessing feed", ex));
            }

            return ret;
		}
		private void LoadGoogleGroups()
		{
            string message = "Error Loading Google Groups. Cannot connect to Google.\r\nPlease ensure you are connected to the internet. If you are behind a proxy, change your proxy configuration!";
            try
            {
                Logger.Log("Loading Google Groups...", EventType.Information);
                GroupsQuery query = new GroupsQuery(GroupsQuery.CreateGroupsUri("default"));
                query.NumberToRetrieve = 256;
                query.StartIndex = 0;
                //query.ShowDeleted = false;

                GoogleGroups = new Collection<Group>();

                Feed<Group> feed = ContactsRequest.Get<Group>(query);               

                while (feed != null)
                {
                    foreach (Group a in feed.Entries)
                    {
                        GoogleGroups.Add(a);
                    }
                    query.StartIndex += query.NumberToRetrieve;
                    feed = ContactsRequest.Get<Group>(feed, FeedRequestType.Next);

                }

                ////Only for debugging or reset purpose: Delete all Gougle Groups:
                //for (int i = GoogleGroups.Count; i > 0;i-- )
                //    _googleService.Delete(GoogleGroups[i-1]);
            }            
			catch (System.Net.WebException ex)
			{                               				
				//Logger.Log(message, EventType.Error);
                throw new GDataRequestException(message, ex);
			}
            catch (System.NullReferenceException ex)
            {
                //Logger.Log(message, EventType.Error);
                throw new GDataRequestException(message, new System.Net.WebException("Error accessing feed", ex));
            }

		}

        private void LoadGoogleNotes()
        {
            LoadGoogleNotes(null, null);
            Logger.Log("Google Notes Found: " + GoogleNotes.Count, EventType.Debug);                
        }

        internal Document LoadGoogleNotes(string folderUri, AtomId id)
        {
            string message = "Error Loading Google Notes. Cannot connect to Google.\r\nPlease ensure you are connected to the internet. If you are behind a proxy, change your proxy configuration!";

            Document ret = null;
            try
            {
                if (folderUri == null && id == null)
                {
                    // Only log, if not specific Google Notes are searched
                    Logger.Log("Loading Google Notes...", EventType.Information);
                    GoogleNotes = new Collection<Document>();
                }

                if (googleNotesFolder == null)
                    googleNotesFolder = GetOrCreateGoogleFolder(null, "Notes");//ToDo: Make the folder name Notes configurable in SettingsForm, for now hardcode to "Notes");

               
                if (folderUri == null)
                {
                    if (id == null)
                        folderUri = googleNotesFolder.DocumentEntry.Content.AbsoluteUri;
                    else //if newly created
                        folderUri = DocumentsRequest.BaseUri;
                }

                DocumentQuery query = new DocumentQuery(folderUri);
                query.Categories.Add(new QueryCategory(new AtomCategory("document")));
                query.NumberToRetrieve = 256;
                query.StartIndex = 0;                

                //query.ShowDeleted = false;
                //query.OrderBy = "lastmodified";
                Feed<Document> feed = DocumentsRequest.Get<Document>(query);

                while (feed != null)
                {
                    foreach (Document a in feed.Entries)
                    {
                        if (id == null)
                            GoogleNotes.Add(a);
                        else if (id.Equals(a.DocumentEntry.Id))
                        {
                            ret = a;
                            return ret;
                        }
                    }
                    query.StartIndex += query.NumberToRetrieve;
                    feed = DocumentsRequest.Get<Document>(feed, FeedRequestType.Next);

                }

            }
            catch (System.Net.WebException ex)
            {
                //Logger.Log(message, EventType.Error);
                throw new GDataRequestException(message, ex);
            }
            catch (System.NullReferenceException ex)
            {
                //Logger.Log(message, EventType.Error);
                throw new GDataRequestException(message, new System.Net.WebException("Error accessing feed", ex));
            }

            return ret;
        }

        private void LoadGoogleAppointments()
        {
            LoadGoogleAppointments(null);
            Logger.Log("Google Appointments Found: " + GoogleAppointments.Count, EventType.Debug);
        }

        private EventEntry LoadGoogleAppointments(AtomId id)
        {
            string message = "Error Loading Google appointments. Cannot connect to Google.\r\nPlease ensure you are connected to the internet. If you are behind a proxy, change your proxy configuration!";

            EventEntry ret = null;
            try
            {
                if (id == null) // Only log, if not specific Google Appointments are searched                    
                    Logger.Log("Loading Google appointments...", EventType.Information);

                GoogleAppointments = new Collection<EventEntry>();

                EventQuery query = new EventQuery("https://www.google.com/calendar/feeds/default/private/full");
                query.NumberToRetrieve = 256;
                query.StartIndex = 0;                               

                //Only Load events from last month
                if (MonthsInPast != 0)
                    query.StartTime = DateTime.Now.AddMonths(-MonthsInPast);
                if (MonthsInFuture != 0)
                    query.EndTime = DateTime.Now.AddMonths(MonthsInFuture);            

                

                EventFeed feed = CalendarService.Query(query);

                while (feed != null && feed.Entries != null && feed.Entries.Count > 0)
                {
                    foreach (EventEntry a in feed.Entries)
                    {
                        if (!a.Status.Equals(Google.GData.Calendar.EventEntry.EventStatus.CANCELED))
                        {//only return not yet cancelled events
                            GoogleAppointments.Add(a);
                            if (id != null && id.Equals(a.Id))
                                ret = a;
                        }
                        //else
                        //{
                        //    Logger.Log("Skipped Appointment because it was cancelled on Google side: " + (a.Title==null?null:a.Title.Text) + " - " + (a.Times.Count==0?null:a.Times[0].StartTime.ToString()), EventType.Information);
                        //    SkippedCount++;
                        //}
                    }
                    query.StartIndex += query.NumberToRetrieve;
                    feed = CalendarService.Query(query);

                }

            }
            catch (System.Net.WebException ex)
            {
                //Logger.Log(message, EventType.Error);
                throw new GDataRequestException(message, ex);
            }
            catch (System.NullReferenceException ex)
            {
                //Logger.Log(message, EventType.Error);
                throw new GDataRequestException(message, new System.Net.WebException("Error accessing feed", ex));
            }

            return ret;
        }

        /// <summary>
        /// Cleanup empty GoogleNotesFolders (for Outlook categories)
        /// </summary>
        internal void CleanUpGoogleCategories()
        {

            DocumentQuery query;
            Feed<Document> feed;
            List<Document> categoryFolders = GetGoogleGroups();

            if (categoryFolders != null)
            {
                foreach (Document categoryFolder in categoryFolders)
                {
                    query = new DocumentQuery(categoryFolder.DocumentEntry.Content.AbsoluteUri);
                    query.NumberToRetrieve = 256;
                    query.StartIndex = 0;

                    //query.ShowDeleted = false;
                    //query.OrderBy = "lastmodified";
                    feed = DocumentsRequest.Get<Document>(query);

                    bool isEmpty = true;
                    while (feed != null)
                    {
                        foreach (Document a in feed.Entries)
                        {
                            isEmpty = false;
                            break;
                        }
                        if (!isEmpty)
                            break;
                        query.StartIndex += query.NumberToRetrieve;
                        feed = DocumentsRequest.Get<Document>(feed, FeedRequestType.Next);
                    }

                    if (isEmpty)
                    {
                        DocumentsRequest.Delete(new Uri(Google.GData.Documents.DocumentsListQuery.documentsBaseUri + "/" +categoryFolder.ResourceId), categoryFolder.ETag);
                        Logger.Log("Deleted empty Google category folder: " + categoryFolder.Title, EventType.Information);
                    }

                }
            }
        }

        internal List<Document> GetGoogleGroups()
        {
            List<Document> categoryFolders;

            DocumentQuery query = new DocumentQuery(googleNotesFolder.DocumentEntry.Content.AbsoluteUri);
            query.Categories.Add(new QueryCategory(new AtomCategory("folder")));
            query.NumberToRetrieve = 256;
            query.StartIndex = 0;

            //query.ShowDeleted = false;
            //query.OrderBy = "lastmodified";
            Feed<Document> feed = DocumentsRequest.Get<Document>(query);
            categoryFolders = new List<Document>();

            while (feed != null)
            {
                foreach (Document a in feed.Entries)
                {
                    categoryFolders.Add(a);
                }
                query.StartIndex += query.NumberToRetrieve;
                feed = DocumentsRequest.Get<Document>(feed, FeedRequestType.Next);

            }

            return categoryFolders;
        }

        private Document GetOrCreateGoogleFolder(Document parentFolder, string title)
        {
            Document ret=null;                       

            lock (this) //Synchronize the threads
            {
                ret = GetGoogleFolder(parentFolder, title, null);
                
                if (ret == null)
                {
                    ret = new Document();
                    ret.Type = Document.DocumentType.Folder;
                    //ret.Categories.Add(new AtomCategory("http://schemas.google.com/docs/2007#folder"));
                    ret.Title = title;
                    ret = SaveGoogleNote(parentFolder, ret, DocumentsRequest);
                }
            }

            return ret;
        }

        internal Document GetGoogleFolder(Document parentFolder, string title, string uri)
        {
            Document ret = null;           

            //First get the Notes folder or create it, if not yet existing            
            DocumentQuery query = new DocumentQuery(DocumentsRequest.BaseUri);
            //Doesn't work, therefore used IsInFolder below: DocumentQuery query = new DocumentQuery((parentFolder == null) ? DocumentsRequest.BaseUri : parentFolder.DocumentEntryContent.AbsoluteUri);
            query.Categories.Add(new QueryCategory(new AtomCategory("folder")));
            if (!string.IsNullOrEmpty(title))
                query.Title = title;
            
            Feed<Document> feed = DocumentsRequest.Get<Document>(query);

            if (feed != null)
            {
                foreach (Document a in feed.Entries)
                {
                    if ((string.IsNullOrEmpty(uri) || a.Self == uri) && 
                        (parentFolder == null || IsInFolder(parentFolder, a)))
                    {
                        ret = a;
                        break;
                    }
                }
                query.StartIndex += query.NumberToRetrieve;
                feed = DocumentsRequest.Get<Document>(feed, FeedRequestType.Next);
            }
            return ret;
        }
        /// <summary>
        /// Load the contacts from Google and Outlook
        /// </summary>
        public void LoadContacts()
        {
            LoadOutlookContacts();
            LoadGoogleGroups();
            LoadGoogleContacts();
        }

        public void LoadNotes()
        {
            LoadOutlookNotes();
            LoadGoogleNotes();
        }

        public void LoadAppointments()
        {
            LoadOutlookAppointments();
            LoadGoogleAppointments();           
        }


        /// <summary>
        /// Load the contacts from Google and Outlook and match them
        /// </summary>
        public void MatchContacts()
		{
            LoadContacts();                            

			DuplicateDataException duplicateDataException;
			Contacts = ContactsMatcher.MatchContacts(this, out duplicateDataException);
			if (duplicateDataException != null)
			{
				
				if (DuplicatesFound != null)
                    DuplicatesFound("Google duplicates found", duplicateDataException.Message);
                else
                    Logger.Log(duplicateDataException.Message, EventType.Warning);
			}

            Logger.Log("Contact Matches Found: " + Contacts.Count, EventType.Debug);
		}

        /// <summary>
        /// Load the contacts from Google and Outlook and match them
        /// </summary>
        public void MatchNotes()
        {
            LoadNotes();
            Notes = NotesMatcher.MatchNotes(this);
            /*DuplicateDataException duplicateDataException;
            _matches = ContactsMatcher.MatchContacts(this, out duplicateDataException);
            if (duplicateDataException != null)
            {

                if (DuplicatesFound != null)
                    DuplicatesFound("Google duplicates found", duplicateDataException.Message);
                else
                    Logger.Log(duplicateDataException.Message, EventType.Warning);
            }*/
            Logger.Log("Note Matches Found: " + Notes.Count, EventType.Debug);
        }

        /// <summary>
        /// Load the appointments from Google and Outlook and match them
        /// </summary>
        public void MatchAppointments()
        {
            LoadAppointments();
            Appointments = AppointmentsMatcher.MatchAppointments(this);
            Logger.Log("Appointment Matches Found: " + Appointments.Count, EventType.Debug);
        }


		public void Sync()
		{
            lock (_syncRoot)
            {

                try
                {

                    if (string.IsNullOrEmpty(SyncProfile))
                    {
                        Logger.Log("Must set a sync profile. This should be different on each user/computer you sync on.", EventType.Error);
                        return;
                    }


                    SyncedCount = 0;
                    DeletedCount = 0;
                    ErrorCount = 0;
                    SkippedCount = 0;
                    SkippedCountNotMatches = 0;
                    ConflictResolution = ConflictResolution.Cancel;
                    DeleteGoogleResolution = DeleteResolution.Cancel;
                    DeleteOutlookResolution = DeleteResolution.Cancel;

                    if (SyncContacts)
                        MatchContacts();

                    if (SyncNotes)
                        MatchNotes();

                    if (SyncAppointments)
                        MatchAppointments();

#if debug
                        this.DebugContacts();
#endif

                    if (SyncContacts)
                    {
                        if (Contacts == null)
                            return;

                        TotalCount = Contacts.Count + SkippedCountNotMatches;

                        //Resolve Google duplicates from matches to be synced
                        ResolveDuplicateContacts(GoogleContactDuplicates);

                        //Remove Outlook duplicates from matches to be synced
                        if (OutlookContactDuplicates != null)
                        {
                            for (int i = OutlookContactDuplicates.Count - 1; i >= 0; i--)
                            {
                                ContactMatch match = OutlookContactDuplicates[i];
                                if (Contacts.Contains(match))
                                {
                                    //ToDo: If there has been a resolution for a duplicate above, there is still skipped increased, check how to distinguish
                                    SkippedCount++;
                                    Contacts.Remove(match);
                                }
                            }
                        }                                                            


                        Logger.Log("Syncing groups...", EventType.Information);
                        ContactsMatcher.SyncGroups(this);

                        Logger.Log("Syncing contacts...", EventType.Information);
                        ContactsMatcher.SyncContacts(this);

                        SaveContacts(Contacts);
                    }

                    if (SyncNotes)
                    {
                        if (Notes == null)
                            return;

                        TotalCount += Notes.Count;

                        Logger.Log("Syncing notes...", EventType.Information);
                        NotesMatcher.SyncNotes(this);

                        SaveNotes(Notes);

                        int timeout = 10;//seconds to wait for asynchronous upload
                        //Because notes are uploaded asynchonously, wait until all notes have been successfully uploaded
                        foreach (NoteMatch match in Notes)
                        {
                            for (int i = 0; match.AsyncUpdateCompleted.HasValue && !match.AsyncUpdateCompleted.Value && i < timeout; i++)
                            {
                                Application.DoEvents();
                                System.Threading.Thread.Sleep(1000);//DoNothing, until the Async Update is complete, but only wait maximum 10 seconds
                                Application.DoEvents();
                            }

                            if (match.AsyncUpdateCompleted.HasValue && !match.AsyncUpdateCompleted.Value)
                                Logger.Log("Asynchronous upload of note didn't finish within " + timeout + " seconds: " + match.GoogleNote.Title, EventType.Warning);
                        }
                        


                        //Delete empty Google note folders
                        CleanUpGoogleCategories();
                    }

                    if (SyncAppointments)
                    {
                        if (Appointments == null)
                            return;

                        TotalCount += Appointments.Count;

                        Logger.Log("Syncing appointments...", EventType.Information);
                        AppointmentsMatcher.SyncAppointments(this);

                        DeleteAppointments(Appointments);

                    }

                }
                finally
                {

                    if (OutlookContacts != null)
                    {
                        Marshal.ReleaseComObject(OutlookContacts);
                        OutlookContacts = null;
                    }
                    if (OutlookNotes != null)
                    {
                        Marshal.ReleaseComObject(OutlookNotes);
                        OutlookNotes = null;
                    }
                    if (OutlookAppointments != null)
                    {
                        Marshal.ReleaseComObject(OutlookAppointments);
                        OutlookAppointments = null;
                    }
                    GoogleContacts = null;
                    GoogleNotes = null;
                    GoogleAppointments = null;
                    OutlookContactDuplicates = null;
                    GoogleContactDuplicates = null;
                    GoogleGroups = null;
                    Contacts = null;
                    Notes = null;
                    Appointments = null;

                }
            }
		}

        private void ResolveDuplicateContacts(Collection<ContactMatch> googleContactDuplicates)
        {
            if (googleContactDuplicates != null)
            {
                for (int i = googleContactDuplicates.Count - 1; i >= 0; i--)
                    ResolveDuplicateContact(googleContactDuplicates[i]);
            }
        }

        private void ResolveDuplicateContact(ContactMatch match)
        { 
            if (Contacts.Contains(match))
            {
                if (_syncOption == SyncOption.MergePrompt)
                {
                    //For each OutlookDuplicate: Ask user for the GoogleContact to be synced with
                    for (int j = match.AllOutlookContactMatches.Count - 1; j >= 0 && match.AllGoogleContactMatches.Count > 0; j--)
                    {
                        OutlookContactInfo olci = match.AllOutlookContactMatches[j];
                        Outlook.ContactItem outlookContactItem = olci.GetOriginalItemFromOutlook();

                        try
                        {
                            Contact googleContact;
                            ConflictResolver r = new ConflictResolver();
                            switch (r.ResolveDuplicate(olci, match.AllGoogleContactMatches, out googleContact))
                            {
                                case ConflictResolution.Skip:
                                case ConflictResolution.SkipAlways: //Keep both entries and sync it to both sides
                                    match.AllGoogleContactMatches.Remove(googleContact);
                                    match.AllOutlookContactMatches.Remove(olci);
                                    Contacts.Add(new ContactMatch(null, googleContact));
                                    Contacts.Add(new ContactMatch(olci, null));
                                    break;
                                case ConflictResolution.OutlookWins:
                                case ConflictResolution.OutlookWinsAlways: //Keep Outlook and overwrite Google
                                    match.AllGoogleContactMatches.Remove(googleContact);
                                    match.AllOutlookContactMatches.Remove(olci);
                                    UpdateContact(outlookContactItem, googleContact);
                                    SaveContact(new ContactMatch(olci, googleContact));
                                    break;
                                case ConflictResolution.GoogleWins:
                                case ConflictResolution.GoogleWinsAlways: //Keep Google and overwrite Outlook
                                    match.AllGoogleContactMatches.Remove(googleContact);
                                    match.AllOutlookContactMatches.Remove(olci);
                                    UpdateContact(googleContact, outlookContactItem);
                                    SaveContact(new ContactMatch(olci, googleContact));
                                    break;
                                default:
                                    throw new ApplicationException("Cancelled");
                            }
                        }
                        finally
                        {
                            if (outlookContactItem != null)
                            {
                                Marshal.ReleaseComObject(outlookContactItem);
                                outlookContactItem = null;
                            }
                        }

                        //Cleanup the match, i.e. assign a proper OutlookContact and GoogleContact, because can be deleted before
                        if (match.AllOutlookContactMatches.Count == 0)
                            match.OutlookContact = null;
                        else
                            match.OutlookContact = match.AllOutlookContactMatches[0];
                    }
                }

                //Cleanup the match, i.e. assign a proper OutlookContact and GoogleContact, because can be deleted before
                if (match.AllGoogleContactMatches.Count == 0)
                    match.GoogleContact = null;
                else
                    match.GoogleContact = match.AllGoogleContactMatches[0];


                if (match.AllOutlookContactMatches.Count == 0)
                {
                    //If all OutlookContacts have been assigned by the users ==> Create one match for each remaining Google Contact to sync them to Outlook
                    Contacts.Remove(match);
                    foreach (Contact googleContact in match.AllGoogleContactMatches)
                        Contacts.Add(new ContactMatch(null, googleContact));
                }
                else if (match.AllGoogleContactMatches.Count == 0)
                {
                    //If all GoogleContacts have been assigned by the users ==> Create one match for each remaining Outlook Contact to sync them to Google
                    Contacts.Remove(match);
                    foreach (OutlookContactInfo outlookContact in match.AllOutlookContactMatches)
                        Contacts.Add(new ContactMatch(outlookContact, null));
                }
                else // if (match.AllGoogleContactMatches.Count > 1 ||
                //         match.AllOutlookContactMatches.Count > 1)
                {
                    SkippedCount++;
                    Contacts.Remove(match);
                }
                //else
                //{
                //    //If there remains a modified ContactMatch with only a single OutlookContact and GoogleContact
                //    //==>Remove all outlookContactDuplicates for this Outlook Contact to not remove it later from the Contacts to sync
                //    foreach (ContactMatch duplicate in OutlookContactDuplicates)
                //    {
                //        if (duplicate.OutlookContact.EntryID == match.OutlookContact.EntryID)
                //        {
                //            OutlookContactDuplicates.Remove(duplicate);
                //            break;
                //        }
                //    }
                //}
            }
        }

        public void DeleteAppointments(List<AppointmentMatch> appointments)
        {
            foreach (AppointmentMatch match in appointments)
            {
                try
                {
                    DeleteAppointment(match);
                }
                catch (Exception ex)
                {
                    if (ErrorEncountered != null)
                    {
                        ErrorCount++;
                        SyncedCount--;
                        string message = String.Format("Failed to synchronize appointment: {0}:\n{1}", match.OutlookAppointment != null ? match.OutlookAppointment.Subject + "(" + match.OutlookAppointment.Start + ")" : (match.GoogleAppointment.Title==null?null:match.GoogleAppointment.Title.Text) + "(" + (match.GoogleAppointment.Times.Count==0?null:match.GoogleAppointment.Times[0].StartTime.ToString()) + ")", ex.Message);
                        Exception newEx = new Exception(message, ex);
                        ErrorEncountered("Error", newEx, EventType.Error);
                    }
                    else
                        throw;
                }
            }
        }

        
        public void DeleteAppointment(AppointmentMatch match)
        {
            if (match.GoogleAppointment != null && match.OutlookAppointment != null)
            {
                // Do nothing: Outlook appointments are not saved here anymore, they have already been saved and counted, just delete items

                ////bool googleChanged, outlookChanged;
                ////SaveAppointmentGroups(match, out googleChanged, out outlookChanged);
                //if (!match.GoogleAppointment.Saved)
                //{
                //    //Google appointment was modified. save.
                //    SyncedCount++;
                //    AppointmentPropertiesUtils.SetProperty(match.GoogleAppointment, Syncronizer.OutlookAppointmentsFolder, match.OutlookAppointment.EntryID);
                //    match.GoogleAppointment.Save();
                //    Logger.Log("Updated Google appointment from Outlook: \"" + match.GoogleAppointment.Title.Text + "\".", EventType.Information);
                //}

                //if (!match.OutlookAppointment.Saved)// || outlookChanged)
                //{
                //    //outlook appointment was modified. save.
                //    SyncedCount++;
                //    AppointmentPropertiesUtils.SetProperty(match.OutlookAppointment, Syncronizer.GoogleAppointmentsFolder, match.GoogleAppointment.EntryID);
                //    match.OutlookAppointment.Save();
                //    Logger.Log("Updated Outlook appointment from Google: \"" + match.OutlookAppointment.Subject + "\".", EventType.Information);
                //}                
            }
            else if (match.GoogleAppointment == null && match.OutlookAppointment != null)
            {                
                if (match.OutlookAppointment.ItemProperties[this.OutlookPropertyNameId] != null)
                {
                    string name = match.OutlookAppointment.Subject;
                    if (_syncOption == SyncOption.OutlookToGoogleOnly)
                    {
                        SkippedCount++;
                        Logger.Log("Skipped Deletion of Outlook appointment because of SyncOption " + _syncOption + ":" + name + ".", EventType.Information);
                    }
                    else if (!SyncDelete)
                    {
                        SkippedCount++;
                        Logger.Log("Skipped Deletion of Outlook appointment because SyncDeletion is switched off: " + name + ".", EventType.Information);
                    }
                    else
                    {
                        // Google appointment was deleted, delete outlook appointment
                        Outlook.AppointmentItem item = match.OutlookAppointment;
                        //try
                        //{
                        string outlookAppointmentId = AppointmentPropertiesUtils.GetOutlookGoogleAppointmentId(this, match.OutlookAppointment);
                        try
                        {
                            //First reset OutlookGoogleContactId to restore it later from trash
                            AppointmentPropertiesUtils.ResetOutlookGoogleAppointmentId(this, item);
                            item.Save();
                        }
                        catch (Exception)
                        {
                            Logger.Log("Error resetting match for Outlook appointment: \"" + name + "\".", EventType.Warning);
                        }

                        item.Delete();

                        DeletedCount++;
                        Logger.Log("Deleted Outlook appointment: \"" + name + "\".", EventType.Information);
                        //}
                        //finally
                        //{
                        //    Marshal.ReleaseComObject(outlookContact);
                        //    outlookContact = null;
                        //}
                    }
                }
            }
            else if (match.GoogleAppointment != null && match.OutlookAppointment == null)
            {
                if (AppointmentPropertiesUtils.GetGoogleOutlookAppointmentId(SyncProfile, match.GoogleAppointment) != null)
                {
                    string name = match.GoogleAppointment.Title.Text;
                    if (_syncOption == SyncOption.OutlookToGoogleOnly)
                    {
                        SkippedCount++;
                        Logger.Log("Skipped Deletion of Google appointment because of SyncOption " + _syncOption + ":" + name + ".", EventType.Information);
                    }
                    else if (!SyncDelete)
                    {
                        SkippedCount++;
                        Logger.Log("Skipped Deletion of Google appointment because SyncDeletion is switched off: " + name + ".", EventType.Information);
                    }
                    else
                    {
                        // outlook appointment was deleted, delete Google appointment
                        EventEntry item = match.GoogleAppointment;
                        ////try
                        ////{
                        //string outlookAppointmentId = AppointmentPropertiesUtils.GetGoogleOutlookAppointmentId(SyncProfile, match.GoogleAppointment);
                        //try
                        //{
                        //    //First reset OutlookGoogleContactId to restore it later from trash
                        //    AppointmentPropertiesUtils.ResetOutlookGoogleAppointmentId(this, item);
                        //    item.Save();
                        //}
                        //catch (Exception)
                        //{
                        //    Logger.Log("Error resetting match for Google appointment: \"" + name + "\".", EventType.Warning);
                        //}

                        item.Delete();

                        DeletedCount++;
                        Logger.Log("Deleted Google appointment: \"" + name + "\".", EventType.Information);
                        //}
                        //finally
                        //{
                        //    Marshal.ReleaseComObject(outlookContact);
                        //    outlookContact = null;
                        //}
                    }
                }
            }
            else
            {
                //TODO: ignore for now: 
                throw new ArgumentNullException("To save appointments, at least a GoogleAppointment or OutlookAppointment must be present.");
                //Logger.Log("Both Google and Outlook appointment: \"" + match.OutlookAppointment.FileAs + "\" have been changed! Not implemented yet.", EventType.Warning);
            }
        }

		public void SaveContacts(List<ContactMatch> contacts)
		{
			foreach (ContactMatch match in contacts)
			{
				try
				{
					SaveContact(match);
				}
				catch (Exception ex)
				{
					if (ErrorEncountered != null)
					{
                        ErrorCount++;
                        SyncedCount--;
                        string message = String.Format("Failed to synchronize contact: {0}. \nPlease check the contact, if any Email already exists on Google contacts side or if there is too much or invalid data in the notes field. \nIf the problem persists, please try recreating the contact or report the error on OutlookForge:\n{1}", match.OutlookContact != null ? match.OutlookContact.FileAs : match.GoogleContact.Title, ex.Message);
						Exception newEx = new Exception(message, ex);
						ErrorEncountered("Error", newEx, EventType.Error);
					}
					else
						throw;
				}
			}
		}

        public void SaveNotes(List<NoteMatch> notes)
        {
            foreach (NoteMatch match in notes)
            {
                try
                {
                    SaveNote(match);
                }
                catch (Exception ex)
                {
                    if (ErrorEncountered != null)
                    {
                        ErrorCount++;
                        SyncedCount--;
                        string message = String.Format("Failed to synchronize note: {0}.", match.OutlookNote.Subject);
                        Exception newEx = new Exception(message, ex);
                        ErrorEncountered("Error", newEx, EventType.Error);
                    }
                    else
                        throw;
                }
            }
        }

       
        public void SaveContact(ContactMatch match)
        {
            if (match.GoogleContact != null && match.OutlookContact != null)
			{
				//bool googleChanged, outlookChanged;
				//SaveContactGroups(match, out googleChanged, out outlookChanged);
                if (match.GoogleContact.ContactEntry.Dirty || match.GoogleContact.ContactEntry.IsDirty())
                {
                    //google contact was modified. save.
                    SyncedCount++;					
					SaveGoogleContact(match);
					Logger.Log("Updated Google contact from Outlook: \"" + match.OutlookContact.FileAs + "\".", EventType.Information);
				}

               
			}
            else if (match.GoogleContact == null && match.OutlookContact != null)
			{
                if (match.OutlookContact.UserProperties.GoogleContactId != null)
				{
                    string name = match.OutlookContact.FileAs;
                    if (_syncOption == SyncOption.OutlookToGoogleOnly)
                    {
                        SkippedCount++;
                        Logger.Log("Skipped Deletion of Outlook contact because of SyncOption " + _syncOption + ":" + name + ".", EventType.Information);
                    }
                    else if (!SyncDelete)
                    {
                        SkippedCount++;
                        Logger.Log("Skipped Deletion of Outlook contact because SyncDeletion is switched off: " + name + ".", EventType.Information);
                    }
                    else
                    {
                        // peer google contact was deleted, delete outlook contact
                        Outlook.ContactItem item = match.OutlookContact.GetOriginalItemFromOutlook();
                        try
                        {
                            try
                            {
                                //First reset OutlookGoogleContactId to restore it later from trash
                                ContactPropertiesUtils.ResetOutlookGoogleContactId(this, item);
                                item.Save();                                
                            }
                            catch (Exception)
                            {
                                Logger.Log("Error resetting match for Outlook contact: \"" + name + "\".", EventType.Warning);
                            }

                            item.Delete();
                            DeletedCount++;
                            Logger.Log("Deleted Outlook contact: \"" + name + "\".", EventType.Information);
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(item);
                            item = null;
                        }
                    }
				}
			}
            else if (match.GoogleContact != null && match.OutlookContact == null)
			{
				if (ContactPropertiesUtils.GetGoogleOutlookContactId(SyncProfile, match.GoogleContact) != null)
				{                    

                    if (_syncOption == SyncOption.GoogleToOutlookOnly)
                    {
                        SkippedCount++;
                        Logger.Log("Skipped Deletion of Google contact because of SyncOption " + _syncOption + ":" + ContactMatch.GetName(match.GoogleContact) + ".", EventType.Information);
                    }
                    else if (!SyncDelete)
                    {
                        SkippedCount++;
                        Logger.Log("Skipped Deletion of Google contact because SyncDeletion is switched off :" + ContactMatch.GetName(match.GoogleContact) + ".", EventType.Information);
                    }
                    else
                    {
                        //commented oud, because it causes precondition failed error, if the ResetMatch is short before the Delete
                        //// peer outlook contact was deleted, delete google contact
                        //try
                        //{
                        //    //First reset GoogleOutlookContactId to restore it later from trash
                        //    match.GoogleContact = ResetMatch(match.GoogleContact);
                        //}
                        //catch (Exception)
                        //{
                        //    Logger.Log("Error resetting match for Google contact: \"" + ContactMatch.GetName(match.GoogleContact) + "\".", EventType.Warning);
                        //}

                        ContactsRequest.Delete(match.GoogleContact);
                        DeletedCount++;
                        Logger.Log("Deleted Google contact: \"" + ContactMatch.GetName(match.GoogleContact) + "\".", EventType.Information);
                    }
				}
			}
			else
			{
				//TODO: ignore for now: 
                throw new ArgumentNullException("To save contacts, at least a GoogleContacat or OutlookContact must be present.");
				//Logger.Log("Both Google and Outlook contact: \"" + match.OutlookContact.FileAs + "\" have been changed! Not implemented yet.", EventType.Warning);
			}
		}

        public void SaveNote(NoteMatch match)
        {
            if (match.GoogleNote != null && match.OutlookNote != null)
            {
                //bool googleChanged, outlookChanged;
                //SaveNoteGroups(match, out googleChanged, out outlookChanged);
                if (match.GoogleNote.DocumentEntry.Dirty || match.GoogleNote.DocumentEntry.IsDirty())
                {
                    //google note was modified. save.
                    SyncedCount++;
                    SaveGoogleNote(match);
                    //Don't log here, because the DocumentsRequest uses async upload, log when async upload was successful
                    //Logger.Log("Updated Google note from Outlook: \"" + match.OutlookNote.Subject + "\".", EventType.Information);
                } 
                else if (!match.OutlookNote.Saved)// || outlookChanged) //If google note is saved above, Saving the OutlookNote not necessary anymore, because this will be done when updating NoteMatchId during saving the Google Note above
                {
                    //outlook note was modified. save.
                    SyncedCount++;
                    NotePropertiesUtils.SetOutlookGoogleNoteId(this, match.OutlookNote, match.GoogleNote);
                    match.OutlookNote.Save();
                    Logger.Log("Updated Outlook note from Google: \"" + match.OutlookNote.Subject + "\".", EventType.Information);
                }                

                // save photos
                //SaveNotePhotos(match);
            }
            else if (match.GoogleNote == null && match.OutlookNote != null)
            {
                if (match.OutlookNote.ItemProperties[this.OutlookPropertyNameId] != null)
                {
                    string name = match.OutlookNote.Subject;
                    if (_syncOption == SyncOption.OutlookToGoogleOnly)
                    {
                        SkippedCount++;
                        Logger.Log("Skipped Deletion of Outlook note because of SyncOption " + _syncOption + ":" + name + ".", EventType.Information);
                    }
                    else if (!SyncDelete)
                    {
                        SkippedCount++;
                        Logger.Log("Skipped Deletion of Outlook note because SyncDeletion is switched off: " + name + ".", EventType.Information);
                    }
                    else
                    {
                        // peer google note was deleted, delete outlook note
                        Outlook.NoteItem item = match.OutlookNote;
                        //try
                        //{
                        string outlookNoteId = NotePropertiesUtils.GetOutlookGoogleNoteId(this, match.OutlookNote);
                            try
                            {
                                //First reset OutlookGoogleContactId to restore it later from trash
                                NotePropertiesUtils.ResetOutlookGoogleNoteId(this, item);
                                item.Save();
                            }
                            catch (Exception)
                            {
                                Logger.Log("Error resetting match for Outlook note: \"" + name + "\".", EventType.Warning);
                            }

                            item.Delete();
                            try
                            { //Delete also the according temporary NoteFile
                                File.Delete(NotePropertiesUtils.GetFileName(outlookNoteId, SyncProfile));
                            }
                            catch (Exception)
                            { }
                            DeletedCount++;
                            Logger.Log("Deleted Outlook note: \"" + name + "\".", EventType.Information);
                        //}
                        //finally
                        //{
                        //    Marshal.ReleaseComObject(item);
                        //    item = null;
                        //}
                    }
                }
            }
            else if (match.GoogleNote != null && match.OutlookNote == null)
            {
                if (NotePropertiesUtils.NoteFileExists(match.GoogleNote.Id, SyncProfile))
                {
                    string name = match.GoogleNote.Title;                    

                    if (_syncOption == SyncOption.GoogleToOutlookOnly)
                    {
                        SkippedCount++;
                        Logger.Log("Skipped Deletion of Google note because of SyncOption " + _syncOption + ":" + name + ".", EventType.Information);
                    }
                    else if (!SyncDelete)
                    {
                        SkippedCount++;
                        Logger.Log("Skipped Deletion of Google note because SyncDeletion is switched off :" + name + ".", EventType.Information);
                    }
                    else
                    {
                        // peer outlook note was deleted, delete google note
                        DocumentsRequest.Delete(new Uri(Google.GData.Documents.DocumentsListQuery.documentsBaseUri + "/" + match.GoogleNote.ResourceId), match.GoogleNote.ETag);
                        //DocumentsRequest.Service.Delete(match.GoogleNote.DocumentEntry); //ToDo: Currently, the Delete only removes the Notes label from the document but keeps the document in the root folder, therefore I use the URI Delete above for now: "https://docs.google.com/feeds/default/private/full"
                        //DocumentsRequest.Delete(match.GoogleNote);

                        ////ToDo: Currently, the Delete only removes the Notes label from the document but keeps the document in the root folder, therefore the following workaround
                        //Document deletedNote = LoadGoogleNotes(match.GoogleNote.DocumentEntry.Id);
                        //if (deletedNote != null)
                        //    DocumentsRequest.Delete(deletedNote);
                        
                        try
                         {//Delete also the according temporary NoteFile
                             File.Delete(NotePropertiesUtils.GetFileName(match.GoogleNote.Id, SyncProfile));
                         }
                         catch (Exception)
                         {}

                        DeletedCount++;
                        Logger.Log("Deleted Google note: \"" + name + "\".", EventType.Information);
                    }
                }
            }
            else
            {
                //TODO: ignore for now: 
                throw new ArgumentNullException("To save notes, at least a GoogleContacat or OutlookNote must be present.");
                //Logger.Log("Both Google and Outlook note: \"" + match.OutlookNote.FileAs + "\" have been changed! Not implemented yet.", EventType.Warning);
            }
        }

        /// <summary>
        /// Updates Outlook appointment from master to slave (including groups/categories)
        /// </summary>
        public void UpdateAppointment(Outlook.AppointmentItem master, ref EventEntry slave)
        {
            AppointmentSync.UpdateAppointment(master, slave);

            AppointmentPropertiesUtils.SetGoogleOutlookAppointmentId(SyncProfile, slave, master);
            slave = SaveGoogleAppointment(slave);

            AppointmentPropertiesUtils.SetOutlookGoogleAppointmentId(this, master,slave);
            master.Save();                    

            SyncedCount++;
            Logger.Log("Updated appointment from Outlook to Google: \"" + master.Subject + " - " + master.Start + "\".", EventType.Information);
        }

        /// <summary>
        /// Updates Outlook appointment from master to slave (including groups/categories)
        /// </summary>
        public void UpdateAppointment(ref EventEntry master, Outlook.AppointmentItem slave)
        {
            AppointmentSync.UpdateAppointment(master, slave);

            AppointmentPropertiesUtils.SetOutlookGoogleAppointmentId(this, slave, master);
            slave.Save();

            AppointmentPropertiesUtils.SetGoogleOutlookAppointmentId(SyncProfile, master, slave);
            master = SaveGoogleAppointment(master);

            SyncedCount++;
            Logger.Log("Updated appointment from Google to Outlook: \"" + (master.Title == null ? null : master.Title.Text) + " - " + (master.Times.Count==0?null:master.Times[0].StartTime.ToString()) + "\".", EventType.Information);
        }

        private void SaveOutlookContact(ref Contact googleContact, Outlook.ContactItem outlookContact)
        {
            ContactPropertiesUtils.SetOutlookGoogleContactId(this, outlookContact, googleContact);
            outlookContact.Save();
            //Because Outlook automatically sets the EmailDisplayName to default value when the email is changed, update the emails again, to also sync the DisplayName
            ContactSync.SetEmails(googleContact, outlookContact);
            ContactPropertiesUtils.SetGoogleOutlookContactId(SyncProfile, googleContact, outlookContact);

            Contact updatedEntry = SaveGoogleContact(googleContact);
            //try
            //{
            //    updatedEntry = _googleService.Update(match.GoogleContact);
            //}
            //catch (GDataRequestException tmpEx)
            //{
            //    // check if it's the known HTCData problem, or if there is any invalid XML element or any unescaped XML sequence
            //    //if (tmpEx.ResponseString.Contains("HTCData") || tmpEx.ResponseString.Contains("&#39") || match.GoogleContact.Content.Contains("<"))
            //    //{
            //    //    bool wasDirty = match.GoogleContact.ContactEntry.Dirty;
            //    //    // XML escape the content
            //    //    match.GoogleContact.Content = EscapeXml(match.GoogleContact.Content);
            //    //    // set dirty to back, cause we don't want the changed content go back to Google without reason
            //    //    match.GoogleContact.ContactEntry.Content.Dirty = wasDirty;
            //    //    updatedEntry = _googleService.Update(match.GoogleContact);
                    
            //    //}
            //    //else 
            //    if (!String.IsNullOrEmpty(tmpEx.ResponseString))
            //        throw new ApplicationException(tmpEx.ResponseString, tmpEx);
            //    else
            //        throw;
            //}            
            googleContact = updatedEntry;

            ContactPropertiesUtils.SetOutlookGoogleContactId(this, outlookContact, googleContact);
            outlookContact.Save();
            SaveOutlookPhoto(googleContact, outlookContact);
        }
		private static string EscapeXml(string xml)
		{
			string encodedXml = System.Security.SecurityElement.Escape(xml);
			return encodedXml;
		}
		public void SaveGoogleContact(ContactMatch match)
		{
            Outlook.ContactItem outlookContactItem = match.OutlookContact.GetOriginalItemFromOutlook();
            try
            {
                ContactPropertiesUtils.SetGoogleOutlookContactId(SyncProfile, match.GoogleContact, outlookContactItem);
                match.GoogleContact = SaveGoogleContact(match.GoogleContact);
                ContactPropertiesUtils.SetOutlookGoogleContactId(this, outlookContactItem, match.GoogleContact);
                outlookContactItem.Save();

                //Now save the Photo
                SaveGooglePhoto(match, outlookContactItem);                       

            }
            finally
            {
                Marshal.ReleaseComObject(outlookContactItem);
                outlookContactItem = null;
            }
		}

        public void SaveGoogleNote(NoteMatch match)
        {
            Outlook.NoteItem outlookNoteItem = match.OutlookNote;
            //try
            //{  

                //ToDo: Somewhow, the content is not uploaded to Google, only an empty document                
                //match.GoogleNote = SaveGoogleNote(match.GoogleNote);

            //New approach how to update an existing document: https://developers.google.com/google-apps/documents-list/#updatingchanging_documents_and_files
            // Instantiate the ResumableUploader component.      
            ResumableUploader uploader = new ResumableUploader();
            // Set the handlers for the completion and progress events                  
            //uploader.AsyncOperationProgress += new AsyncOperationProgressEventHandler(OnProgress);
                
                //ToDo: Therefoe I use DocumentService.UploadDocument instead and move it to the NotesFolder
                string oldOutlookGoogleNoteId = NotePropertiesUtils.GetOutlookGoogleNoteId(this, outlookNoteItem);
                if (match.GoogleNote.DocumentEntry.Id.Uri != null)
                {
                    //DocumentsRequest.Delete(new Uri(Google.GData.Documents.DocumentsListQuery.documentsBaseUri + "/" + match.GoogleNote.ResourceId), match.GoogleNote.ETag);
                    ////DocumentsRequest.Delete(match.GoogleNote); //ToDo: Currently, the Delete only removes the Notes label from the document but keeps the document in the root folder
                    //NotePropertiesUtils.ResetOutlookGoogleNoteId(this, outlookNoteItem);                                        

                    ////ToDo: Currently, the Delete only removes the Notes label from the document but keeps the document in the root folder
                    //Document deletedNote = LoadGoogleNotes(match.GoogleNote.DocumentEntry.Id);
                    //if (deletedNote != null)
                    //    DocumentsRequest.Delete(deletedNote);

                    // Start the update process.  
                    uploader.AsyncOperationCompleted += new AsyncOperationCompletedEventHandler(OnGoogleNoteUpdated);    
                    uploader.UpdateAsync(authenticator, match.GoogleNote.DocumentEntry, match);
                    //uploader.Update(_authenticator, match.GoogleNote.DocumentEntry);

                }
                else
                {
                    uploader.AsyncOperationCompleted += new AsyncOperationCompletedEventHandler(OnGoogleNoteCreated);
                    CreateGoogleNote(match.GoogleNote, match, DocumentsRequest, uploader, authenticator);                
                }

                match.AsyncUpdateCompleted = false;

                //Google.GData.Documents.DocumentEntry entry = DocumentsRequest.Service.UploadDocument(NotePropertiesUtils.GetFileName(outlookNoteItem.EntryID, SyncProfile), match.GoogleNote.Title.Replace(":", String.Empty));                               
                //Document newNote = LoadGoogleNotes(entry.Id);
                //match.GoogleNote = DocumentsRequest.MoveDocumentTo(GoogleNotesFolder, newNote);

                //First delete old temporary file, because it was saved with old GoogleNoteID, because every sync to Google becomes a new ID, because updateMedia doesn't work
                //File.Delete(NotePropertiesUtils.GetFileName(oldOutlookGoogleNoteId, SyncProfile));
                //UpdateNoteMatchId(match);
            //}
            //finally
            //{
            //    Marshal.ReleaseComObject(outlookNoteItem);
            //    outlookNoteItem = null;
            //}
        }

        public static void CreateGoogleNote(/*Document parentFolder, */Document googleNote, object UserData, DocumentsRequest documentsRequest, ResumableUploader uploader, ClientLoginAuthenticator authenticator)
        {
            // Define the resumable upload link      
            Uri createUploadUrl = new Uri("https://docs.google.com/feeds/upload/create-session/default/private/full");
            //Uri createUploadUrl = new Uri(GoogleNotesFolder.AtomEntry.EditUri.ToString()); 
            AtomLink link = new AtomLink(createUploadUrl.AbsoluteUri);
            link.Rel = ResumableUploader.CreateMediaRelation;
            googleNote.DocumentEntry.Links.Add(link);
            //if (parentFolder != null)
            //    googleNote.DocumentEntry.ParentFolders.Add(new AtomLink(parentFolder.DocumentEntry.SelfUri.ToString()));
            // Set the service to be used to parse the returned entry 
            googleNote.DocumentEntry.Service = documentsRequest.Service;
            // Start the upload process   
            //uploader.InsertAsync(_authenticator, match.GoogleNote.DocumentEntry, new object());
            uploader.InsertAsync(authenticator, googleNote.DocumentEntry, UserData);
        }

        private void UpdateNoteMatchId(NoteMatch match)
        {
            NotePropertiesUtils.SetOutlookGoogleNoteId(this, match.OutlookNote, match.GoogleNote);
            match.OutlookNote.Save();

            //As GoogleDocuments don't have UserProperties, we have to use the file to check, if Note was already synced or not
            File.Delete(NotePropertiesUtils.GetFileName(match.GoogleNote.Id, SyncProfile));
            File.Move(NotePropertiesUtils.GetFileName(match.OutlookNote.EntryID, SyncProfile), NotePropertiesUtils.GetFileName(match.GoogleNote.Id, SyncProfile));
        }

        private void OnGoogleNoteCreated(object sender, AsyncOperationCompletedEventArgs e)
        {           
            MoveGoogleNote(e.Entry as DocumentEntry, e.UserState as NoteMatch, true, e.Error, e.Cancelled);
        }

        private void OnGoogleNoteUpdated(object sender, AsyncOperationCompletedEventArgs e)
        {
            MoveGoogleNote(e.Entry as DocumentEntry, e.UserState as NoteMatch, false, e.Error, e.Cancelled);
        }
        private void MoveGoogleNote(DocumentEntry entry, NoteMatch match, bool create, Exception ex, bool cancelled)
        {
            if (ex != null)
            {
                ErrorHandler.Handle(new Exception("Google Note couldn't be " + (create?"created":"updated") + " :" + entry == null ? null : entry.Title.Text, ex));
                return;
            }

            if (cancelled || entry == null)
            {
                ErrorHandler.Handle(new Exception("Google Note " + (create ? "creation" : "update") + " was cancelled: " + entry == null ? null : entry.Title.Text));
                return;
            }

           //Get updated Google Note
            Document newNote = LoadGoogleNotes(null, entry.Id);
            match.GoogleNote = newNote;

            //Doesn't work because My Drive is not listed as parent folder: Remove all parent folders except for the Notes subfolder
            //if (create)
            //{
            //    foreach (string parentFolder in newNote.ParentFolders)
            //        if (parentFolder != googleNotesFolder.Self)
            //            DocumentsRequest.Delete(new Uri(googleNotesFolder.DocumentEntry.Content.AbsoluteUri + "/" + newNote.ResourceId),newNote.ETag);
            //}

            //first delete the note from all categories, the still valid categories are assigned again later           
            foreach (string parentFolder in newNote.ParentFolders)
                if (parentFolder != googleNotesFolder.Self) //Except for Notes root folder
                {
                    Document deletedNote = LoadGoogleNotes(parentFolder + "/contents", newNote.DocumentEntry.Id);
                    //DocumentsRequest.Delete(new Uri(parentFolder + "/contents/" + newNote.ResourceId), newNote.ETag);
                    DocumentsRequest.Delete(deletedNote); //Just delete it from this category
                }

            //Move now to Notes subfolder (if not already there)
            if (!IsInFolder(googleNotesFolder, newNote))
                newNote = DocumentsRequest.MoveDocumentTo(googleNotesFolder, newNote);

            //Move now to all categories subfolder (if not already there)
            foreach (string category in Utilities.GetOutlookGroups(match.OutlookNote.Categories))
            {                
                Document categoryFolder = GetOrCreateGoogleFolder(googleNotesFolder, category);    
            
                if (!IsInFolder(categoryFolder, newNote))
                    newNote = DocumentsRequest.MoveDocumentTo(categoryFolder, newNote);
            }           

            //Then update the match IDs
            UpdateNoteMatchId(match);

            Logger.Log((create?"Created":"Updated") + " Google note from Outlook: \"" + match.OutlookNote.Subject + "\".", EventType.Information);
            //Then release this match as completed (to not log the summary already before each single note result has been synced
            match.AsyncUpdateCompleted = true;
        }

        /// <summary>
        /// returns true, if the passed document is already in passed parentFolder
        /// </summary>
        /// <param name="parentFolder">the parent folder</param>
        /// <param name="childDocument">the document to be checked</param>
        /// <returns></returns>
        private bool IsInFolder(Document parentFolder, Document childDocument)
        {
            foreach (string parent in childDocument.ParentFolders)
            {
                if (parent == parentFolder.Self)
                {
                    return true;
                }
            }

            return false;
        }

        

		private string GetXml(Contact contact)
		{
			MemoryStream ms = new MemoryStream();
			contact.ContactEntry.SaveToXml(ms);
			StreamReader sr = new StreamReader(ms);
			ms.Seek(0, SeekOrigin.Begin);
			string xml = sr.ReadToEnd();
			return xml;
		}

        private static string GetXml(Document note)
        {
            MemoryStream ms = new MemoryStream();
            note.DocumentEntry.SaveToXml(ms);
            StreamReader sr = new StreamReader(ms);
            ms.Seek(0, SeekOrigin.Begin);
            string xml = sr.ReadToEnd();
            return xml;
        }

        private static string GetXml(EventEntry appointment)
        {
            MemoryStream ms = new MemoryStream();
            appointment.SaveToXml(ms);
            StreamReader sr = new StreamReader(ms);
            ms.Seek(0, SeekOrigin.Begin);
            string xml = sr.ReadToEnd();
            return xml;
        }

        /// <summary>
        /// Only save the google contact without photo update
        /// </summary>
        /// <param name="googleContact"></param>
		internal Contact SaveGoogleContact(Contact googleContact)
		{
			//check if this contact was not yet inserted on google.
			if (googleContact.ContactEntry.Id.Uri == null)
			{
				//insert contact.
				Uri feedUri = new Uri(ContactsQuery.CreateContactsUri("default"));

				try
				{
					Contact createdEntry = ContactsRequest.Insert(feedUri, googleContact);
                    return createdEntry;
				}
                catch (Exception ex)
                {
                    string responseString = "";
                    if (ex is GDataRequestException)
                        responseString = EscapeXml(((GDataRequestException)ex).ResponseString);
                    string xml = GetXml(googleContact);
                    string newEx = String.Format("Error saving NEW Google contact: {0}. \n{1}\n{2}", responseString, ex.Message, xml);
                    throw new ApplicationException(newEx, ex);
                }
			}
			else
			{
				try
				{
					//contact already present in google. just update
					
                    // User can create an empty label custom field on the web, but when I retrieve, and update, it throws this:
                    // Data Request Error Response: [Line 12, Column 44, element gContact:userDefinedField] Missing attribute: &#39;key&#39;
                    // Even though I didn't touch it.  So, I will search for empty keys, and give them a simple name.  Better than deleting...
                    int fieldCount = 0;
                    foreach (UserDefinedField userDefinedField in googleContact.ContactEntry.UserDefinedFields)
                    {
                        fieldCount++;
                        if (String.IsNullOrEmpty(userDefinedField.Key))
                        {
                            userDefinedField.Key = "UserField" + fieldCount.ToString();
                        }
                    }

                    //TODO: this will fail if original contact had an empty name or rpimary email address.
                    Contact updated = ContactsRequest.Update(googleContact);
                    return updated;
				}
                catch (Exception ex)
                {
                    string responseString = "";
                    if (ex is GDataRequestException)
                        responseString = EscapeXml(((GDataRequestException)ex).ResponseString);
                    string xml = GetXml(googleContact);
                    string newEx = String.Format("Error saving EXISTING Google contact: {0}. \n{1}\n{2}", responseString, ex.Message, xml);
                    throw new ApplicationException(newEx, ex);
                }
			}
		}

        /// <summary>
        /// Save the google Appointment
        /// </summary>
        /// <param name="googleAppointment"></param>
        internal EventEntry SaveGoogleAppointment(EventEntry googleAppointment)
        {
            //check if this contact was not yet inserted on google.
            if (googleAppointment.Id.Uri == null)
            {
                //insert contact.
                Uri feedUri = new Uri("https://www.google.com/calendar/feeds/default/private/full");

                try
                {
                    EventEntry createdEntry = CalendarService.Insert(feedUri, googleAppointment);
                    return createdEntry;
                }
                catch (Exception ex)
                {
                    string responseString = "";
                    if (ex is GDataRequestException)
                        responseString = EscapeXml(((GDataRequestException)ex).ResponseString);
                    string xml = GetXml(googleAppointment);
                    string newEx = String.Format("Error saving NEW Google appointment: {0}. \n{1}\n{2}", responseString, ex.Message, xml);
                    throw new ApplicationException(newEx, ex);
                }
            }
            else
            {
                try
                {
                    //contact already present in google. just update
                   
                    EventEntry updated = CalendarService.Update(googleAppointment);
                    return updated;
                }
                catch (Exception ex)
                {
                    string responseString = "";
                    if (ex is GDataRequestException)
                        responseString = EscapeXml(((GDataRequestException)ex).ResponseString);
                    string xml = GetXml(googleAppointment);
                    string newEx = String.Format("Error saving EXISTING Google appointment: {0}. \n{1}\n{2}", responseString, ex.Message, xml);
                    throw new ApplicationException(newEx, ex);
                }
            }
        }


        /// <summary>
        /// save the google note
        /// </summary>
        /// <param name="googleNote"></param>
        public static Document SaveGoogleNote(Document parentFolder, Document googleNote, DocumentsRequest documentsRequest)
        {
            //check if this contact was not yet inserted on google.
            if (googleNote.DocumentEntry.Id.Uri == null)
            {
                //insert contact.
                Uri feedUri = null;

                if (parentFolder != null)
                {
                    try
                    {//In case of Notes folder creation, the GoogleNotesFolder.DocumentEntry.Content.AbsoluteUri throws a NullReferenceException
                        feedUri = new Uri(parentFolder.DocumentEntry.Content.AbsoluteUri);
                    }
                    catch (Exception)
                    { }
                }

                if (feedUri == null)                
                    feedUri = new Uri(documentsRequest.BaseUri);               

                try
                {
                    Document createdEntry = documentsRequest.Insert(feedUri, googleNote);
                    //ToDo: Workaround also doesn't help: Utilities.SaveGoogleNoteContent(this, createdEntry, googleNote);    
                    Logger.Log("Created new Google folder: " + createdEntry.Title, EventType.Information);
                    return createdEntry;
                }
                catch (Exception ex)
                {
                    string responseString = "";
                    if (ex is GDataRequestException)
                        responseString = EscapeXml(((GDataRequestException)ex).ResponseString);
                    string xml = GetXml(googleNote);
                    string newEx = String.Format("Error saving NEW Google note: {0}. \n{1}\n{2}", responseString, ex.Message, xml);
                    throw new ApplicationException(newEx, ex);
                }
            }
            else
            {
                try
                {
                    //note already present in google. just update
                    Document updated = documentsRequest.Update(googleNote);

                    //ToDo: Workaround also doesn't help: Utilities.SaveGoogleNoteContent(this, updated, googleNote);                   

                    return updated;
                }
                catch (Exception ex)
                {
                    string responseString = "";
                    if (ex is GDataRequestException)
                        responseString = EscapeXml(((GDataRequestException)ex).ResponseString);
                    string xml = GetXml(googleNote);
                    string newEx = String.Format("Error saving EXISTING Google note: {0}. \n{1}\n{2}", responseString, ex.Message, xml);
                    throw new ApplicationException(newEx, ex);
                }
            }
        }         

        //public void SaveContactPhotos(ContactMatch match)
        //{
        //    bool hasGooglePhoto = Utilities.HasPhoto(match.GoogleContact);
        //    bool hasOutlookPhoto = Utilities.HasPhoto(match.OutlookContact);

        //    if (!hasGooglePhoto && !hasOutlookPhoto)
        //        return;
        //    else if (hasGooglePhoto && _syncOption != SyncOption.OutlookToGoogleOnly)
        //    {
        //        // add google photo to outlook
        //        Image googlePhoto = Utilities.GetGooglePhoto(this, match.GoogleContact);
        //        Utilities.SetOutlookPhoto(match.OutlookContact, googlePhoto);
        //        match.OutlookContact.Save();

        //        googlePhoto.Dispose();
        //    }
        //    else if (hasOutlookPhoto && _syncOption != SyncOption.GoogleToOutlookOnly)
        //    {
        //        // add outlook photo to google
        //        Image outlookPhoto = Utilities.GetOutlookPhoto(match.OutlookContact);
        //        if (outlookPhoto != null)
        //        {
        //            outlookPhoto = Utilities.CropImageGoogleFormat(outlookPhoto);
        //            bool saved = Utilities.SaveGooglePhoto(this, match.GoogleContact, outlookPhoto);
        //            if (!saved)
        //                throw new Exception("Could not save");

        //            outlookPhoto.Dispose();
        //        }
        //    }
        //    else
        //    {
        //        // TODO: if both contacts have photos and one is updated, the
        //        // other will not be updated.
        //    }

        //    //Utilities.DeleteTempPhoto();
        //}

        public void SaveGooglePhoto(ContactMatch match, Outlook.ContactItem outlookContactitem)
        {
            bool hasGooglePhoto = Utilities.HasPhoto(match.GoogleContact);
            bool hasOutlookPhoto = Utilities.HasPhoto(outlookContactitem);

            if (hasOutlookPhoto)
            {
                // add outlook photo to google
                Image outlookPhoto = Utilities.GetOutlookPhoto(outlookContactitem);

                if (outlookPhoto != null)
                {
                    //Try up to 5 times to overcome Google issue
                    for (int retry = 0; retry < 5; retry++)
                    {
                        try
                        {
                           
                            using (MemoryStream stream = new MemoryStream(Utilities.BitmapToBytes(new Bitmap(outlookPhoto))))
                            {
                                // Save image to stream.
                                //outlookPhoto.Save(stream, System.Drawing.Imaging.ImageFormat.Bmp);

                                //Don'T crop, because maybe someone wants to keep his photo like it is on Outlook
                                //outlookPhoto = Utilities.CropImageGoogleFormat(outlookPhoto);                        
                                ContactsRequest.SetPhoto(match.GoogleContact, stream);                        

                                //Just save the Outlook Contact to have the same lastUpdate date as Google
                                ContactPropertiesUtils.SetOutlookGoogleContactId(this, outlookContactitem, match.GoogleContact);
                                outlookContactitem.Save();
                                outlookPhoto.Dispose();
                        
                            }

                            break; //Exit because photo save succeeded
                        }
                        catch (GDataRequestException ex)
                        { //If Google found a picture for a new Google account, it sets it automatically and throws an error, if updating it with the Outlook photo. 
                            //Therefore save it again and try again to save the photo
                            if (retry == 4)
                                ErrorHandler.Handle(new Exception("Photo of contact " + match.GoogleContact.Title + "couldn't be saved after 5 tries, maybe Google found its own photo and doesn't allow updating it", ex));
                            else
                            {
                                System.Threading.Thread.Sleep(1000);
                                //LoadGoogleContact again to get latest ETag
                                //match.GoogleContact = LoadGoogleContacts(match.GoogleContact.AtomEntry.Id);
                                match.GoogleContact = SaveGoogleContact(match.GoogleContact);
                            }
                        }
                    }
                }
            }
            else if (hasGooglePhoto)
            {
                //Delete Photo on Google side, if no Outlook photo exists
                ContactsRequest.Delete(match.GoogleContact.PhotoUri, match.GoogleContact.PhotoEtag);
            }

            Utilities.DeleteTempPhoto();
        }

        //public void SaveOutlookPhoto(ContactMatch match)
        //{
        //    bool hasGooglePhoto = Utilities.HasPhoto(match.GoogleContact);
        //    bool hasOutlookPhoto = Utilities.HasPhoto(match.OutlookContact);

        //    if (hasGooglePhoto)
        //    {
        //        Image image = new Image(match.GoogleContact.PhotoUri);
        //        Utilities.SetOutlookPhoto(match.OutlookContact, image);
        //        ContactPropertiesUtils.SetOutlookGoogleContactId(this, match.OutlookContact, match.GoogleContact);
        //        match.OutlookContact.Save();

        //        //googlePhoto.Dispose();
        //    }
        //    else if (hasOutlookPhoto)
        //    {
        //        match.OutlookContact.RemovePicture();
        //        ContactPropertiesUtils.SetOutlookGoogleContactId(this, match.OutlookContact, match.GoogleContact);
        //        match.OutlookContact.Save();
        //    }
        //}

        //public void SaveGooglePhoto(ContactMatch match)
        //{
        //    bool hasGooglePhoto = Utilities.HasPhoto(match.GoogleContact);
        //    bool hasOutlookPhoto = Utilities.HasPhoto(match.OutlookContact);

        //    if (hasOutlookPhoto)
        //    {
        //        // add outlook photo to google
        //        Image outlookPhoto = Utilities.GetOutlookPhoto(match.OutlookContact);
        //        if (outlookPhoto != null)
        //        {
        //            //outlookPhoto = Utilities.CropImageGoogleFormat(outlookPhoto);
        //            bool saved = Utilities.SaveGooglePhoto(this, match.GoogleContact, outlookPhoto);
        //            if (!saved)
        //                throw new Exception("Could not save");

        //            //Just save the Outlook Contact to have the same lastUpdate date as Google
        //            ContactPropertiesUtils.SetOutlookGoogleContactId(this, match.OutlookContact, match.GoogleContact);
        //            match.OutlookContact.Save();
        //            outlookPhoto.Dispose();
        //        }
        //    }
        //    else if (hasGooglePhoto)
        //    {
        //        //ToDo: Delete Photo on Google side, if no Outlook photo exists
        //        //match.GoogleContact.PhotoUri = null;
        //    }

        //    //Utilities.DeleteTempPhoto();
        //}

        public void SaveOutlookPhoto(Contact googleContact, Outlook.ContactItem outlookContact)
        {
            bool hasGooglePhoto = Utilities.HasPhoto(googleContact);
            bool hasOutlookPhoto = Utilities.HasPhoto(outlookContact);

            if (hasGooglePhoto)
            {
                // add google photo to outlook
                //ToDo: add google photo to outlook with new Google API
                //Stream stream = _googleService.GetPhoto(match.GoogleContact);
                Image googlePhoto = Utilities.GetGooglePhoto(this, googleContact);
                if (googlePhoto != null)    // Google may have an invalid photo
                {
                    Utilities.SetOutlookPhoto(outlookContact, googlePhoto);
                    ContactPropertiesUtils.SetOutlookGoogleContactId(this, outlookContact, googleContact);
                    outlookContact.Save();

                    googlePhoto.Dispose();
                }
            }
            else if (hasOutlookPhoto)
            {
                outlookContact.RemovePicture();
                ContactPropertiesUtils.SetOutlookGoogleContactId(this, outlookContact, googleContact);
                outlookContact.Save();
            }
        }

	
		public Group SaveGoogleGroup(Group group)
		{
			//check if this group was not yet inserted on google.
			if (group.GroupEntry.Id.Uri == null)
			{
				//insert group.
				Uri feedUri = new Uri(GroupsQuery.CreateGroupsUri("default"));

				try
				{
					return ContactsRequest.Insert(feedUri, group);
				}
				catch
				{
					//TODO: save google group xml for diagnistics
					throw;
				}
			}
			else
			{
				try
				{
					//group already present in google. just update
					return ContactsRequest.Update(group);
				}
				catch
				{
					//TODO: save google group xml for diagnistics
					throw;
				}
			}
		}

        /// <summary>
        /// Updates Google contact from Outlook (including groups/categories)
        /// </summary>
        public void UpdateContact(Outlook.ContactItem master, Contact slave)
        {
            ContactSync.UpdateContact(master, slave, UseFileAs);
            OverwriteContactGroups(master, slave);
        }

        /// <summary>
        /// Updates Outlook contact from Google (including groups/categories)
        /// </summary>
        public void UpdateContact(Contact master, Outlook.ContactItem slave)
        {
            ContactSync.UpdateContact(master, slave, UseFileAs);
            OverwriteContactGroups(master, slave);

            // -- Immediately save the Outlook contact (including groups) so it can be released, and don't do it in the save loop later
            SaveOutlookContact(ref master, slave);
            SyncedCount++;
            Logger.Log("Updated Outlook contact from Google: \"" + slave.FileAs + "\".", EventType.Information);
        }

        /// <summary>
        /// Updates Google note from Outlook
        /// </summary>
        public void UpdateNote(Outlook.NoteItem master, Document slave)
        {
            if (!string.IsNullOrEmpty(master.Subject))            
                slave.Title = master.Subject.Replace(":", String.Empty);

            string fileName = NotePropertiesUtils.CreateNoteFile(master.EntryID, master.Body, SyncProfile);

            string contentType = MediaFileSource.GetContentTypeForFileName(fileName);

            //ToDo: Somewhow, the content is not uploaded to Google, only an empty document
            //Therefoe I use DocumentService.UploadDocument instead.
            slave.MediaSource = new MediaFileSource(fileName, contentType);

        }

        


        /// <summary>
        /// Updates Outlook contact from Google
        /// </summary>
        public void UpdateNote(Document master, Outlook.NoteItem slave)
        {
            //slave.Subject = master.Title; //The Subject is readonly and set automatically by Outlook
            string body = NotePropertiesUtils.GetBody(this, master);

            if (string.IsNullOrEmpty(body) && slave.Body != null)
            {
                //DialogResult result = MessageBox.Show("The body of Google note '" + master.Title + "' is empty. Do you really want to syncronize an empty Google note to a not yet empty Outlook note?", "Empty Google Note", MessageBoxButtons.YesNo);

                //if (result != DialogResult.Yes)
                //{
                //    Logger.Log("The body of Google note '" + master.Title + "' is empty. The user decided to skip this note and not to syncronize an empty Google note to a not yet empty Outlook note.", EventType.Information);
                    Logger.Log("The body of Google note '" + master.Title + "' is empty. It is skipped from syncing, because Outlook note is not empty.", EventType.Warning);
                    SkippedCount++;
                    return;
                //}
                //Logger.Log("The body of Google note '" + master.Title + "' is empty. The user decided to syncronize an empty Google note to a not yet empty Outlook note (" + slave.Body + ").", EventType.Warning);                
                
            }

            slave.Body = body;

            slave.Categories = string.Empty;
            List<string> newCats = new List<string>();
            foreach (string category in master.ParentFolders)
            {
                Document categoryFolder = GetGoogleFolder(googleNotesFolder, null, category);                

                if (categoryFolder != null)
                    newCats.Add(categoryFolder.Title);
                
            }

            slave.Categories = string.Join(", ", newCats.ToArray());

            NotePropertiesUtils.CreateNoteFile(master.Id, body, SyncProfile);

        }

		/// <summary>
		/// Updates Google contact's groups from Outlook contact
		/// </summary>
		private void OverwriteContactGroups(Outlook.ContactItem master, Contact slave)
		{
			Collection<Group> currentGroups = Utilities.GetGoogleGroups(this, slave);

			// get outlook categories
			string[] cats = Utilities.GetOutlookGroups(master.Categories);

			// remove obsolete groups
			Collection<Group> remove = new Collection<Group>();
			bool found;
			foreach (Group group in currentGroups)
			{
				found = false;
				foreach (string cat in cats)
				{
					if (group.Title == cat)
					{
						found = true;
						break;
					}
				}
				if (!found)
					remove.Add(group);
			}
			while (remove.Count != 0)
			{
				Utilities.RemoveGoogleGroup(slave, remove[0]);
				remove.RemoveAt(0);
			}

			// add new groups
			Group g;
			foreach (string cat in cats)
			{
				if (!Utilities.ContainsGroup(this, slave, cat))
				{
					// add group to contact
					g = GetGoogleGroupByName(cat);
					if (g == null)
					{
						throw new Exception(string.Format("Google Groups were supposed to be created prior to saving", cat));
					}
					Utilities.AddGoogleGroup(slave, g);
				}
			}

            //add system Group My Contacts            
            if (!Utilities.ContainsGroup(this, slave, myContactsGroup))
            {
                // add group to contact
                g = GetGoogleGroupByName(myContactsGroup);
                if (g == null)
                {
                    throw new Exception(string.Format("Google System Group: My Contacts doesn't exist", myContactsGroup));
                }
                Utilities.AddGoogleGroup(slave, g);
            }
		}

		/// <summary>
		/// Updates Outlook contact's categories (groups) from Google groups
		/// </summary>
		private void OverwriteContactGroups(Contact master, Outlook.ContactItem slave)
		{
			Collection<Group> newGroups = Utilities.GetGoogleGroups(this, master);

			List<string> newCats = new List<string>(newGroups.Count);
			foreach (Group group in newGroups)
            {   //Only add groups that are no SystemGroup (e.g. "System Group: Meine Kontakte") automatically tracked by Google
                if (group.Title != null && !group.Title.Equals(myContactsGroup))
				    newCats.Add(group.Title);
			}

			slave.Categories = string.Join(", ", newCats.ToArray());
		}

		/// <summary>
		/// Resets associantions of Outlook contacts with Google contacts via user props
		/// and resets associantions of Google contacts with Outlook contacts via extended properties.
		/// </summary>
		public void ResetContactMatches()
		{
			Debug.Assert(OutlookContacts != null, "Outlook Contacts object is null - this should not happen. Please inform Developers.");
            Debug.Assert(GoogleContacts != null, "Google Contacts object is null - this should not happen. Please inform Developers.");

            try
            {
                if (string.IsNullOrEmpty(SyncProfile))
                {
                    Logger.Log("Must set a sync profile. This should be different on each user/computer you sync on.", EventType.Error);
                    return;
                }
               

			    lock (_syncRoot)
			    {
                    Logger.Log("Resetting Google Contact matches...", EventType.Information);
				    foreach (Contact googleContact in GoogleContacts)
				    {
                        try
                        {
                            if (googleContact != null)
                                ResetMatch(googleContact);
                        }
                        catch (Exception ex)
                        {                           
                            Logger.Log("The match of Google contact " + ContactMatch.GetName(googleContact) + " couldn't be reset: " + ex.Message, EventType.Warning);
                        }
				    }

                    Logger.Log("Resetting Outlook Contact matches...", EventType.Information);
                    //1 based array
                    for (int i=1; i <= OutlookContacts.Count; i++)
                    {
                        Outlook.ContactItem outlookContact = null;

                        try
                        {
                            outlookContact = OutlookContacts[i] as Outlook.ContactItem;
                            if (outlookContact == null)
                            {
                                Logger.Log("Empty Outlook contact found (maybe distribution list). Skipping", EventType.Warning);
                                continue;
                            }
                        }
                        catch (Exception ex)
                        {
                            //this is needed because some contacts throw exceptions
                            Logger.Log("Accessing Outlook contact threw and exception. Skipping: " + ex.Message, EventType.Warning);                               
                            continue;
                        }

                        try
                        {
                            ResetMatch(outlookContact);                            
                        }
                        catch (Exception ex)
                        {
                            Logger.Log("The match of Outlook contact " + outlookContact.FileAs + " couldn't be reset: " + ex.Message, EventType.Warning);
                        }
                    }

                }
            }
            finally
            {
                if (OutlookContacts != null)
                {
                    Marshal.ReleaseComObject(OutlookContacts);
                    OutlookContacts = null;
                }
                GoogleContacts = null;
            }
						
		}

        /// <summary>
        /// Resets associantions of Outlook notes with Google contacts via user props
        /// and resets associantions of Google contacts with Outlook contacts via extended properties.
        /// </summary>
        public void ResetNoteMatches()
        {
            Debug.Assert(OutlookNotes != null, "Outlook Notes object is null - this should not happen. Please inform Developers.");            

            //try
            //{
                if (string.IsNullOrEmpty(SyncProfile))
                {
                    Logger.Log("Must set a sync profile. This should be different on each user/computer you sync on.", EventType.Error);
                    return;
                }


                lock (_syncRoot)
                {
                    Logger.Log("Resetting Google Note matches...", EventType.Information);

                    try
                    {
                        NotePropertiesUtils.DeleteNoteFiles(SyncProfile);
                    }
                    catch (Exception ex)
                    {                           
                        Logger.Log("The Google Note matches couldn't be reset: " + ex.Message, EventType.Warning);
                    }
                    

                    Logger.Log("Resetting Outlook Note matches...", EventType.Information);
                    //1 based array
                    for (int i = 1; i <= OutlookNotes.Count; i++)
                    {
                        Outlook.NoteItem outlookNote = null;

                        try
                        {
                            outlookNote = OutlookNotes[i] as Outlook.NoteItem;
                            if (outlookNote == null)
                            {
                                Logger.Log("Empty Outlook Note found (maybe distribution list). Skipping", EventType.Warning);
                                continue;
                            }
                        }
                        catch (Exception ex)
                        {
                            //this is needed because some notes throw exceptions
                            Logger.Log("Accessing Outlook Note threw and exception. Skipping: " + ex.Message, EventType.Warning);
                            continue;
                        }

                        try
                        {
                            ResetMatch(outlookNote);
                        }
                        catch (Exception ex)
                        {
                            Logger.Log("The match of Outlook note " + outlookNote.Subject + " couldn't be reset: " + ex.Message, EventType.Warning);
                        }
                    }

                }
            //}
            //finally
            //{
            //    if (OutlookContacts != null)
            //    {
            //        Marshal.ReleaseComObject(OutlookContacts);
            //        OutlookContacts = null;
            //    }
            //    GoogleContacts = null;
            //}

        }


        ///// <summary>
        ///// Reset the match link between Google and Outlook contact
        ///// </summary>
        ///// <param name="match"></param>
        //public void ResetMatch(ContactMatch match)
        //{           
        //    if (match == null)
        //        throw new ArgumentNullException("match", "Given ContactMatch is null");
            

        //    if (match.GoogleContact != null)
        //    {
        //        ContactPropertiesUtils.ResetGoogleOutlookContactId(SyncProfile, match.GoogleContact);
        //        SaveGoogleContact(match.GoogleContact);
        //    }
            
        //    if (match.OutlookContact != null)
        //    {
        //        Outlook.ContactItem outlookContactItem = match.OutlookContact.GetOriginalItemFromOutlook(this);
        //        try
        //        {
        //            ContactPropertiesUtils.ResetOutlookGoogleContactId(this, outlookContactItem);
        //            outlookContactItem.Save();
        //        }
        //        finally
        //        {
        //            Marshal.ReleaseComObject(outlookContactItem);
        //            outlookContactItem = null;
        //        }
              
        //        //Reset also Google duplicatesC
        //        foreach (Contact duplicate in match.AllGoogleContactMatches)
        //        {
        //            if (duplicate != match.GoogleContact)
        //            { //To save time, only if not match.GoogleContact, because this was already reset above
        //                ContactPropertiesUtils.ResetGoogleOutlookContactId(SyncProfile, duplicate);
        //                SaveGoogleContact(duplicate);
        //            }
        //        }
        //    }

            
        //}

        /// <summary>
        /// Resets associations of Outlook appointments with Google appointments via user props
        /// and vice versa
        /// </summary>
        public void ResetAppointmentMatches()
        {
            Debug.Assert(OutlookAppointments != null, "Outlook Appointments object is null - this should not happen. Please inform Developers.");

            //try
            //{

            lock (_syncRoot)
            {
                Logger.Log("Resetting Google appointment matches...", EventType.Information);

                for (int i = 0; i < GoogleAppointments.Count; i++)
                {
                    EventEntry googleAppointment = null;

                    try
                    {
                        googleAppointment = GoogleAppointments[i];
                        if (googleAppointment == null)
                        {
                            Logger.Log("Empty Google appointment found (maybe distribution list). Skipping", EventType.Warning);
                            continue;
                        }
                    }
                    catch (Exception ex)
                    {
                        //this is needed because some appointments throw exceptions
                        Logger.Log("Accessing Google appointment threw and exception. Skipping: " + ex.Message, EventType.Warning);
                        continue;
                    }

                    try
                    {
                        ResetMatch(googleAppointment);
                    }
                    catch (Exception ex)
                    {
                        Logger.Log("The match of Google appointment " + googleAppointment.Title.Text + " couldn't be reset: " + ex.Message, EventType.Warning);
                    }
                }


                Logger.Log("Resetting Outlook appointment matches...", EventType.Information);
                //1 based array
                for (int i = 1; i <= OutlookAppointments.Count; i++)
                {
                    Outlook.AppointmentItem outlookAppointment = null;

                    try
                    {
                        outlookAppointment = OutlookAppointments[i] as Outlook.AppointmentItem;
                        if (outlookAppointment == null)
                        {
                            Logger.Log("Empty Outlook appointment found (maybe distribution list). Skipping", EventType.Warning);
                            continue;
                        }
                    }
                    catch (Exception ex)
                    {
                        //this is needed because some appointments throw exceptions
                        Logger.Log("Accessing Outlook appointment threw and exception. Skipping: " + ex.Message, EventType.Warning);
                        continue;
                    }

                    try
                    {
                        ResetMatch(outlookAppointment);
                    }
                    catch (Exception ex)
                    {
                        Logger.Log("The match of Outlook appointment " + outlookAppointment.Subject + " couldn't be reset: " + ex.Message, EventType.Warning);
                    }
                }

            }
            //}
            //finally
            //{
            //    if (OutlookContacts != null)
            //    {
            //        Marshal.ReleaseComObject(OutlookContacts);
            //        OutlookContacts = null;
            //    }
            //    GoogleContacts = null;
            //}

        }

        /// <summary>
        /// Reset the match link between Google and Outlook contact        
        /// </summary>
        public Contact ResetMatch(Contact googleContact)
        {

            if (googleContact != null)
            {
                ContactPropertiesUtils.ResetGoogleOutlookContactId(SyncProfile, googleContact);
                return SaveGoogleContact(googleContact);
            }
            else
                return googleContact;
        }

        public EventEntry ResetMatch(EventEntry googleAppointment)
        {

            if (googleAppointment != null)
            {
                AppointmentPropertiesUtils.ResetGoogleOutlookAppointmentId(SyncProfile, googleAppointment);
                return SaveGoogleAppointment(googleAppointment);
            }
            else
                return googleAppointment;
        }

        /// <summary>
        /// Reset the match link between Outlook and Google contact
        /// </summary>
        public void ResetMatch(Outlook.ContactItem outlookContact)
        {           

            if (outlookContact != null)
            {
                try
                {
                    ContactPropertiesUtils.ResetOutlookGoogleContactId(this, outlookContact);
                    outlookContact.Save();
                }
                finally
                {
                    Marshal.ReleaseComObject(outlookContact);
                    outlookContact = null;
                }
                
            }


        }

        /// <summary>
        /// Reset the match link between Outlook and Google note
        /// </summary>
        public void ResetMatch(Outlook.NoteItem outlookNote)
        {

            if (outlookNote != null)
            {
                //try
                //{
                    NotePropertiesUtils.ResetOutlookGoogleNoteId(this, outlookNote);
                    outlookNote.Save();
                //}
                //finally
                //{
                //    Marshal.ReleaseComObject(outlookNote);
                //    outlookNote = null;
                //}

            }


        }


        /// <summary>
        /// Reset the match link between Outlook and Google appointment
        /// </summary>
        public void ResetMatch(Outlook.AppointmentItem outlookAppointment)
        {

            if (outlookAppointment != null)
            {
                //try
                //{
                AppointmentPropertiesUtils.ResetOutlookGoogleAppointmentId(this, outlookAppointment);
                outlookAppointment.Save();
                //}
                //finally
                //{
                //    Marshal.ReleaseComObject(OutlookAppointment);
                //    OutlookAppointment = null;
                //}

            }


        }

        public ContactMatch ContactByProperty(string name, string email)
        {            
            foreach (ContactMatch m in Contacts)
            {
                if (m.GoogleContact != null &&
                    ((m.GoogleContact.PrimaryEmail != null && m.GoogleContact.PrimaryEmail.Address == email) ||
                    m.GoogleContact.Title == name ||
                    m.GoogleContact.Name != null && m.GoogleContact.Name.FullName == name))
                {
                    return m;
                }
                else if (m.OutlookContact != null && (
                    (m.OutlookContact.Email1Address != null && m.OutlookContact.Email1Address == email) ||
                    m.OutlookContact.FileAs == name))
                {
                    return m;
                }
            }
            return null;
        }
		
        //public ContactMatch ContactEmail(string email)
        //{
        //    foreach (ContactMatch m in Contacts)
        //    {
        //        if (m.GoogleContact != null &&
        //            (m.GoogleContact.PrimaryEmail != null && m.GoogleContact.PrimaryEmail.Address == email))
        //        {
        //            return m;
        //        }
        //        else if (m.OutlookContact != null && (
        //            m.OutlookContact.Email1Address != null && m.OutlookContact.Email1Address == email))
        //        {
        //            return m;
        //        }
        //    }
        //    return null;
        //}

		/// <summary>
		/// Used to find duplicates.
		/// </summary>
		/// <param name="name"></param>
		/// <param name="value"></param>
		/// <returns></returns>
		public Collection<OutlookContactInfo> OutlookContactByProperty(string name, string value)
		{
            Collection<OutlookContactInfo> col = new Collection<OutlookContactInfo>();
            //foreach (Outlook.ContactItem outlookContact in OutlookContacts)
            //{
            //    if (outlookContact != null && (
            //        (outlookContact.Email1Address != null && outlookContact.Email1Address == email) ||
            //        outlookContact.FileAs == name))
            //    {
            //        col.Add(outlookContact);
            //    }
            //}
            Outlook.ContactItem item = null;
            try
            {
                item = OutlookContacts.Find("["+name+"] = \"" + value + "\"") as Outlook.ContactItem;
                while (item != null)
                {
                    col.Add(new OutlookContactInfo(item, this));
                    Marshal.ReleaseComObject(item);
                    item = OutlookContacts.FindNext() as Outlook.ContactItem;
                }
            }
            catch (Exception)
			{
				//TODO: should not get here.
			}

			return col;
		}
        
		public Group GetGoogleGroupById(string id)
		{
			//return GoogleGroups.FindById(new AtomId(id)) as Group;
            foreach (Group group in GoogleGroups)
            {
                if (group.GroupEntry.Id.Equals(new AtomId(id)))
                    return group;
            }
            return null;
		}

		public Group GetGoogleGroupByName(string name)
		{
			foreach (Group group in GoogleGroups)
			{
				if (group.Title == name)
					return group;
			}
			return null;
		}

        public Contact GetGoogleContactById(string id)
        {
            foreach (Contact contact in GoogleContacts)
            {
                if (contact.ContactEntry.Id.Equals(new AtomId(id)))
                    return contact;
            }
            return null;
        }

        public Document GetGoogleNoteById(string id)
        {
            foreach (Document note in GoogleNotes)
            {
                if (note.DocumentEntry.Id.Equals(new AtomId(id)))
                    return note;
            }
            return null;
        }

        public EventEntry GetGoogleAppointmentById(string id)
        {
            foreach (EventEntry appointment in GoogleAppointments)
            {
                if (appointment.Id.Equals(new AtomId(id)))
                    return appointment;
            }
            return null;
        }

		public Group CreateGroup(string name)
		{
			Group group = new Group();
			group.Title = name;
			group.GroupEntry.Dirty = true;
			return group;
		}

		public static bool AreEqual(Outlook.ContactItem c1, Outlook.ContactItem c2)
		{
			return c1.Email1Address == c2.Email1Address;
		}
		public static int IndexOf(Collection<Outlook.ContactItem> col, Outlook.ContactItem outlookContact)
		{

			for (int i = 0; i < col.Count; i++)
			{
				if (AreEqual(col[i], outlookContact))
					return i;
			}
			return -1;
		}

		internal void DebugContacts()
		{
			string msg = "DEBUG INFORMATION\nPlease submit to developer:\n\n{0}\n{1}\n{2}";
          

            if (SyncContacts)
            {
                string oCount = "Outlook Contact Count: " + OutlookContacts.Count;
                string gCount = "Google Contact Count: " + GoogleContacts.Count;
                string mCount = "Matches Count: " + Contacts.Count;

                MessageBox.Show(string.Format(msg, oCount, gCount, mCount), "DEBUG INFO", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            if (SyncNotes)
            {
                string oCount = "Outlook Notes Count: " + OutlookNotes.Count;
                string gCount = "Google Notes Count: " + GoogleNotes.Count;
                string mCount = "Matches Count: " + Notes.Count;

                MessageBox.Show(string.Format(msg, oCount, gCount, mCount), "DEBUG INFO", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            if (SyncAppointments)
            {
                string oCount = "Outlook appointments Count: " + OutlookAppointments.Count;
                string gCount = "Google appointments Count: " + GoogleAppointments.Count;
                string mCount = "Matches Count: " + Appointments.Count;

                MessageBox.Show(string.Format(msg, oCount, gCount, mCount), "DEBUG INFO", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
		}
        public static Outlook.ContactItem CreateOutlookContactItem(string syncContactsFolder)
        {
            //outlookContact = OutlookApplication.CreateItem(Outlook.OlItemType.olContactItem) as Outlook.ContactItem; //This will only create it in the default folder, but we have to consider the selected folder
            Outlook.ContactItem outlookContact = null;
            Outlook.MAPIFolder contactsFolder = null;
            Outlook.Items items = null;

            try
            {
                contactsFolder = OutlookNameSpace.GetFolderFromID(syncContactsFolder);
                items = contactsFolder.Items;
                outlookContact = items.Add(Outlook.OlItemType.olContactItem) as Outlook.ContactItem;
            }
            finally
            {
                if (items != null) Marshal.ReleaseComObject(items);
                if (contactsFolder != null) Marshal.ReleaseComObject(contactsFolder);
            }
            return outlookContact;
        }

        public static Outlook.NoteItem CreateOutlookNoteItem(string syncNotesFolder)
        {
            //outlookNote = OutlookApplication.CreateItem(Outlook.OlItemType.olNoteItem) as Outlook.NoteItem; //This will only create it in the default folder, but we have to consider the selected folder
            Outlook.NoteItem outlookNote = null;
            Outlook.MAPIFolder notesFolder = null;
            Outlook.Items items = null;

            try
            {
                notesFolder = OutlookNameSpace.GetFolderFromID(syncNotesFolder);
                items = notesFolder.Items;
                outlookNote = items.Add(Outlook.OlItemType.olNoteItem) as Outlook.NoteItem;
            }
            finally
            {
                if (items != null) Marshal.ReleaseComObject(items);
                if (notesFolder != null) Marshal.ReleaseComObject(notesFolder);
            }
            return outlookNote;
        }
	

        public static Outlook.AppointmentItem CreateOutlookAppointmentItem(string syncAppointmentsFolder)
        {
            //OutlookAppointment = OutlookApplication.CreateItem(Outlook.OlItemType.olAppointmentItem) as Outlook.AppointmentItem; //This will only create it in the default folder, but we have to consider the selected folder
            Outlook.AppointmentItem outlookAppointment = null;
            Outlook.MAPIFolder appointmentsFolder = null;
            Outlook.Items items = null;

            try
            {
                appointmentsFolder = OutlookNameSpace.GetFolderFromID(syncAppointmentsFolder);
                items = appointmentsFolder.Items;
                outlookAppointment = items.Add(Outlook.OlItemType.olAppointmentItem) as Outlook.AppointmentItem;
            }
            finally
            {
                if (items != null) Marshal.ReleaseComObject(items);
                if (appointmentsFolder != null) Marshal.ReleaseComObject(appointmentsFolder);
            }
            return outlookAppointment;
        }
	}

	internal enum SyncOption
	{
		MergePrompt,
		MergeOutlookWins,
		MergeGoogleWins,
		OutlookToGoogleOnly,
		GoogleToOutlookOnly,
	}
}
