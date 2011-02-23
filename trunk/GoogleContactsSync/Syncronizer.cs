using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using Google.GData.Contacts;
using Google.GData.Client;
using Google.GData.Extensions;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Drawing;
using System.IO;
using System.Collections.ObjectModel;
using Microsoft.Win32;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace GoContactSyncMod
{
	internal class Syncronizer
	{
		public const int OutlookUserPropertyMaxLength = 32;
		public const string OutlookUserPropertyTemplate = "g/con/{0}/";
		private static object _syncRoot = new object();

		private int _totalCount;
		public int TotalCount
		{
			get { return _totalCount; }
		}

		private int _syncedCount;
		public int SyncedCount
		{
			get { return _syncedCount; }
		}

		private int _deletedCount;
		public int DeletedCount
		{
			get { return _deletedCount; }
		}

        private int _errorCount;
        public int ErrorCount
        {
            get { return _errorCount; }
        }

        private int _skippedCount;
        public int SkippedCount
        {
            set { _skippedCount = value; }
            get { return _skippedCount; }
        }


		public delegate void NotificationHandler(string title, string message, EventType eventType);
		public delegate void ErrorNotificationHandler(string title, Exception ex, EventType eventType);
		public event NotificationHandler DuplicatesFound;
		public event ErrorNotificationHandler ErrorEncountered;

		private ContactsService _googleService;
		public ContactsService GoogleService
		{
			get { return _googleService; }
		}

		private Outlook.NameSpace _outlookNamespace;

		private Outlook.Application _outlookApp;
		public Outlook.Application OutlookApplication
		{
			get { return _outlookApp; }
		}

		private Outlook.Items _outlookContacts;
		public Outlook.Items OutlookContacts
		{
			get { return _outlookContacts; }
		}

        private Collection<ContactMatch> _outlookContactDuplicates;
        public Collection<ContactMatch> OutlookContactDuplicates
        {
            get { return _outlookContactDuplicates; }
            set { _outlookContactDuplicates = value; }
        }

        private Collection<ContactMatch> _googleContactDuplicates;
        public Collection<ContactMatch> GoogleContactDuplicates
        {
            get { return _googleContactDuplicates; }
            set { _googleContactDuplicates = value; }
        }

		private AtomEntryCollection _googleContacts;
		public AtomEntryCollection GoogleContacts
		{
			get { return _googleContacts; }
		}

		private AtomEntryCollection _googleGroups;
		public AtomEntryCollection GoogleGroups
		{
			get { return _googleGroups; }
		}

		private string _propertyPrefix;
		public string OutlookPropertyPrefix
		{
			get { return _propertyPrefix; }
		}

		public string OutlookPropertyNameId
		{
			get { return _propertyPrefix + "id"; }
		}

		/*public string OutlookPropertyNameUpdated
		{
			get { return _propertyPrefix + "up"; }
		}*/

		public string OutlookPropertyNameSynced
		{
			get { return _propertyPrefix + "up"; }
		}

		private SyncOption _syncOption = SyncOption.MergeOutlookWins;
		public SyncOption SyncOption
		{
			get { return _syncOption; }
			set { _syncOption = value; }
		}

		private string _syncProfile = "";
		public string SyncProfile
		{
			get { return _syncProfile; }
			set { _syncProfile = value; }
		}

		//private ConflictResolution? _conflictResolution;
		//public ConflictResolution? CResolution
		//{
		//    get { return _conflictResolution; }
		//    set { _conflictResolution = value; }
		//}

		private ContactMatchList _matches;
		public ContactMatchList Contacts
		{
			get { return _matches; }
		}

		private string _authToken;
		public string AuthToken
		{
			get
			{
				return _authToken;
			}
		}

		private bool _syncDelete;
		/// <summary>
		/// If true deletes contacts if synced before, but one is missing. Otherwise contacts will bever be automatically deleted
		/// </summary>
		public bool SyncDelete
		{
			get { return _syncDelete; }
			set { _syncDelete = value; }
		}


		public Syncronizer()
		{

		}

		public Syncronizer(SyncOption syncOption)
		{
			_syncOption = syncOption;
		}

		public void LoginToGoogle(string username, string password)
		{
			Logger.Log("Connecting to Google...", EventType.Information);
			if (_googleService == null)
				_googleService = new ContactsService("GoogleContactSyncMod");

			_googleService.setUserCredentials(username, password);
			_authToken = _googleService.QueryAuthenticationToken();

			int maxUserIdLength = Syncronizer.OutlookUserPropertyMaxLength - (Syncronizer.OutlookUserPropertyTemplate.Length - 3 + 2);//-3 = to remove {0}, +2 = to add length for "id" or "up"
			string userId = _googleService.Credentials.Username;
			if (userId.Length > maxUserIdLength)
				userId = userId.GetHashCode().ToString("X"); //if a user id would overflow UserProperty name, then use that user id hash code as id.

			_propertyPrefix = string.Format(Syncronizer.OutlookUserPropertyTemplate, userId);
		}

		public void LoginToOutlook()
		{
			Logger.Log("Connecting to Outlook...", EventType.Information);

			try
			{
				if (_outlookApp == null)
				{
					try
					{
						_outlookApp = new Outlook.Application();
					}
					catch (Exception ex)
					{
						throw new NotSupportedException("Could not create instance of 'Microsoft Outlook'. Make sure Outlook 2003 or above version is installed and retry.", ex);
					}

					try
					{
						_outlookNamespace = _outlookApp.GetNamespace("mapi");
					}
					catch (COMException comEx)
					{
						throw new NotSupportedException("Could not conncet to 'Microsoft Outlook'. Make sure Outlook 2003 or above version is installed and running.", comEx);
					}
				}

				// Get default profile name from registry, as this is not always "Outlook" and would popup a dialog to choose profile
				// no matter if default profile is set or not. So try to read the default profile, fallback is still "Outlook"
				string profileName = "Outlook";
				using (RegistryKey k = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\Outlook\SocialConnector", false))
				{
					if (k != null)
						profileName = k.GetValue("PrimaryOscProfile", "Outlook").ToString();
				}
				_outlookNamespace.Logon(profileName, null, true, false);
			}
			catch (System.Runtime.InteropServices.COMException)
			{
				try
				{
					// If outlook was closed/terminated inbetween, we will receive an Exception
					// System.Runtime.InteropServices.COMException (0x800706BA): The RPC server is unavailable. (Exception from HRESULT: 0x800706BA)
					// so recreate outlook instance
					Logger.Log("Cannot connect to Outlook, creating new instance....", EventType.Information);
					_outlookApp = new Outlook.Application();
					_outlookNamespace = _outlookApp.GetNamespace("mapi");
					_outlookNamespace.Logon();
				}
				catch (Exception ex)
				{
					string message = String.Format("Cannot connect to Outlook: {0}. \nPlease restart GO Contact Sync Mod and try again. If error persists, please inform developers on SourceForge.", ex.Message);
					// Error again? We need full stacktrace, display it!
					ErrorHandler.Handle(new ApplicationException(message, ex));
				}
			}
		}

		public void LogoffOutlook()
		{
			try
			{
				Logger.Log("Disconnecting from Outlook...", EventType.Information);
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
		}

		public void LoadOutlookContacts()
		{
			Logger.Log("Loading Outlook contacts...", EventType.Information);
			Outlook.MAPIFolder contactsFolder = _outlookNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
            _outlookContacts = contactsFolder.Items;
		}
        ///// <summary>
        ///// Moves duplicates from _outlookContacts to _outlookContactDuplicates
        ///// </summary>
        //private void FilterOutlookContactDuplicates()
        //{
        //    _outlookContactDuplicates = new Collection<Outlook.ContactItem>();
            
        //    if (_outlookContacts.Count < 2)
        //        return;

        //    Outlook.ContactItem main, other;
        //    bool found = true;
        //    int index = 0;

        //    while (found)
        //    {
        //        found = false;

        //        for (int i = index; i <= _outlookContacts.Count - 1; i++)
        //        {
        //            main = _outlookContacts[i] as Outlook.ContactItem;

        //            // only look forward
        //            for (int j = i + 1; j <= _outlookContacts.Count; j++)
        //            {
        //                other = _outlookContacts[j] as Outlook.ContactItem;

        //                if (other.FileAs == main.FileAs &&
        //                    other.Email1Address == main.Email1Address)
        //                {
        //                    _outlookContactDuplicates.Add(other);
        //                    _outlookContacts.Remove(j);
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

		public void LoadGoogleContacts()
		{
			try
			{
				Logger.Log("Loading Google Contacts...", EventType.Information);
				ContactsQuery query = new ContactsQuery(ContactsQuery.CreateContactsUri("default"));
				query.NumberToRetrieve = 256;
				query.StartIndex = 0;
				query.ShowDeleted = false;
				//query.OrderBy = "lastmodified";

				ContactsFeed feed;
				feed = _googleService.Query(query);
				_googleContacts = feed.Entries;
				while (feed.Entries.Count == query.NumberToRetrieve)
				{
                    query.StartIndex = _googleContacts.Count +1;
					feed = _googleService.Query(query);
					foreach (AtomEntry a in feed.Entries)
					{
						_googleContacts.Add(a);
					}
				}
			}
			catch (System.Net.WebException ex)
			{
				string message = string.Format("Cannot connect to Google: {0}. \nPlease ensure you are connected to the internet. If you are behind a proxy, change your proxy configuration!", ex.Message);
				Logger.Log(message, EventType.Error);
			}
		}
		public void LoadGoogleGroups()
		{
			Logger.Log("Loading Google Groups...", EventType.Information);
			GroupsQuery query = new GroupsQuery(GroupsQuery.CreateGroupsUri("default"));
			query.NumberToRetrieve = 256;
			query.StartIndex = 0;
			query.ShowDeleted = false;

			GroupsFeed feed;
			feed = _googleService.Query(query);
			_googleGroups = feed.Entries;
			while (feed.Entries.Count == query.NumberToRetrieve)
			{
				query.StartIndex = _googleGroups.Count + 1;
				feed = _googleService.Query(query);
				foreach (AtomEntry a in feed.Entries)
				{
					_googleGroups.Add(a);
				}
			}
		}

       

        /// <summary>
        /// Load the contacts from Google and Outlook and match them
        /// </summary>
		public void Load()
		{
			LoadOutlookContacts();
			LoadGoogleContacts();
			LoadGoogleGroups();

			DuplicateDataException duplicateDataException;
			_matches = ContactsMatcher.MatchContacts(this, out duplicateDataException);
			if (duplicateDataException != null)
			{
				Logger.Log(duplicateDataException.Message, EventType.Error);
				if (DuplicatesFound != null)
					DuplicatesFound("Google duplicates found", duplicateDataException.Message, EventType.Error);
			}
		}

		public void Sync()
		{
			lock (_syncRoot)
			{
                if (_syncProfile.Length == 0)
                {
                    Logger.Log("Must set a sync profile. This should be different on each user/computer you sync on.", EventType.Error);
                    return;
                }

				_syncedCount = 0;
				_deletedCount = 0;
                _errorCount = 0;
                _skippedCount = 0;

				Load();

#if debug
			this.DebugContacts();
#endif

				if (_matches == null)
					return;

                _totalCount = _matches.Count;

                //Remove Google duplicates from matches to be synced
                if (_googleContactDuplicates != null)
                    foreach (ContactMatch match in _googleContactDuplicates)
                        if (_matches.Contains(match))
                        {
                            _skippedCount++;
                            _matches.Remove(match);
                        }

                //Remove Outlook duplicates from matches to be synced
                if (_outlookContactDuplicates != null)
                    foreach (ContactMatch match in _outlookContactDuplicates)
                        if (_matches.Contains(match))
                        {
                            _skippedCount++;
                            _matches.Remove(match);
                        }

				Logger.Log("Syncing groups...", EventType.Information);
				ContactsMatcher.SyncGroups(this);

				Logger.Log("Syncing contacts...", EventType.Information);
				ContactsMatcher.SyncContacts(this);

				SaveContacts(_matches);
			}
		}

		public void SaveContacts(ContactMatchList contacts)
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
                        _errorCount++;
                        _syncedCount--;
                        string message = String.Format("Failed to synchronize contact: {0}. \nPlease check the contact, if any Email already exists on Google contacts side or if there is too much or invalid data in the notes field. \nIf the problem persists, please try recreating the contact or report the error on SourceForge.", match.OutlookContact.FullNameAndCompany);
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
                if (match.GoogleContact.IsDirty() || !match.OutlookContact.Saved)
                    _syncedCount++;
				
				if (match.GoogleContact.IsDirty())// || googleChanged)
				{
					//google contact was modified. save.
					SaveGoogleContact(match);
					Logger.Log("Updated Google contact from Outlook: \"" + match.GoogleContact.Title.Text + "\".", EventType.Information);
				}

				if (!match.OutlookContact.Saved)// || outlookChanged)
				{
                    //outlook contact was modified. save.
                    SaveOutlookContact(match);
					Logger.Log("Updated Outlook contact from Google: \"" + match.OutlookContact.FileAs + "\".", EventType.Information);

					//TODO: this will cause the google contact to be updated on next run because Outlook's contact will be marked as saved later that Google's contact.
				}                

				// save photos
				//SaveContactPhotos(match);
			}
			else if (match.GoogleContact == null && match.OutlookContact != null)
			{
				if (ContactPropertiesUtils.GetOutlookGoogleContactId(this, match.OutlookContact) != null)
				{
                    string name = match.OutlookContact.FileAs;
                    if (_syncOption == SyncOption.OutlookToGoogleOnly)
                    {
                        _skippedCount++;
                        Logger.Log("Skipped Deletion of Outlook contact because of SyncOption " + _syncOption + ":" + name + ".", EventType.Information);
                    }
                    else if (!_syncDelete)
                    {
                        _skippedCount++;
                        Logger.Log("Skipped Deletion of Outlook contact because SyncDeletion is switched off: " + name + ".", EventType.Information);
                    }
                    else
                    {
                        // peer google contact was deleted, delete outlook contact
                        match.OutlookContact.Delete();
                        _deletedCount++;
                        Logger.Log("Deleted Outlook contact: \"" + name + "\".", EventType.Information);
                    }
				}
			}
			else if (match.GoogleContact != null && match.OutlookContact == null)
			{
				if (ContactPropertiesUtils.GetGoogleOutlookContactId(SyncProfile, match.GoogleContact) != null)
				{
                    if (_syncOption == SyncOption.GoogleToOutlookOnly)
                    {
                        _skippedCount++;
                        Logger.Log("Skipped Deletion of Google contact because of SyncOption " + _syncOption + ":" + match.GoogleContact.Title.Text + ".", EventType.Information);
                    }
                    else if (!_syncDelete)
                    {
                        _skippedCount++;
                        Logger.Log("Skipped Deletion of Google contact because SyncDeletion is switched off :" + match.GoogleContact.Title.Text + ".", EventType.Information);
                    }
                    else
                    {
                        // peer outlook contact was deleted, delete google contact
                        match.GoogleContact.Delete();
                        _deletedCount++;
                        Logger.Log("Deleted Google contact: \"" + match.GoogleContact.Title.Text + "\".", EventType.Information);
                    }
				}
			}
			else
			{
				//TODO: ignore for now: throw new ArgumentNullException("To save contacts both ContactMatch peers must be present.");
				Logger.Log("Both Google and Outlook contact: \"" + match.GoogleContact.Title.Text + "\" have been changed! Not implemented yet.", EventType.Warning);
			}
		}

        private void SaveOutlookContact(ContactMatch match)
        {
            match.OutlookContact.Save();
            ContactPropertiesUtils.SetGoogleOutlookContactId(SyncProfile, match.GoogleContact, match.OutlookContact);

            ContactEntry updatedEntry;
            try
            {
                updatedEntry = match.GoogleContact.Update() as ContactEntry;
            }
            catch (GDataRequestException tmpEx)
            {
                // check if it's the known HTCData problem, or if there is any invalid XML element or any unescaped XML sequence
                if (tmpEx.ResponseString.Contains("HTCData") || tmpEx.ResponseString.Contains("&#39") || match.GoogleContact.Content.Content.Contains("<"))
                {
                    bool wasDirty = match.GoogleContact.Content.Dirty;
                    // XML escape the content
                    match.GoogleContact.Content.Content = EscapeXml(match.GoogleContact.Content.Content);
                    // set dirty to back, cause we don't want the changed content go back to Google without reason
                    match.GoogleContact.Content.Dirty = wasDirty;
                    updatedEntry = match.GoogleContact.Update() as ContactEntry;
                }
                else if (!String.IsNullOrEmpty(tmpEx.ResponseString))
                    throw new ApplicationException(tmpEx.ResponseString, tmpEx);
                else
                    throw;
            }
            match.GoogleContact = updatedEntry;

            ContactPropertiesUtils.SetOutlookGoogleContactId(this, match.OutlookContact, match.GoogleContact);
            match.OutlookContact.Save();
            SaveOutlookPhoto(match);
        }
		private string EscapeXml(string xml)
		{
			string encodedXml = System.Security.SecurityElement.Escape(xml);
			return encodedXml;
		}
		public void SaveGoogleContact(ContactMatch match)
		{
			//check if this contact was not yet inserted on google.
			if (match.GoogleContact.Id.Uri == null)
			{
				//insert contact.
				Uri feedUri = new Uri(ContactsQuery.CreateContactsUri("default"));

				try
				{
					ContactPropertiesUtils.SetGoogleOutlookContactId(SyncProfile, match.GoogleContact, match.OutlookContact);

                    //ToDo: This will fail, if another account with the same email already exists
					ContactEntry createdEntry = (ContactEntry)_googleService.Insert(feedUri, match.GoogleContact);
					match.GoogleContact = createdEntry;

					ContactPropertiesUtils.SetOutlookGoogleContactId(this, match.OutlookContact, match.GoogleContact);
					match.OutlookContact.Save();

                    SaveGooglePhoto(match);
				}
				catch (Exception ex)
				{
					string xml = GetContactXml(match.GoogleContact);
					string newEx = String.Format("Error saving NEW Google contact: {0}. \n{1}", ex.Message, xml);
					throw new ApplicationException(newEx, ex);
				}
			}
			else
			{
				try
				{
					//contact already present in google. just update
					ContactPropertiesUtils.SetGoogleOutlookContactId(SyncProfile, match.GoogleContact, match.OutlookContact);

					//TODO: this will fail if original contact had an empty name or primary email address.
					ContactEntry updatedEntry = match.GoogleContact.Update() as ContactEntry;
					match.GoogleContact = updatedEntry;

					ContactPropertiesUtils.SetOutlookGoogleContactId(this, match.OutlookContact, match.GoogleContact);
					match.OutlookContact.Save();

                    SaveGooglePhoto(match);
				}
				catch (Exception ex)
				{
					//match.GoogleContact.Summary
					string xml = GetContactXml(match.GoogleContact);
					string newEx = String.Format("Error saving EXISTING Google contact: {0}. \n{1}", ex.Message, xml);
					throw new ApplicationException(newEx, ex);
				}
			}
		}

		private string GetContactXml(ContactEntry contactEntry)
		{
			MemoryStream ms = new MemoryStream();
			contactEntry.SaveToXml(ms);
			StreamReader sr = new StreamReader(ms);
			ms.Seek(0, SeekOrigin.Begin);
			string xml = sr.ReadToEnd();
			return xml;
		}

        /// <summary>
        /// Only save the google contact without photo update
        /// </summary>
        /// <param name="googleContact"></param>
		public void SaveGoogleContact(ContactEntry googleContact)
		{
			//check if this contact was not yet inserted on google.
			if (googleContact.Id.Uri == null)
			{
				//insert contact.
				Uri feedUri = new Uri(ContactsQuery.CreateContactsUri("default"));

				try
				{
					ContactEntry createdEntry = (ContactEntry)_googleService.Insert(feedUri, googleContact);
				}
				catch
				{
					//TODO: save google contact xml for diagnistics
					throw;
				}
			}
			else
			{
				try
				{
					//contact already present in google. just update
					//TODO: this will fail if original contact had an empty name or rpimary email address.
					ContactEntry updatedEntry = googleContact.Update() as ContactEntry;
				}
				catch
				{
					//TODO: save google contact xml for diagnistics
					throw;
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

        public void SaveGooglePhoto(ContactMatch match)
        {
            bool hasGooglePhoto = Utilities.HasPhoto(match.GoogleContact);
            bool hasOutlookPhoto = Utilities.HasPhoto(match.OutlookContact);

            if (hasOutlookPhoto)
            {
                // add outlook photo to google
                Image outlookPhoto = Utilities.GetOutlookPhoto(match.OutlookContact);
                if (outlookPhoto != null)
                {
                    //outlookPhoto = Utilities.CropImageGoogleFormat(outlookPhoto);
                    bool saved = Utilities.SaveGooglePhoto(this, match.GoogleContact, outlookPhoto);
                    if (!saved)
                        throw new Exception("Could not save");

                    //Just save the Outlook Contact to have the same lastUpdate date as Google
                    match.OutlookContact.Save();
                    outlookPhoto.Dispose();
                }
            }
            else if (hasGooglePhoto)
            {                
                //ToDo: Delete Photo on Google side, if no Outlook photo exists
                //match.GoogleContact.PhotoUri = null;
            }

            //Utilities.DeleteTempPhoto();
        }

        public void SaveOutlookPhoto(ContactMatch match)
        {
            bool hasGooglePhoto = Utilities.HasPhoto(match.GoogleContact);
            bool hasOutlookPhoto = Utilities.HasPhoto(match.OutlookContact);

            if (hasGooglePhoto)
            {
                // add google photo to outlook
                Image googlePhoto = Utilities.GetGooglePhoto(this, match.GoogleContact);
                Utilities.SetOutlookPhoto(match.OutlookContact, googlePhoto);
                match.OutlookContact.Save();

                googlePhoto.Dispose();
            }
            else if (hasOutlookPhoto)
            {
                match.OutlookContact.RemovePicture();
                match.OutlookContact.Save();
            }
        }

		
		public GroupEntry SaveGoogleGroup(GroupEntry group)
		{
			//check if this group was not yet inserted on google.
			if (group.Id.Uri == null)
			{
				//insert group.
				Uri feedUri = new Uri(GroupsQuery.CreateGroupsUri("default"));

				try
				{
					GroupEntry createdEntry = _googleService.Insert(feedUri, group) as GroupEntry;
					return createdEntry;
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
					GroupEntry updatedEntry = group.Update() as GroupEntry;
					return updatedEntry;
				}
				catch
				{
					//TODO: save google group xml for diagnistics
					throw;
				}
			}
		}

		
		/// <summary>
		/// Updates Google contact's groups
		/// </summary>
		/// <param name="googleContact"></param>
		/// <param name="currentGroups"></param>
		/// <param name="newGroups"></param>
		public void OverwriteContactGroups(Outlook.ContactItem master, ContactEntry slave)
		{
			Collection<GroupEntry> currentGroups = Utilities.GetGoogleGroups(this, slave);

			// get outlook categories
			string[] cats = Utilities.GetOutlookGroups(master);

			// remove obsolete groups
			Collection<GroupEntry> remove = new Collection<GroupEntry>();
			bool found;
			foreach (GroupEntry group in currentGroups)
			{
				found = false;
				foreach (string cat in cats)
				{
					if (group.Title.Text == cat)
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
			GroupEntry g;
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
		}

		/// <summary>
		/// Updates Outlook contact's categories (groups)
		/// </summary>
		/// <param name="outlookContact"></param>
		/// <param name="currentGroups"></param>
		/// <param name="newGroups"></param>
		public void OverwriteContactGroups(ContactEntry master, Outlook.ContactItem slave)
		{
			Collection<GroupEntry> newGroups = Utilities.GetGoogleGroups(this, master);

			List<string> newCats = new List<string>(newGroups.Count);
			foreach (GroupEntry group in newGroups)
			{
				newCats.Add(group.Title.Text);
			}

			slave.Categories = string.Join(", ", newCats.ToArray());
		}

		/// <summary>
		/// Resets associantions of Outlook contacts with Google contacts via user props
		/// and resets associantions of Google contacts with Outlook contacts via extended properties.
		/// </summary>
		public void ResetMatches()
		{
			Debug.Assert(Contacts != null, "Contacts object is null - this should not happen. Please inform Developers.");

            if (_syncProfile.Length == 0)
            {
                Logger.Log("Must set a sync profile. This should be different on each user/computer you sync on.", EventType.Error);
                return;
            }

            Logger.Log("Resetting matches...", EventType.Information);

			lock (_syncRoot)
			{
				foreach (ContactMatch match in Contacts)
				{
                    try
                    {
                        ResetMatch(match);
                    }
                    catch
                    {
                        Logger.Log("The match of contact " + ((match.OutlookContact!=null)?match.OutlookContact.FileAs:match.GoogleContact.Title.Text) + " couldn't be reset", EventType.Warning);
                    }
				}

				Debug.Assert(Contacts != null, "Contacts object is null after reset - this should not happen. Please inform Developers.");
				Contacts.Clear();
			}
		}

        /// <summary>
        /// Reset the match link between Google and Outlook contact
        /// </summary>
        /// <param name="match"></param>
		public void ResetMatch(ContactMatch match)
		{           
			if (match == null)
				throw new ArgumentNullException("match", "Given ContactMatch is null");
            

            if (match.GoogleContact != null)
            {
                ContactPropertiesUtils.ResetGoogleOutlookContactId(SyncProfile, match.GoogleContact);
                SaveGoogleContact(match.GoogleContact);
			}
            
            if (match.OutlookContact != null)
            {
                ContactPropertiesUtils.ResetOutlookGoogleContactId(this, match.OutlookContact);
				match.OutlookContact.Save();
                
                //Reset also Google duplicates
                foreach (ContactEntry duplicate in match.AllGoogleContactMatches)
                {
                    if (duplicate != match.GoogleContact)
                    { //To save time, only if not match.GoogleContact, because this was already reset above
                        ContactPropertiesUtils.ResetGoogleOutlookContactId(SyncProfile, duplicate);
                        SaveGoogleContact(duplicate);
                    }
                }
			}

            
		}

		public ContactMatch ContactByProperty(string name, string email)
		{
			foreach (ContactMatch m in Contacts)
			{
				if (m.GoogleContact != null &&
					((m.GoogleContact.PrimaryEmail != null && m.GoogleContact.PrimaryEmail.Address == email) ||
					m.GoogleContact.Title.Text == name))
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
		public ContactMatch ContactEmail(string email)
		{
			foreach (ContactMatch m in Contacts)
			{
				if (m.GoogleContact != null &&
					(m.GoogleContact.PrimaryEmail != null && m.GoogleContact.PrimaryEmail.Address == email))
				{
					return m;
				}
				else if (m.OutlookContact != null && (
					m.OutlookContact.Email1Address != null && m.OutlookContact.Email1Address == email))
				{
					return m;
				}
			}
			return null;
		}

		/// <summary>
		/// Used to find duplicates.
		/// </summary>
		/// <param name="name"></param>
		/// <param name="value"></param>
		/// <returns></returns>
		public Collection<Outlook.ContactItem> OutlookContactByProperty(string name, string value)
		{
			Collection<Outlook.ContactItem> col = new Collection<Outlook.ContactItem>();
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
                if (item != null)
                {
                    col.Add(item);
                    do
                    {
                        item = OutlookContacts.FindNext() as Outlook.ContactItem;
                        if (item != null)
                            col.Add(item);
                    } while (item != null);
                }
            }
            catch (Exception)
			{
				//TODO: should not get here.
			}

			return col;
		}
		/// <summary>
		/// Used to find duplicates.
		/// </summary>
		/// <param name="name"></param>
		/// <param name="email"></param>
		/// <returns></returns>
		public Collection<Outlook.ContactItem> OutlookContactByEmail(string email)
		{
			//TODO: optimise by using OutlookContacts.Find

			Collection<Outlook.ContactItem> col = new Collection<Outlook.ContactItem>();
			Outlook.ContactItem item = null;
			try
			{
				item = OutlookContacts.Find("[Email1Address] = \"" + email + "\"") as Outlook.ContactItem;
				if (item != null)
				{
					col.Add(item);
					do
					{
						item = OutlookContacts.FindNext() as Outlook.ContactItem;
						if (item != null)
							col.Add(item);
					} while (item != null);
				}

                item = OutlookContacts.Find("[Email2Address] = \"" + email + "\"") as Outlook.ContactItem;
                if (item != null)
                {
                    col.Add(item);
                    do
                    {
                        item = OutlookContacts.FindNext() as Outlook.ContactItem;
                        if (item != null)
                            col.Add(item);
                    } while (item != null);
                }

                item = OutlookContacts.Find("[Email3Address] = \"" + email + "\"") as Outlook.ContactItem;
                if (item != null)
                {
                    col.Add(item);
                    do
                    {
                        item = OutlookContacts.FindNext() as Outlook.ContactItem;
                        if (item != null)
                            col.Add(item);
                    } while (item != null);
                }
			}
			catch (Exception)
			{
				//TODO: should not get here.
			}

			return col;

		}

		public GroupEntry GetGoogleGroupById(string id)
		{
			return _googleGroups.FindById(new AtomId(id)) as GroupEntry;
		}

		public GroupEntry GetGoogleGroupByName(string name)
		{
			foreach (GroupEntry group in _googleGroups)
			{
				if (group.Title.Text == name)
					return group;
			}
			return null;
		}
		public GroupEntry CreateGroup(string name)
		{
			GroupEntry group = new GroupEntry();
			group.Title.Text = name;
			group.Dirty = true;
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

			string oCount = "Outlook Contact Count: " + _outlookContacts.Count;
			string gCount = "Google Contact Count: " + _googleContacts.Count;
			string mCount = "Matches Count: " + _matches.Count;

			MessageBox.Show(string.Format(msg, oCount, gCount, mCount), "DEBUG INFO", MessageBoxButtons.OK, MessageBoxIcon.Information);
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