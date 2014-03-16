using System;
using System.Collections.Generic;
using System.Text;
using Google.Contacts;
using Google.Documents;
using System.Reflection;
using Google.GData.Extensions;
using Google.GData.Calendar;

namespace GoContactSyncMod
{
    class ConflictResolver : IConflictResolver
    {
        private ConflictResolverForm _form;

        public ConflictResolver()
        {
            _form = new ConflictResolverForm();
        }
        

        #region IConflictResolver Members

        public ConflictResolution Resolve(ContactMatch match, bool isNewMatch)
        {
            string name = match.ToString();

            if (isNewMatch)
            {
                _form.messageLabel.Text =
                    "This is the first time these Outlook and Google Contacts \"" + name +
                    "\" are synced. Choose which you would like to keep.";
                _form.skip.Text = "Keep both";
            }
            else
            {
                _form.messageLabel.Text =
                    "Both the Outlook Contact and the Google Contact \"" + name +
                    "\" have been changed. Choose which you would like to keep.";
            }
            
            _form.OutlookItemTextBox.Text = string.Empty;
            _form.GoogleItemTextBox.Text = string.Empty;
            if (match.OutlookContact != null)
            {
                Microsoft.Office.Interop.Outlook.ContactItem item = match.OutlookContact.GetOriginalItemFromOutlook();
                try
                {
                    _form.OutlookItemTextBox.Text = ContactMatch.GetSummary(item);
                }
                finally
                {
                    if (item != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(item);
                        item = null;
                    }
                }

            }

            if (match.GoogleContact != null)
                _form.GoogleItemTextBox.Text = ContactMatch.GetSummary(match.GoogleContact);
           

            return Resolve();
        }

        public ConflictResolution ResolveDuplicate(OutlookContactInfo outlookContact, List<Contact> googleContacts, out Contact googleContact)
        {
            string name = ContactMatch.GetName(outlookContact);

            _form.messageLabel.Text =
                     "There are multiple Google Contacts (" + googleContacts.Count + ") matching unique properties for Outlook Contact \"" + name +
                     "\". Please choose from the combobox below the Google Contact you would like to match with Outlook and if you want to keep the Google or Outlook properties of the selected contact.";


            _form.OutlookItemTextBox.Text = string.Empty;
            _form.GoogleItemTextBox.Text = string.Empty;
            
            Microsoft.Office.Interop.Outlook.ContactItem item = outlookContact.GetOriginalItemFromOutlook();
            try
            {
                _form.OutlookItemTextBox.Text = ContactMatch.GetSummary(item);
            }
            finally
            {
                if (item != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(item);
                    item = null;
                }
            }

            

            _form.GoogleComboBox.DataSource = googleContacts;
            _form.GoogleComboBox.Visible = true;
            _form.AllCheckBox.Visible = false;
            _form.skip.Text = "Keep both";
            

            ConflictResolution res = Resolve();
            googleContact = _form.GoogleComboBox.SelectedItem as Contact;

            return res;

        }

        public DeleteResolution ResolveDelete(OutlookContactInfo outlookContact)
        {
            string name = ContactMatch.GetName(outlookContact);

            _form.Text = "Google Contact deleted";
            _form.messageLabel.Text =
                "Google Contact \"" + name +
                "\" doesn't exist anymore. Do you want to delete it also on Outlook side?";            

            _form.OutlookItemTextBox.Text = string.Empty;
            _form.GoogleItemTextBox.Text = string.Empty;
            Microsoft.Office.Interop.Outlook.ContactItem item = outlookContact.GetOriginalItemFromOutlook();
            try
            {
                _form.OutlookItemTextBox.Text = ContactMatch.GetSummary(item);
            }
            finally
            {
                if (item != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(item);
                    item = null;
                }
            }       
            
            _form.keepOutlook.Text = "Keep Outlook";
            _form.keepGoogle.Text = "Delete Outlook";
            _form.skip.Enabled = false;

            return ResolveDeletedGoogle();
        }

        public DeleteResolution ResolveDelete(Contact googleContact)
        {
            string name = ContactMatch.GetName(googleContact);

            _form.Text = "Outlook Contact deleted";
            _form.messageLabel.Text =
                "Outlook Contact \"" + name +
                "\" doesn't exist aynmore. Do you want to delete it also on Google side?";                       

            _form.OutlookItemTextBox.Text = string.Empty;
            _form.GoogleItemTextBox.Text = string.Empty;
            _form.GoogleItemTextBox.Text = ContactMatch.GetSummary(googleContact);

            _form.keepOutlook.Text = "Keep Google";
            _form.keepGoogle.Text = "Delete Google";
            _form.skip.Enabled = false;

            return ResolveDeletedOutlook();
        }

        

        private ConflictResolution Resolve()
        {

            switch (SettingsForm.Instance.ShowConflictDialog(_form))
            {
                case System.Windows.Forms.DialogResult.Ignore:
                    // skip
                    return _form.AllCheckBox.Checked ? ConflictResolution.SkipAlways : ConflictResolution.Skip;
                case System.Windows.Forms.DialogResult.No:
                    // google wins
                    return _form.AllCheckBox.Checked ? ConflictResolution.GoogleWinsAlways : ConflictResolution.GoogleWins;
                case System.Windows.Forms.DialogResult.Yes:
                    // outlook wins
                    return _form.AllCheckBox.Checked ? ConflictResolution.OutlookWinsAlways : ConflictResolution.OutlookWins;
                default:
                    return ConflictResolution.Cancel;
            }
        }

        private DeleteResolution ResolveDeletedOutlook()
        {

            switch (SettingsForm.Instance.ShowConflictDialog(_form))
            {              
                case System.Windows.Forms.DialogResult.No:
                    // google wins
                    return _form.AllCheckBox.Checked ? DeleteResolution.DeleteGoogleAlways : DeleteResolution.DeleteGoogle;
                case System.Windows.Forms.DialogResult.Yes:
                    // outlook wins
                    return _form.AllCheckBox.Checked ? DeleteResolution.KeepGoogleAlways : DeleteResolution.KeepGoogle;
                default:
                    return DeleteResolution.Cancel;
            }
        }

        private DeleteResolution ResolveDeletedGoogle()
        {

            switch (SettingsForm.Instance.ShowConflictDialog(_form))
            {               
                case System.Windows.Forms.DialogResult.No:
                    // google wins
                    return _form.AllCheckBox.Checked ? DeleteResolution.DeleteOutlookAlways : DeleteResolution.DeleteOutlook;
                case System.Windows.Forms.DialogResult.Yes:
                    // outlook wins
                    return _form.AllCheckBox.Checked ? DeleteResolution.KeepOutlookAlways : DeleteResolution.KeepOutlook;
                default:
                    return DeleteResolution.Cancel;
            }
        }

        //private string GetPropertyInfos(object myObject)
        //{
        //    // get all public static properties of OutlookItem
        //    PropertyInfo[] propertyInfos;
        //    propertyInfos = myObject.GetType().GetProperties();
        //    // sort properties by name
        //    Array.Sort(propertyInfos,
        //            delegate(PropertyInfo propertyInfo1, PropertyInfo propertyInfo2)
        //            { return propertyInfo1.Name.CompareTo(propertyInfo2.Name); });

        //    string ret = String.Empty;
        //    // write property names
        //    foreach (PropertyInfo propertyInfo in propertyInfos)
        //    {
        //        object value = propertyInfo.GetValue(myObject, null);

        //        if (value != null && !(value is string))
        //            ret += propertyInfo.Name + ": " + value + "\r\n";
        //    }

        //    return ret;
        //}

        public ConflictResolution Resolve(Microsoft.Office.Interop.Outlook.NoteItem outlookNote, Document googleNote, Syncronizer sync, bool isNewMatch)
        {
            string name = string.Empty;

            _form.OutlookItemTextBox.Text = string.Empty;
            _form.GoogleItemTextBox.Text = string.Empty;
            if (outlookNote != null)
            {
                name = outlookNote.Subject;
                _form.OutlookItemTextBox.Text = outlookNote.Body;
            }

            if (googleNote != null)
            {
                name = googleNote.Title;
                _form.GoogleItemTextBox.Text = NotePropertiesUtils.GetBody(sync, googleNote);
            }

            if (isNewMatch)
            {
                _form.messageLabel.Text =
                    "This is the first time these Outlook and Google notes \"" + name +
                    "\" are synced. Choose which you would like to keep.";
                _form.skip.Text = "Keep both";
            }
            else
            {
                _form.messageLabel.Text =
                "Both the Outlook Note and the Google Note \"" + name +
                "\" have been changed. Choose which you would like to keep.";
            }
            

           
            

            return Resolve();
        }
        public DeleteResolution ResolveDelete(Microsoft.Office.Interop.Outlook.NoteItem outlookNote)
        {            

            _form.Text = "Google note deleted";
            _form.messageLabel.Text =
                "Google note \"" + outlookNote.Subject +
                "\" doesn't exist aynmore. Do you want to delete it also on Outlook side?";

            _form.OutlookItemTextBox.Text = outlookNote.Body;
            _form.GoogleItemTextBox.Text = string.Empty;
            
            _form.keepOutlook.Text = "Keep Outlook";
            _form.keepGoogle.Text = "Delete Outlook";
            _form.skip.Enabled = false;

            return ResolveDeletedGoogle();
        }

        public DeleteResolution ResolveDelete(Document googleNote, Syncronizer sync)
        {

            _form.Text = "Outlook note deleted";
            _form.messageLabel.Text =
                "Outlook note \"" + googleNote.Title +
                "\" doesn't exist aynmore. Do you want to delete it also on Google side?";

            _form.OutlookItemTextBox.Text = string.Empty;
            _form.GoogleItemTextBox.Text = NotePropertiesUtils.GetBody(sync, googleNote);

            _form.keepOutlook.Text = "Keep Google";
            _form.keepGoogle.Text = "Delete Google";
            _form.skip.Enabled = false;

            return ResolveDeletedOutlook();
        }

        public ConflictResolution Resolve(Microsoft.Office.Interop.Outlook.AppointmentItem outlookAppointment, EventEntry googleAppointment, Syncronizer sync, bool isNewMatch)
        {
            string name = string.Empty;

            _form.OutlookItemTextBox.Text = string.Empty;
            _form.GoogleItemTextBox.Text = string.Empty;
            if (outlookAppointment != null)
            {
                name = outlookAppointment.Subject + " - " + outlookAppointment.Start;                
                _form.OutlookItemTextBox.Text += outlookAppointment.Body;
            }

            if (googleAppointment != null)
            {
                name = googleAppointment.Title.Text + " - " + Syncronizer.GetTime(googleAppointment);                
                _form.GoogleItemTextBox.Text += googleAppointment.Content.Content;
            }

            if (isNewMatch)
            {
                _form.messageLabel.Text =
                    "This is the first time these appointments \"" + name +
                    "\" are synced. Choose which you would like to keep.";
                _form.skip.Text = "Keep both";
            }
            else
            {
                _form.messageLabel.Text =
                "Both the Outlook and Google Appointment \"" + name +
                "\" have been changed. Choose which you would like to keep.";
            }




            return Resolve();
        }
        public DeleteResolution ResolveDelete(Microsoft.Office.Interop.Outlook.AppointmentItem outlookAppointment)
        {

            _form.Text = "Google appointment deleted";
            _form.messageLabel.Text =
                "Google appointment \"" + outlookAppointment.Subject + " - " + outlookAppointment.Start +
                "\" doesn't exist aynmore. Do you want to delete it also on Outlook side?";

            _form.GoogleItemTextBox.Text = String.Empty;            
            _form.OutlookItemTextBox.Text += outlookAppointment.Body;
            


            _form.keepOutlook.Text = "Keep Outlook";
            _form.keepGoogle.Text = "Delete Outlook";
            _form.skip.Enabled = false;

            return ResolveDeletedGoogle();
        }

        public DeleteResolution ResolveDelete(EventEntry googleAppointment)
        {

            _form.Text = "Outlook appointment deleted";
            _form.messageLabel.Text =
                "Outlook appointment \"" + googleAppointment.Title.Text + " - " + Syncronizer.GetTime(googleAppointment) +
                "\" doesn't exist aynmore. Do you want to delete it also on Google side?";

            _form.OutlookItemTextBox.Text = String.Empty;            
            _form.GoogleItemTextBox.Text += googleAppointment.Content.Content;


            _form.keepOutlook.Text = "Keep Google";
            _form.keepGoogle.Text = "Delete Google";
            _form.skip.Enabled = false;

            return ResolveDeletedOutlook();
        }

        #endregion
    }
}
