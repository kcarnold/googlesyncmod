using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Outlook = Microsoft.Office.Interop.Outlook;

using System.Runtime.InteropServices;
using Google.Apis.Calendar;
using Google.Apis.Calendar.v3.Data;

namespace GoContactSyncMod
{
    internal static class AppointmentsMatcher
    {
        /// <summary>
        /// Time tolerance in seconds - used when comparing date modified.
        /// Less than 60 seconds doesn't make sense, as the lastSync is saved without seconds and if it is compared
        /// with the LastUpdate dates of Google and Outlook, in the worst case you compare e.g. 15:59 with 16:00 and 
        /// after truncating to minutes you compare 15:00 wiht 16:00
        /// </summary>
        public static int TimeTolerance = 60;

        public delegate void NotificationHandler(string message);
        public static event NotificationHandler NotificationReceived;

        /// <summary>
        /// Matches outlook and Google appointment by a) id b) properties.
        /// </summary>
        /// <param name="sync">Syncronizer instance</param>
        /// <returns>Returns a list of match pairs (outlook appointment + Google appointment) for all appointment. Those that weren't matche will have it's peer set to null</returns>
        public static List<AppointmentMatch> MatchAppointments(Syncronizer sync)
        {
            Logger.Log("Matching Outlook and Google appointments...", EventType.Information);
            var result = new List<AppointmentMatch>();

            var googleAppointmentExceptions = new List<Event>();

            //for each outlook appointment try to get Google appointment id from user properties
            //if no match - try to match by properties
            //if no match - create a new match pair without Google appointment. 
            //foreach (Outlook._AppointmentItem olc in outlookAppointments)
            var OutlookAppointmentsWithoutSyncId = new Collection<Outlook.AppointmentItem>();
            #region Match first all outlookAppointments by sync id
            for (int i = 1; i <= sync.OutlookAppointments.Count; i++)
            {
                Outlook.AppointmentItem ola;               

                try
                {
                    ola = sync.OutlookAppointments[i] as Outlook.AppointmentItem;

                    //For debugging:
                    //if (ola.Subject == "Flex Pilot - Latest Status") // - 14.07.2014 11:30:00
                    //    throw new Exception("Debugging: Flex Pilot - Latest Status");

                    if (ola == null || string.IsNullOrEmpty(ola.Subject) && ola.Start == AppointmentSync.outlookDateMin)
                    {
                        Logger.Log("Empty Outlook appointment found. Skipping", EventType.Warning);
                        sync.SkippedCount++;
                        sync.SkippedCountNotMatches++;
                        continue;
                    }
                    else if (ola.MeetingStatus == Outlook.OlMeetingStatus.olMeetingCanceled || ola.MeetingStatus == Outlook.OlMeetingStatus.olMeetingReceivedAndCanceled)
                    {
                        Logger.Log("Skipping Outlook appointment found because it is cancelled: " + ola.Subject + " - " + ola.Start, EventType.Debug);
                        //sync.SkippedCount++;
                        //sync.SkippedCountNotMatches++;
                        continue;
                    }
                    else if (Syncronizer.MonthsInPast > 0 && 
                             (ola.IsRecurring && ola.GetRecurrencePattern().PatternEndDate < DateTime.Now.AddMonths(-Syncronizer.MonthsInPast) ||
                             !ola.IsRecurring && ola.End < DateTime.Now.AddMonths(-Syncronizer.MonthsInPast)) ||
                        Syncronizer.MonthsInFuture > 0 && 
                             (ola.IsRecurring && ola.GetRecurrencePattern().PatternStartDate > DateTime.Now.AddMonths(Syncronizer.MonthsInFuture) ||
                             !ola.IsRecurring && ola.Start > DateTime.Now.AddMonths(Syncronizer.MonthsInFuture)))
                    {
                        Logger.Log("Skipping Outlook appointment because it is out of months range to sync:" + ola.Subject + " - " + ola.Start, EventType.Debug);                       
                        continue;
                    }                   
                    //else if (ola.Subject != "Business Trip")
                    //{//Only for debugging purposes, please comment if not needed
                    //    Logger.Log("Skipping Outlook appointment because of DEBUGGING:" + ola.Subject + " - " + ola.Start, EventType.Error);
                    //    continue;
                    //}

                }
                catch (Exception ex)
                {
                    //this is needed because some appointments throw exceptions
                    Logger.Log("Accessing Outlook appointment threw and exception. Skipping: " + ex.Message, EventType.Warning);
                    sync.SkippedCount++;
                    sync.SkippedCountNotMatches++;
                    continue;
                }

                //try
                //{

                if (NotificationReceived != null)
                    NotificationReceived(String.Format("Matching appointment {0} of {1} by id: {2} ...", i, sync.OutlookAppointments.Count, ola.Subject));

                // Create our own info object to go into collections/lists, so we can free the Outlook objects and not run out of resources / exceed policy limits.
                //OutlookAppointmentInfo olci = new OutlookAppointmentInfo(ola, sync);

                //try to match this appointment to one of Google appointments               
                Outlook.ItemProperties userProperties = ola.ItemProperties;
                Outlook.ItemProperty idProp = userProperties[sync.OutlookPropertyNameId];
                try
                {
                    if (idProp != null)
                    {
                        string googleAppointmentId = string.Copy((string)idProp.Value);
                        Event foundAppointment = sync.GetGoogleAppointmentById(googleAppointmentId);
                        var match = new AppointmentMatch(ola, null);

                        

                        if (foundAppointment != null && !foundAppointment.Status.Equals("cancelled"))
                        {
                            //we found a match by google id, that is not deleted or cancelled yet

                            //ToDo: For an unknown reason, some appointments are duplicate in GoogleAppointments, therefore remove all duplicates before continuing
                            while (foundAppointment != null)
                            {
                                match.AddGoogleAppointment(foundAppointment);
                                result.Add(match);
                                //Remove the appointment from the list to not sync it twice
                                if (sync.GoogleAppointments.Contains(foundAppointment))
                                {
                                    sync.GoogleAppointments.Remove(foundAppointment);
                                    foundAppointment = sync.GetGoogleAppointmentById(googleAppointmentId);
                                }
                                else
                                    foundAppointment = null;                                
                            }


                        }
                        else
                        {
                        
                            OutlookAppointmentsWithoutSyncId.Add(ola);                        
                        }
                      
                    }
                    else
                        OutlookAppointmentsWithoutSyncId.Add(ola);
                }
                finally
                {
                    if (idProp != null)
                        Marshal.ReleaseComObject(idProp);
                    Marshal.ReleaseComObject(userProperties);
                }
                

            }
            #endregion
            #region Match the remaining appointments by properties

            for (int i = 0; i < OutlookAppointmentsWithoutSyncId.Count; i++)
            {
                Outlook.AppointmentItem ola = OutlookAppointmentsWithoutSyncId[i];

                if (NotificationReceived != null)
                    NotificationReceived(String.Format("Matching appointment {0} of {1} by unique properties: {2} ...", i + 1, OutlookAppointmentsWithoutSyncId.Count, ola.Subject));

                //no match found by id => match by subject/title
                //create a default match pair with just outlook appointment.
                var match = new AppointmentMatch(ola, null);

                //foreach Google appointment try to match and create a match pair if found some match(es)
                for (int j = sync.GoogleAppointments.Count - 1; j >= 0; j--)
                {
                    var googleAppointment = sync.GoogleAppointments[j];
                    // only match if there is a appointment targetBody, else
                    // a matching Google appointment will be created at each sync                
                    if (!googleAppointment.Status.Equals("cancelled") && ola.Subject == googleAppointment.Summary && googleAppointment.Start.DateTime != null && ola.Start == googleAppointment.Start.DateTime)                         
                    {                        
                        match.AddGoogleAppointment(googleAppointment);
                        sync.GoogleAppointments.Remove(googleAppointment);

                        //ToDo: For an unknown reason, some appointments are duplicate in GoogleAppointments, therefore remove all duplicates before continuing                                                
                        Event foundAppointment = sync.GetGoogleAppointmentById(AppointmentPropertiesUtils.GetGoogleId(googleAppointment));
                        while (foundAppointment != null)
                        {
                            //we found a match by google id, that is not deleted yet                       
                            match.AddGoogleAppointment(foundAppointment);
                            //Remove the appointment from the list to not sync it twice
                            if (sync.GoogleAppointments.Contains(foundAppointment))
                            {
                                sync.GoogleAppointments.Remove(foundAppointment);
                                foundAppointment = sync.GetGoogleAppointmentById(AppointmentPropertiesUtils.GetGoogleId(googleAppointment));
                            }
                            else
                                foundAppointment = null;
                        }
                    }

                }

                if (match.GoogleAppointment == null)
                    Logger.Log(string.Format("No match found for outlook appointment ({0}) => {1}", match.OutlookAppointment.Subject + " - " + match.OutlookAppointment.Start, (AppointmentPropertiesUtils.GetOutlookGoogleAppointmentId(sync, match.OutlookAppointment) != null ? "Delete from Outlook" : "Add to Google")), EventType.Information);

                result.Add(match);
            }
            #endregion

            
            //for each Google appointment that's left (they will be nonmatched) create a new match pair without outlook appointment. 
            for (int i = 0; i < sync.GoogleAppointments.Count; i++)
            {                
                var googleAppointment = sync.GoogleAppointments[i];

                //For debugging:
                //if (googleAppointment.Summary == "Flex Pilot - Latest Status")
                //    throw new Exception("Debugging: Flex Pilot - Latest Status");

               

                if (NotificationReceived != null)
                    NotificationReceived(String.Format("Adding new Google appointment {0} of {1} by unique properties: {2} ...", i + 1, sync.GoogleAppointments.Count, googleAppointment.Summary));

                
                if (googleAppointment.RecurringEventId != null)
                {
                    sync.SkippedCountNotMatches++;
                    googleAppointmentExceptions.Add(googleAppointment);
                }
                else if (googleAppointment.Status.Equals("cancelled"))
                {
                    Logger.Log("Skipping Google appointment found because it is cancelled: " + googleAppointment.Summary + " - " + Syncronizer.GetTime(googleAppointment), EventType.Debug);
                    //sync.SkippedCount++;
                    //sync.SkippedCountNotMatches++;
                }
                else if (string.IsNullOrEmpty(googleAppointment.Summary) && (googleAppointment.Start == null || googleAppointment.Start.DateTime == null && googleAppointment.Start.Date == null))
                {
                    // no title or time
                    sync.SkippedCount++;
                    sync.SkippedCountNotMatches++;
                    Logger.Log("Skipped GoogleAppointment because no unique property found (Subject or StartDate):" + googleAppointment.Summary + " - " + Syncronizer.GetTime(googleAppointment), EventType.Warning);
                }
                //else if (googleAppointment.Summary != "Business Trip")
                //{//Only for debugging purposes, please comment if not needed
                //    Logger.Log("Skipping Google appointment because of DEBUGGING:" + googleAppointment.Summary + " - " + googleAppointment.Times[0].StartTime, EventType.Error);
                //    continue;
                //}
                else
                {
                    Logger.Log(string.Format("No match found for Google appointment ({0}) => {1}", googleAppointment.Summary + " - " + Syncronizer.GetTime(googleAppointment), (!string.IsNullOrEmpty(AppointmentPropertiesUtils.GetGoogleOutlookAppointmentId(sync.SyncProfile, googleAppointment)) ? "Delete from Google" : "Add to Outlook")), EventType.Information);
                    var match = new AppointmentMatch(null, googleAppointment);
                    result.Add(match);
                }
            }

            //for each Google appointment exception, assign to proper match
            for (int i = 0; i < googleAppointmentExceptions.Count; i++)
            {
                var googleAppointment = googleAppointmentExceptions[i];
                if (NotificationReceived != null)
                    NotificationReceived(String.Format("Adding Google appointment exception {0} of {1} ...", i + 1, googleAppointmentExceptions.Count, googleAppointment.Summary + " - " + Syncronizer.GetTime(googleAppointment)));

                //Search for original recurrent event in matches
                //AtomId atomId = new AtomId(googleAppointment.Id.AbsoluteUri.Substring(0, googleAppointment.Id.AbsoluteUri.LastIndexOf("/") + 1) + googleAppointment.RecurringEventId);
                bool found = false;
                foreach (AppointmentMatch match in result)
                {
                    if (match.GoogleAppointment != null && googleAppointment.RecurringEventId.Equals(match.GoogleAppointment.Id))
                    {
                        match.GoogleAppointmentExceptions.Add(googleAppointment);
                        found = true;
                        break;
                    }
                }

                if (!found)
                    Logger.Log(string.Format("No match found for Google appointment exception: {0}", googleAppointment.Summary + " - " + Syncronizer.GetTime(googleAppointment)), EventType.Debug);
            }

            return result;
        }





        public static void SyncAppointments(Syncronizer sync)
        {
            for (int i = 0; i < sync.Appointments.Count; i++)
            {
                AppointmentMatch match = sync.Appointments[i];
                if (NotificationReceived != null)
                {
                    string name = string.Empty;
                    if (match.OutlookAppointment != null)
                        name = match.OutlookAppointment.Subject + " - " + match.OutlookAppointment.Start;
                    else if (match.GoogleAppointment != null)
                        name = match.GoogleAppointment.Summary + " - " + Syncronizer.GetTime(match.GoogleAppointment);
                    NotificationReceived(String.Format("Syncing appointment {0} of {1}: {2} ...", i + 1, sync.Appointments.Count, name));
                }

                SyncAppointment(match, sync);
            }
        }
        public static void SyncAppointment(AppointmentMatch match, Syncronizer sync)
        {

            if (match.GoogleAppointment == null && match.OutlookAppointment != null)
            {
                //no Google appointment                               
                string googleAppointmentId = AppointmentPropertiesUtils.GetOutlookGoogleAppointmentId(sync, match.OutlookAppointment);
                if (!string.IsNullOrEmpty(googleAppointmentId))
                {
                    //if (match.OutlookAppointment.IsRecurring && match.OutlookAppointment.RecurrenceState == Outlook.OlRecurrenceState.olApptMaster &&
                    //    (Syncronizer.MonthsInPast == 0 || new DateTime(DateTime.Now.AddMonths(-Syncronizer.MonthsInPast).Year, match.OutlookAppointment.End.Month, match.OutlookAppointment.End.Day) >= DateTime.Now.AddMonths(-Syncronizer.MonthsInPast)) &&
                    //    (Syncronizer.MonthsInFuture == 0 || new DateTime(DateTime.Now.AddMonths(-Syncronizer.MonthsInPast).Year, match.OutlookAppointment.Start.Month, match.OutlookAppointment.Start.Day) <= DateTime.Now.AddMonths(Syncronizer.MonthsInFuture))                        
                    //    ||
                    //    (Syncronizer.MonthsInPast == 0 || match.OutlookAppointment.End >= DateTime.Now.AddMonths(-Syncronizer.MonthsInPast)) &&
                    //    (Syncronizer.MonthsInFuture == 0 || match.OutlookAppointment.Start <= DateTime.Now.AddMonths(Syncronizer.MonthsInFuture))                        
                    //    )
                    //{

                        //Redundant check if exist, but in case an error occurred in MatchAppointments or not all appointments have been loaded (e.g. because months before/after constraint)
                        Event matchingGoogleAppointment = null;
                        if (sync.AllGoogleAppointments != null)
                            matchingGoogleAppointment = sync.GetGoogleAppointmentById(googleAppointmentId);
                        else
                            matchingGoogleAppointment = sync.LoadGoogleAppointments(googleAppointmentId, 0, 0, null, null);
                        if (matchingGoogleAppointment == null)
                        {
                            

                            if (!sync.PromptDelete && match.OutlookAppointment.Recipients.Count == 0)
                                sync.DeleteOutlookResolution = DeleteResolution.DeleteOutlookAlways;
                            else if (sync.DeleteOutlookResolution != DeleteResolution.DeleteOutlookAlways &&
                                     sync.DeleteOutlookResolution != DeleteResolution.KeepOutlookAlways)
                            {
                                var r = new ConflictResolver();
                                sync.DeleteOutlookResolution = r.ResolveDelete(match.OutlookAppointment);
                            }
                            switch (sync.DeleteOutlookResolution)
                            {
                                case DeleteResolution.KeepOutlook:
                                case DeleteResolution.KeepOutlookAlways:
                                    AppointmentPropertiesUtils.ResetOutlookGoogleAppointmentId(sync, match.OutlookAppointment);
                                    break;
                                case DeleteResolution.DeleteOutlook:
                                case DeleteResolution.DeleteOutlookAlways:
                                    
                                    if (match.OutlookAppointment.Recipients.Count > 1)
                                    {
                                        //ToDo:Maybe find as better way, e.g. to ask the user, if he wants to overwrite the invalid appointment                                
                                        Logger.Log("Outlook Appointment not deleted, because multiple participants found,  invitation maybe NOT sent by Google: " + match.OutlookAppointment.Subject + " - " + match.OutlookAppointment.Start, EventType.Information);
                                        AppointmentPropertiesUtils.ResetOutlookGoogleAppointmentId(sync, match.OutlookAppointment);
                                        break;
                                    }
                                    else
                                        //Avoid recreating a GoogleAppointment already existing
                                        //==> Delete this OutlookAppointment instead if previous match existed but no match exists anymore
                                        return;
                                default:
                                    throw new ApplicationException("Cancelled");
                            }
                        }
                        else
                        {
                            sync.SkippedCount++;
                            match.GoogleAppointment = matchingGoogleAppointment;
                            Logger.Log("Outlook Appointment not deleted, because still existing on Google side, maybe because months restriction: " + match.OutlookAppointment.Subject + " - " + match.OutlookAppointment.Start, EventType.Information);
                            return;
                        }

                    }
                //}

                if (sync.SyncOption == SyncOption.GoogleToOutlookOnly)
                {
                    sync.SkippedCount++;
                    Logger.Log(string.Format("Outlook appointment not added to Google, because of SyncOption " + sync.SyncOption.ToString() + ": {0}", match.OutlookAppointment.Subject), EventType.Information);
                    return;
                }

                //create a Google appointment from Outlook appointment
                match.GoogleAppointment = Factory.NewEvent();

                sync.UpdateAppointment(match.OutlookAppointment, ref match.GoogleAppointment);

            }
            else if (match.OutlookAppointment == null && match.GoogleAppointment != null)
            {
                //no Outlook appointment                               
                string outlookAppointmenttId = AppointmentPropertiesUtils.GetGoogleOutlookAppointmentId(sync.SyncProfile,match.GoogleAppointment);
                if (!string.IsNullOrEmpty(outlookAppointmenttId))
                {
                    if (!sync.PromptDelete)
                        sync.DeleteGoogleResolution = DeleteResolution.DeleteGoogleAlways;
                    else if (sync.DeleteGoogleResolution != DeleteResolution.DeleteGoogleAlways &&
                             sync.DeleteGoogleResolution != DeleteResolution.KeepGoogleAlways)
                    {
                        var r = new ConflictResolver();
                        sync.DeleteGoogleResolution = r.ResolveDelete(match.GoogleAppointment);
                    }
                    switch (sync.DeleteGoogleResolution)
                    {
                        case DeleteResolution.KeepGoogle:
                        case DeleteResolution.KeepGoogleAlways:
                            AppointmentPropertiesUtils.ResetGoogleOutlookAppointmentId(sync.SyncProfile,match.GoogleAppointment);
                            break;
                        case DeleteResolution.DeleteGoogle:
                        case DeleteResolution.DeleteGoogleAlways:
                            //Avoid recreating a OutlookAppointment already existing
                            //==> Delete this googleAppointment instead if previous match existed but no match exists anymore 
                            return;
                        default:
                            throw new ApplicationException("Cancelled");
                    }
                }


                if (sync.SyncOption == SyncOption.OutlookToGoogleOnly)
                {
                    sync.SkippedCount++;
                    Logger.Log(string.Format("Google appointment not added to Outlook, because of SyncOption " + sync.SyncOption.ToString() + ": {0}", match.GoogleAppointment.Summary), EventType.Information);
                    return;
                }

                //create a Outlook appointment from Google appointment
                match.OutlookAppointment = Syncronizer.CreateOutlookAppointmentItem(Syncronizer.SyncAppointmentsFolder);

                sync.UpdateAppointment(ref match.GoogleAppointment, match.OutlookAppointment, match.GoogleAppointmentExceptions);
            }
            else if (match.OutlookAppointment != null && match.GoogleAppointment != null)
            {
                //ToDo: Check how to overcome appointment recurrences, which need more than 60 seconds to update and therefore get updated again and again because of time tolerance 60 seconds violated again and again
                

                //merge appointment details                

                //determine if this appointment pair were syncronized
                //DateTime? lastUpdated = GetOutlookPropertyValueDateTime(match.OutlookAppointment, sync.OutlookPropertyNameUpdated);
                DateTime? lastSynced = AppointmentPropertiesUtils.GetOutlookLastSync(sync, match.OutlookAppointment);
                if (lastSynced.HasValue)
                {
                    //appointment pair was syncronysed before.

                    //determine if Google appointment was updated since last sync

                    //lastSynced is stored without seconds. take that into account.
                    DateTime lastUpdatedOutlook = match.OutlookAppointment.LastModificationTime.AddSeconds(-match.OutlookAppointment.LastModificationTime.Second);
                    DateTime lastUpdatedGoogle = match.GoogleAppointment.Updated.Value.AddSeconds(-match.GoogleAppointment.Updated.Value.Second);
                    //consider GoogleAppointmentExceptions, because if they are updated, the master appointment doesn't have a new Saved TimeStamp
                    foreach (Event googleAppointment in match.GoogleAppointmentExceptions)
                    {
                        if (googleAppointment.Updated != null)//happens for cancelled events
                        {
                            DateTime lastUpdatedGoogleException = googleAppointment.Updated.Value.AddSeconds(-googleAppointment.Updated.Value.Second);
                            if (lastUpdatedGoogleException > lastUpdatedGoogle)
                                lastUpdatedGoogle = lastUpdatedGoogleException;
                        }
                    }

                    //check if both outlok and Google appointments where updated sync last sync
                    if ((int)lastUpdatedOutlook.Subtract(lastSynced.Value).TotalSeconds > TimeTolerance
                        && (int)lastUpdatedGoogle.Subtract(lastSynced.Value).TotalSeconds > TimeTolerance)
                    {
                        //both appointments were updated.
                        //options: 1) ignore 2) loose one based on SyncOption
                        //throw new Exception("Both appointments were updated!");

                        switch (sync.SyncOption)
                        {
                            case SyncOption.MergeOutlookWins:
                            case SyncOption.OutlookToGoogleOnly:
                                //overwrite Google appointment
                                Logger.Log("Outlook and Google appointment have been updated, Outlook appointment is overwriting Google because of SyncOption " + sync.SyncOption + ": " + match.OutlookAppointment.Subject + ".", EventType.Information);
                                sync.UpdateAppointment(match.OutlookAppointment, ref match.GoogleAppointment);
                                break;
                            case SyncOption.MergeGoogleWins:
                            case SyncOption.GoogleToOutlookOnly:
                                //overwrite outlook appointment
                                Logger.Log("Outlook and Google appointment have been updated, Google appointment is overwriting Outlook because of SyncOption " + sync.SyncOption + ": " + match.GoogleAppointment.Summary + ".", EventType.Information);
                                sync.UpdateAppointment(ref match.GoogleAppointment, match.OutlookAppointment, match.GoogleAppointmentExceptions);
                                break;
                            case SyncOption.MergePrompt:
                                //promp for sync option
                                if (sync.ConflictResolution != ConflictResolution.GoogleWinsAlways &&
                                    sync.ConflictResolution != ConflictResolution.OutlookWinsAlways &&
                                    sync.ConflictResolution != ConflictResolution.SkipAlways)
                                {
                                    var r = new ConflictResolver();
                                    sync.ConflictResolution = r.Resolve(match.OutlookAppointment, match.GoogleAppointment, sync, false);
                                }
                                switch (sync.ConflictResolution)
                                {
                                    case ConflictResolution.Skip:
                                    case ConflictResolution.SkipAlways:
                                        Logger.Log(string.Format("User skipped appointment ({0}).", match.ToString()), EventType.Information);
                                        sync.SkippedCount++;
                                        break;
                                    case ConflictResolution.OutlookWins:
                                    case ConflictResolution.OutlookWinsAlways:
                                        sync.UpdateAppointment(match.OutlookAppointment, ref match.GoogleAppointment);
                                        break;
                                    case ConflictResolution.GoogleWins:
                                    case ConflictResolution.GoogleWinsAlways:
                                        sync.UpdateAppointment(ref match.GoogleAppointment, match.OutlookAppointment, match.GoogleAppointmentExceptions);
                                        break;
                                    default:
                                        throw new ApplicationException("Cancelled");
                                }
                                break;
                        }
                        return;
                    }


                    //check if Outlook appointment was updated (with X second tolerance)
                    if (sync.SyncOption != SyncOption.GoogleToOutlookOnly &&
                        ((int)lastUpdatedOutlook.Subtract(lastSynced.Value).TotalSeconds > TimeTolerance ||
                         (int)lastUpdatedGoogle.Subtract(lastSynced.Value).TotalSeconds > TimeTolerance &&
                         sync.SyncOption == SyncOption.OutlookToGoogleOnly
                        )
                       )
                    {
                        //Outlook appointment was changed or changed Google appointment will be overwritten

                        if ((int)lastUpdatedGoogle.Subtract(lastSynced.Value).TotalSeconds > TimeTolerance &&
                            sync.SyncOption == SyncOption.OutlookToGoogleOnly)
                            Logger.Log("Google appointment has been updated since last sync, but Outlook appointment is overwriting Google because of SyncOption " + sync.SyncOption + ": " + match.OutlookAppointment.Subject + ".", EventType.Information);

                        sync.UpdateAppointment(match.OutlookAppointment, ref match.GoogleAppointment);

                        //at the moment use Outlook as "master" source of appointments - in the event of a conflict Google appointment will be overwritten.
                        //TODO: control conflict resolution by SyncOption
                        return;
                    }

                    //check if Google appointment was updated (with X second tolerance)
                    if (sync.SyncOption != SyncOption.OutlookToGoogleOnly &&
                        ((int)lastUpdatedGoogle.Subtract(lastSynced.Value).TotalSeconds > TimeTolerance ||
                         (int)lastUpdatedOutlook.Subtract(lastSynced.Value).TotalSeconds > TimeTolerance &&
                         sync.SyncOption == SyncOption.GoogleToOutlookOnly
                        )
                       )
                    {
                        //google appointment was changed or changed Outlook appointment will be overwritten

                        if ((int)lastUpdatedOutlook.Subtract(lastSynced.Value).TotalSeconds > TimeTolerance &&
                            sync.SyncOption == SyncOption.GoogleToOutlookOnly)
                            Logger.Log("Outlook appointment has been updated since last sync, but Google appointment is overwriting Outlook because of SyncOption " + sync.SyncOption + ": " + match.OutlookAppointment.Subject + ".", EventType.Information);

                        sync.UpdateAppointment(ref match.GoogleAppointment, match.OutlookAppointment, match.GoogleAppointmentExceptions);
                    }
                }
                else
                {
                    //appointments were never synced.
                    //merge appointments.
                    switch (sync.SyncOption)
                    {
                        case SyncOption.MergeOutlookWins:
                        case SyncOption.OutlookToGoogleOnly:
                            //overwrite Google appointment
                            sync.UpdateAppointment(match.OutlookAppointment, ref match.GoogleAppointment);
                            break;
                        case SyncOption.MergeGoogleWins:
                        case SyncOption.GoogleToOutlookOnly:
                            //overwrite outlook appointment
                            sync.UpdateAppointment(ref match.GoogleAppointment, match.OutlookAppointment, match.GoogleAppointmentExceptions);
                            break;
                        case SyncOption.MergePrompt:
                            //promp for sync option
                            if (sync.ConflictResolution != ConflictResolution.GoogleWinsAlways &&
                                sync.ConflictResolution != ConflictResolution.OutlookWinsAlways &&
                                    sync.ConflictResolution != ConflictResolution.SkipAlways)
                            {
                                var r = new ConflictResolver();
                                sync.ConflictResolution = r.Resolve(match.OutlookAppointment, match.GoogleAppointment, sync, true);
                            }
                            switch (sync.ConflictResolution)
                            {
                                case ConflictResolution.Skip:
                                case ConflictResolution.SkipAlways: //Keep both, Google AND Outlook
                                    sync.Appointments.Add(new AppointmentMatch(match.OutlookAppointment, null));
                                    sync.Appointments.Add(new AppointmentMatch(null, match.GoogleAppointment));
                                    break;
                                case ConflictResolution.OutlookWins:
                                case ConflictResolution.OutlookWinsAlways:
                                    sync.UpdateAppointment(match.OutlookAppointment, ref match.GoogleAppointment);
                                    break;
                                case ConflictResolution.GoogleWins:
                                case ConflictResolution.GoogleWinsAlways:
                                    sync.UpdateAppointment(ref match.GoogleAppointment, match.OutlookAppointment, match.GoogleAppointmentExceptions);
                                    break;
                                default:
                                    throw new ApplicationException("Canceled");
                            }
                            break;
                    }
                }

            }
            else
                throw new ArgumentNullException("AppointmenttMatch has all peers null.");
            //}
            //finally
            //{
            //if (outlookAppointment != null &&
            //    match.OutlookAppointment != null)
            //{
            //    match.OutlookAppointment.Update(outlookAppointment, sync);
            //    Marshal.ReleaseComObject(outlookAppointment);
            //    outlookAppointment = null;
            //}
            //}

        }        
    }



    internal class AppointmentMatch
    {
        //ToDo: OutlookappointmentInfo
        public Outlook.AppointmentItem OutlookAppointment;
        public Event GoogleAppointment;
        public readonly List<Event> AllGoogleAppointmentMatches = new List<Event>(1);
        public Event LastGoogleAppointment;
        public List<Event> GoogleAppointmentExceptions = new List<Event>();

        public AppointmentMatch(Outlook.AppointmentItem outlookAppointment, Event googleAppointment)
        {
            OutlookAppointment = outlookAppointment;
            GoogleAppointment = googleAppointment;
        }

        public void AddGoogleAppointment(Event googleAppointment)
        {
            if (googleAppointment == null)
                return;
            //throw new ArgumentNullException("googleAppointment must not be null.");

            if (GoogleAppointment == null)
                GoogleAppointment = googleAppointment;

            //this to avoid searching the entire collection. 
            //if last appointment it what we are trying to add the we have already added it earlier
            if (LastGoogleAppointment == googleAppointment)
                return;

            if (!AllGoogleAppointmentMatches.Contains(googleAppointment))
                AllGoogleAppointmentMatches.Add(googleAppointment);

            LastGoogleAppointment = googleAppointment;
        }

    }



}
