using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Outlook = Microsoft.Office.Interop.Outlook;

using System.Runtime.InteropServices;
using Google.GData.Calendar;

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
        /// Matches outlook and target appointment by a) target id b) properties.
        /// </summary>
        /// <param name="sync">Syncronizer instance</param>
        /// <returns>Returns a list of match pairs (outlook appointment + target appointment) for all appointment. Those that weren't matche will have it's peer set to null</returns>
        public static List<AppointmentMatch> MatchAppointments(Syncronizer sync)
        {
            Logger.Log("Matching Source and Target appointments...", EventType.Information);
            var result = new List<AppointmentMatch>();

            //string duplicateTargetMatches = "";
            //string duplicateOutlookAppointments = "";
            //sync.GoogleAppointmentDuplicates = new Collection<AppointmentMatch>();
            //sync.OutlookAppointmentDuplicates = new Collection<AppointmentMatch>();

            //for each outlook appointment try to get target appointment id from user properties
            //if no match - try to match by properties
            //if no match - create a new match pair without target appointment. 
            //foreach (Outlook._AppointmentItem olc in outlookAppointments)
            var OutlookAppointmentsWithoutSyncId = new Collection<Outlook.AppointmentItem>();
            #region Match first all outlookAppointments by sync id
            for (int i = 1; i <= sync.OutlookAppointments.Count; i++)
            {
                Outlook.AppointmentItem oln;
                try
                {
                    oln = sync.OutlookAppointments[i] as Outlook.AppointmentItem;
                    if (oln == null || string.IsNullOrEmpty(oln.Subject))
                    {
                        Logger.Log("Empty Outlook appointment found. Skipping", EventType.Warning);
                        sync.SkippedCount++;
                        sync.SkippedCountNotMatches++;
                        continue;
                    }
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
                    NotificationReceived(String.Format("Matching appointment {0} of {1} by id: {2} ...", i, sync.OutlookAppointments.Count, oln.Subject));

                // Create our own info object to go into collections/lists, so we can free the Outlook objects and not run out of resources / exceed policy limits.
                //OutlookAppointmentInfo olci = new OutlookAppointmentInfo(ola, sync);

                //try to match this appointment to one of target appointments
                var propertyName = string.Format(Syncronizer.OutlookUserPropertyTemplate, Syncronizer.GoogleAppointmentsFolder.GetHashCode().ToString(), Syncronizer.OutlookUserPropertyTemplateId);
                Outlook.ItemProperties userProperties = oln.ItemProperties;
                Outlook.ItemProperty idProp = userProperties[propertyName];
                try
                {
                    if (idProp != null)
                    {
                        string googleAppointmentId = string.Copy((string)idProp.Value);
                        EventEntry foundAppointment = sync.GetGoogleAppointmentById(googleAppointmentId);
                        var match = new AppointmentMatch(oln, null);

                        //Check first, that this is not a duplicate 
                        //e.g. by copying an existing Outlook appointment
                        //or by Outlook checked this as duplicate, but the user selected "Add new"
                        //    Collection<OutlookAppointmentInfo> duplicates = sync.OutlookAppointmentByProperty(sync.OutlookPropertyNameId, googleAppointmentId);
                        //    if (duplicates.Count > 1)
                        //    {
                        //        foreach (OutlookAppointmentInfo duplicate in duplicates)
                        //        {
                        //            if (!string.IsNullOrEmpty(googleAppointmentId))
                        //            {
                        //                Logger.Log("Duplicate Outlook appointment found, resetting match and try to match again: " + duplicate.FileAs, EventType.Warning);
                        //                idProp.Value = "";
                        //            }
                        //        }

                        //        if (foundAppointment != null && !foundAppointment.Deleted)
                        //        {
                        //            AppointmentPropertiesUtils.ResetGoogleOutlookAppointmentId(sync.SyncProfile, foundAppointment);
                        //        }

                        //        OutlookAppointmentsWithoutSyncId.Add(olci);
                        //    }
                        //    else
                        //    {

                        if (foundAppointment != null)
                        {
                            //we found a match by target id, that is not deleted yet
                            match.AddGoogleAppointment(foundAppointment);
                            result.Add(match);
                            //Remove the appointment from the list to not sync it twice
                            sync.GoogleAppointments.Remove(foundAppointment);
                        }
                        else
                        {
                            ////If no match found, is the appointment either deleted on Google side or was a copy on Outlook side 
                            ////If it is a copy on Outlook side, the idProp.Value must be emptied to assure, the appointment is created on Google side and not deleted on Outlook side
                            ////bool matchIsDuplicate = false;
                            //foreach (AppointmentMatch existingMatch in result)
                            //{
                            //    if (existingMatch.OutlookAppointment.UserProperties[sync.OutlookPropertyNameId].Value.Equals(idProp.Value))
                            //    {
                            //        //matchIsDuplicate = true;
                            //        idProp.Value = "";
                            //        break;
                            //    }

                            //}
                            OutlookAppointmentsWithoutSyncId.Add(oln);

                            //if (!matchIsDuplicate)
                            //    result.Add(match);
                        }
                        //    }
                    }
                    else
                        OutlookAppointmentsWithoutSyncId.Add(oln);
                }
                finally
                {
                    if (idProp != null)
                        Marshal.ReleaseComObject(idProp);
                    Marshal.ReleaseComObject(userProperties);
                }
                //}

                //finally
                //{
                //    Marshal.ReleaseComObject(ola);
                //    ola = null;
                //}

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

                //foreach target appointment try to match and create a match pair if found some match(es)
                for (int j = sync.GoogleAppointments.Count - 1; j >= 0; j--)
                {
                    var googleAppointment = sync.GoogleAppointments[j];
                    // only match if there is a appointment targetBody, else
                    // a matching target appointment will be created at each sync                
                    if (ola.Subject == googleAppointment.Title && ola.Start == googleAppointment.Start)
                    {
                        match.AddGoogleAppointment(googleAppointment);
                        sync.GoogleAppointments.Remove(googleAppointment);
                    }

                }

                #region find duplicates not needed now
                //if (match.GoogleAppointment == null && match.OutlookAppointment != null)
                //{//If GoogleAppointment, we have to expect a conflict because of Google insert of duplicates
                //    foreach (Appointment googleAppointment in sync.GoogleAppointments)
                //    {                        
                //        if (!string.IsNullOrEmpty(olc.FullName) && olc.FullName.Equals(googleAppointment.Title, StringComparison.InvariantCultureIgnoreCase) ||
                //         !string.IsNullOrEmpty(olc.FileAs) && olc.FileAs.Equals(googleAppointment.Title, StringComparison.InvariantCultureIgnoreCase) ||
                //         !string.IsNullOrEmpty(olc.Email1Address) && FindEmail(olc.Email1Address, googleAppointment.Emails) != null ||
                //         !string.IsNullOrEmpty(olc.Email2Address) && FindEmail(olc.Email1Address, googleAppointment.Emails) != null ||
                //         !string.IsNullOrEmpty(olc.Email3Address) && FindEmail(olc.Email1Address, googleAppointment.Emails) != null ||
                //         olc.MobileTelephoneNumber != null && FindPhone(olc.MobileTelephoneNumber, googleAppointment.Phonenumbers) != null
                //         )
                //    }
                //// check for each email 1,2 and 3 if a duplicate exists with same email, because Google doesn't like inserting new appointments with same email
                //Collection<Outlook.AppointmentItem> duplicates1 = new Collection<Outlook.AppointmentItem>();
                //Collection<Outlook.AppointmentItem> duplicates2 = new Collection<Outlook.AppointmentItem>();
                //Collection<Outlook.AppointmentItem> duplicates3 = new Collection<Outlook.AppointmentItem>();
                //if (!string.IsNullOrEmpty(olc.Email1Address))
                //    duplicates1 = sync.OutlookAppointmentByEmail(olc.Email1Address);

                //if (!string.IsNullOrEmpty(olc.Email2Address))
                //    duplicates2 = sync.OutlookAppointmentByEmail(olc.Email2Address);

                //if (!string.IsNullOrEmpty(olc.Email3Address))
                //    duplicates3 = sync.OutlookAppointmentByEmail(olc.Email3Address);


                //if (duplicates1.Count > 1 || duplicates2.Count > 1 || duplicates3.Count > 1)
                //{
                //    if (string.IsNullOrEmpty(duplicatesEmailList))
                //        duplicatesEmailList = "Outlook appointments with the same email have been found and cannot be synchronized. Please delete duplicates of:";

                //    if (duplicates1.Count > 1)
                //        foreach (Outlook.AppointmentItem duplicate in duplicates1)
                //        {
                //            string str = olc.FileAs + " (" + olc.Email1Address + ")";
                //            if (!duplicatesEmailList.Contains(str))
                //                duplicatesEmailList += Environment.NewLine + str;
                //        }
                //    if (duplicates2.Count > 1)
                //        foreach (Outlook.AppointmentItem duplicate in duplicates2)
                //        {
                //            string str = olc.FileAs + " (" + olc.Email2Address + ")";
                //            if (!duplicatesEmailList.Contains(str))
                //                duplicatesEmailList += Environment.NewLine + str;
                //        }
                //    if (duplicates3.Count > 1)
                //        foreach (Outlook.AppointmentItem duplicate in duplicates3)
                //        {
                //            string str = olc.FileAs + " (" + olc.Email3Address + ")";
                //            if (!duplicatesEmailList.Contains(str))
                //                duplicatesEmailList += Environment.NewLine + str;
                //        }
                //    continue;
                //}
                //else if (!string.IsNullOrEmpty(olc.Email1Address))
                //{
                //    AppointmentMatch dup = result.Find(delegate(AppointmentMatch match)
                //    {
                //        return match.OutlookAppointment != null && match.OutlookAppointment.Email1Address == olc.Email1Address;
                //    });
                //    if (dup != null)
                //    {
                //        Logger.Log(string.Format("Duplicate appointment found by Email1Address ({0}). Skipping", olc.FileAs), EventType.Information);
                //        continue;
                //    }
                //}

                //// check for unique mobile phone, because this sync tool uses the also the mobile phone to identify matches between Google and Outlook
                //Collection<Outlook.AppointmentItem> duplicatesMobile = new Collection<Outlook.AppointmentItem>();
                //if (!string.IsNullOrEmpty(olc.MobileTelephoneNumber))
                //    duplicatesMobile = sync.OutlookAppointmentByProperty("MobileTelephoneNumber", olc.MobileTelephoneNumber);

                //if (duplicatesMobile.Count > 1)
                //{
                //    if (string.IsNullOrEmpty(duplicatesMobileList))
                //        duplicatesMobileList = "Outlook appointments with the same mobile phone have been found and cannot be synchronized. Please delete duplicates of:";

                //    foreach (Outlook.AppointmentItem duplicate in duplicatesMobile)
                //    {
                //        sync.OutlookAppointmentDuplicates.Add(olc);
                //        string str = olc.FileAs + " (" + olc.MobileTelephoneNumber + ")";
                //        if (!duplicatesMobileList.Contains(str))
                //            duplicatesMobileList += Environment.NewLine + str;
                //    }
                //    continue;
                //}
                //else if (!string.IsNullOrEmpty(olc.MobileTelephoneNumber))
                //{
                //    AppointmentMatch dup = result.Find(delegate(AppointmentMatch match)
                //    {
                //        return match.OutlookAppointment != null && match.OutlookAppointment.MobileTelephoneNumber == olc.MobileTelephoneNumber;
                //    });
                //    if (dup != null)
                //    {
                //        Logger.Log(string.Format("Duplicate appointment found by MobileTelephoneNumber ({0}). Skipping", olc.FileAs), EventType.Information);
                //        continue;
                //    }
                //}

                #endregion

                //    if (match.AllGoogleAppointmentMatches == null || match.AllGoogleAppointmentMatches.Count == 0)
                //    {
                //        //Check, if this Outlook appointment has a match in the target duplicates
                //        bool duplicateFound = false;
                //        foreach (AppointmentMatch duplicate in sync.GoogleAppointmentDuplicates)
                //        {
                //            if (duplicate.AllGoogleAppointmentMatches.Count > 0 &&
                //                (!string.IsNullOrEmpty(olci.FileAs) && !string.IsNullOrEmpty(duplicate.AllGoogleAppointmentMatches[0].Title) && olci.FileAs.Equals(duplicate.AllGoogleAppointmentMatches[0].Title, StringComparison.InvariantCultureIgnoreCase) ||  //Replace twice to not replace a \r\n by \r\r\n. This is necessary because \r\n are saved as \n only to google
                //                 !string.IsNullOrEmpty(olci.FileAs) && !string.IsNullOrEmpty(duplicate.AllGoogleAppointmentMatches[0].Name.FullName) && olci.FileAs.Equals(duplicate.AllGoogleAppointmentMatches[0].Name.FullName, StringComparison.InvariantCultureIgnoreCase) ||
                //                 !string.IsNullOrEmpty(olci.FullName) && !string.IsNullOrEmpty(duplicate.AllGoogleAppointmentMatches[0].Name.FullName) && olci.FullName.Equals(duplicate.AllGoogleAppointmentMatches[0].Name.FullName, StringComparison.InvariantCultureIgnoreCase) ||
                //                 !string.IsNullOrEmpty(olci.Email1Address) && duplicate.AllGoogleAppointmentMatches[0].Emails.Count > 0 && olci.Email1Address.Equals(duplicate.AllGoogleAppointmentMatches[0].Emails[0].Address, StringComparison.InvariantCultureIgnoreCase) ||
                //                //!string.IsNullOrEmpty(olci.Email2Address) && FindEmail(olci.Email2Address, duplicate.AllGoogleAppointmentMatches[0].Emails) != null ||
                //                //!string.IsNullOrEmpty(olci.Email3Address) && FindEmail(olci.Email3Address, duplicate.AllGoogleAppointmentMatches[0].Emails) != null ||
                //                 olci.MobileTelephoneNumber != null && FindPhone(olci.MobileTelephoneNumber, duplicate.AllGoogleAppointmentMatches[0].Phonenumbers) != null ||
                //                 !string.IsNullOrEmpty(olci.FileAs) && string.IsNullOrEmpty(duplicate.AllGoogleAppointmentMatches[0].Title) && duplicate.AllGoogleAppointmentMatches[0].Organizations.Count > 0 && olci.FileAs.Equals(duplicate.AllGoogleAppointmentMatches[0].Organizations[0].Name, StringComparison.InvariantCultureIgnoreCase)
                //                ) ||
                //                !string.IsNullOrEmpty(olci.FileAs) && olci.FileAs.Equals(duplicate.OutlookAppointment.Subject, StringComparison.InvariantCultureIgnoreCase) ||
                //                !string.IsNullOrEmpty(olci.FullName) && olci.FullName.Equals(duplicate.OutlookAppointment.FullName, StringComparison.InvariantCultureIgnoreCase) ||
                //                !string.IsNullOrEmpty(olci.Email1Address) && olci.Email1Address.Equals(duplicate.OutlookAppointment.Email1Address, StringComparison.InvariantCultureIgnoreCase) ||
                //                //                                              olci.Email1Address.Equals(duplicate.OutlookAppointment.Email2Address, StringComparison.InvariantCultureIgnoreCase) ||
                //                //                                              olci.Email1Address.Equals(duplicate.OutlookAppointment.Email3Address, StringComparison.InvariantCultureIgnoreCase)
                //                //                                              ) ||
                //                //!string.IsNullOrEmpty(olci.Email2Address) && (olci.Email2Address.Equals(duplicate.OutlookAppointment.Email1Address, StringComparison.InvariantCultureIgnoreCase) ||
                //                //                                              olci.Email2Address.Equals(duplicate.OutlookAppointment.Email2Address, StringComparison.InvariantCultureIgnoreCase) ||
                //                //                                              olci.Email2Address.Equals(duplicate.OutlookAppointment.Email3Address, StringComparison.InvariantCultureIgnoreCase)
                //                //                                              ) ||
                //                //!string.IsNullOrEmpty(olci.Email3Address) && (olci.Email3Address.Equals(duplicate.OutlookAppointment.Email1Address, StringComparison.InvariantCultureIgnoreCase) ||
                //                //                                              olci.Email3Address.Equals(duplicate.OutlookAppointment.Email2Address, StringComparison.InvariantCultureIgnoreCase) ||
                //                //                                              olci.Email3Address.Equals(duplicate.OutlookAppointment.Email3Address, StringComparison.InvariantCultureIgnoreCase)
                //                //                                              ) ||
                //                olci.MobileTelephoneNumber != null && olci.MobileTelephoneNumber.Equals(duplicate.OutlookAppointment.MobileTelephoneNumber) ||
                //                !string.IsNullOrEmpty(olci.FileAs) && string.IsNullOrEmpty(duplicate.GoogleAppointment.Title) && duplicate.GoogleAppointment.Organizations.Count > 0 && olci.FileAs.Equals(duplicate.GoogleAppointment.Organizations[0].Name, StringComparison.InvariantCultureIgnoreCase)
                //               )
                //            {
                //                duplicateFound = true;
                //                sync.OutlookAppointmentDuplicates.Add(match);
                //                if (string.IsNullOrEmpty(duplicateOutlookAppointments))
                //                    duplicateOutlookAppointments = "Outlook appointment found that has been already identified as duplicate Google appointment (either same email, Mobile or FullName) and cannot be synchronized. Please delete or resolve duplicates of:";

                //                string str = olci.FileAs + " (" + olci.Email1Address + ", " + olci.MobileTelephoneNumber + ")";
                //                if (!duplicateOutlookAppointments.Contains(str))
                //                    duplicateOutlookAppointments += Environment.NewLine + str;
                //            }
                //        }

                //        if (!duplicateFound)
                if (match.GoogleAppointment == null)
                    Logger.Log(string.Format("No match found for outlook appointment ({0}) => {1}", match.OutlookAppointment.Subject + " - " + match.OutlookAppointment.Start, (AppointmentPropertiesUtils.GetOutlookGoogleAppointmentId(sync, match.OutlookAppointment) != null ? "Delete from Source" : "Add to Target")), EventType.Information);

                //    }
                //    else
                //    {
                //        //Remember Google duplicates to later react to it when resetting matches or syncing
                //        //ResetMatches: Also reset the duplicates
                //        //Sync: Skip duplicates (don't sync duplicates to be fail safe)
                //        if (match.AllGoogleAppointmentMatches.Count > 1)
                //        {
                //            sync.GoogleAppointmentDuplicates.Add(match);
                //            foreach (Appointment googleAppointment in match.AllGoogleAppointmentMatches)
                //            {
                //                //Create message for duplicatesFound exception
                //                if (string.IsNullOrEmpty(duplicateTargetMatches))
                //                    duplicateTargetMatches = "Outlook appointments matching with multiple Google appointments have been found (either same email, Mobile, FullName or company) and cannot be synchronized. Please delete or resolve duplicates of:";

                //                string str = olci.FileAs + " (" + olci.Email1Address + ", " + olci.MobileTelephoneNumber + ")";
                //                if (!duplicateTargetMatches.Contains(str))
                //                    duplicateTargetMatches += Environment.NewLine + str;
                //            }
                //        }



                //    }                

                result.Add(match);
            }
            #endregion

            //if (!string.IsNullOrEmpty(duplicateTargetMatches) || !string.IsNullOrEmpty(duplicateOutlookAppointments))
            //    duplicatesFound = new DuplicateDataException(duplicateTargetMatches + Environment.NewLine + Environment.NewLine + duplicateOutlookAppointments);
            //else
            //    duplicatesFound = null;

            //return result;

            //for each target appointment that's left (they will be nonmatched) create a new match pair without outlook appointment. 
            for (int i = 0; i < sync.GoogleAppointments.Count; i++)
            {
                var googleAppointment = sync.GoogleAppointments[i] as Outlook.AppointmentItem;
                if (NotificationReceived != null)
                    NotificationReceived(String.Format("Adding new Google appointment {0} of {1} by unique properties: {2} ...", i + 1, sync.GoogleAppointments.Count, googleAppointment.Subject));

                //string syncId = AppointmentPropertiesUtils.GetGoogleOutlookAppointmentId(sync.SyncProfile, googleAppointment);
                //if (!String.IsNullOrEmpty(syncId) && skippedOutlookIds.Contains(syncId))
                //{
                //    Logger.Log("Skipped GoogleAppointment because Outlook appointment couldn't be matched beacause of previous problem (see log): " + googleAppointment.Title, EventType.Warning);
                //}
                //else 
                if (string.IsNullOrEmpty(googleAppointment.Subject) && googleAppointment.Start == default(DateTime))
                {
                    // no title or content
                    sync.SkippedCount++;
                    sync.SkippedCountNotMatches++;
                    Logger.Log("Skipped GoogleAppointment because no unique property found (Subject or StartDate):" + googleAppointment.Subject, EventType.Warning);
                }
                else
                {
                    Logger.Log(string.Format("No match found for target appointment ({0}) => {1}", googleAppointment.Subject + " - "+googleAppointment.Start, (!string.IsNullOrEmpty(AppointmentPropertiesUtils.GetGoogleOutlookAppointmentId(sync.SyncProfile, googleAppointment)) ? "Delete from Target" : "Add to Source")), EventType.Information);
                    var match = new AppointmentMatch(null, googleAppointment);
                    result.Add(match);
                }
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
                        name = match.OutlookAppointment.Subject;
                    else if (match.GoogleAppointment != null)
                        name = match.GoogleAppointment.Subject;
                    NotificationReceived(String.Format("Syncing appointment {0} of {1}: {2} ...", i + 1, sync.Appointments.Count, name));
                }

                SyncAppointment(match, sync);
            }
        }
        public static void SyncAppointment(AppointmentMatch match, Syncronizer sync)
        {
            //Outlook.AppointmentItem outlookAppointment = match.OutlookAppointment;
            //Outlook.AppointmentItem googleAppointment = match.GoogleAppointment;

            //try
            //{
            if (match.GoogleAppointment == null && match.OutlookAppointment != null)
            {
                //no target appointment                               
                string googleAppointmenttId = AppointmentPropertiesUtils.GetOutlookGoogleAppointmentId(sync, match.OutlookAppointment);
                if (!string.IsNullOrEmpty(googleAppointmenttId))
                {
                    //Redundant check if exist, but in case an error occurred in MatchAppointments
                    Microsoft.Office.Interop.Outlook.AppointmentItem matchingGoogleAppointment = sync.GetGoogleAppointmentById(googleAppointmenttId);
                    if (matchingGoogleAppointment == null)
                        if (!sync.PromptDelete)
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
                            AppointmentPropertiesUtils.ResetOutlookGoogleAppointmentId(sync,match.OutlookAppointment);
                            break;
                        case DeleteResolution.DeleteOutlook:
                        case DeleteResolution.DeleteOutlookAlways:
                            //Avoid recreating a GoogleAppointment already existing
                            //==> Delete this OutlookAppointment instead if previous match existed but no match exists anymore
                            return;
                        default:
                            throw new ApplicationException("Cancelled");
                    }
                }

                if (sync.SyncOption == SyncOption.GoogleToOutlookOnly)
                {
                    sync.SkippedCount++;
                    Logger.Log(string.Format("Source Appointment not added to Target, because of SyncOption " + sync.SyncOption.ToString() + ": {0}", match.OutlookAppointment.Subject), EventType.Information);
                    return;
                }

                //create a Target appointment from Source appointment
                match.GoogleAppointment = Syncronizer.CreateOutlookAppointmentItem(Syncronizer.GoogleAppointmentsFolder);

                sync.SaveAppointment(match.OutlookAppointment, match.GoogleAppointment, Syncronizer.Source);

            }
            else if (match.OutlookAppointment == null && match.GoogleAppointment != null)
            {
                //no source appointment                               
                string outlookAppointmenttId = AppointmentPropertiesUtils.GetGoogleOutlookAppointmentId(sync.SyncProfile,match.GoogleAppointment);
                if (!string.IsNullOrEmpty(outlookAppointmenttId))
                {
                    if (!sync.PromptDelete)
                        sync.DeleteGoogleResolution = DeleteResolution.DeleteGoogleAlways;
                    else if (sync.DeleteGoogleResolution != DeleteResolution.DeleteGoogleAlways &&
                             sync.DeleteGoogleResolution != DeleteResolution.KeepGoogleAlways)
                    {
                        var r = new ConflictResolver();
                        sync.DeleteGoogleResolution = r.ResolveDelete(match.GoogleAppointment, Syncronizer.Target);
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
                    Logger.Log(string.Format("Target Appointment not added to Source, because of SyncOption " + sync.SyncOption.ToString() + ": {0}", match.GoogleAppointment.Subject), EventType.Information);
                    return;
                }

                //create a Source appointment from Target appointment
                match.OutlookAppointment = Syncronizer.CreateOutlookAppointmentItem(Syncronizer.OutlookAppointmentsFolder);

                sync.SaveAppointment(match.GoogleAppointment, match.OutlookAppointment, Syncronizer.Target);
            }
            else if (match.OutlookAppointment != null && match.GoogleAppointment != null)
            {
                //merge appointment details                

                //determine if this appointment pair were syncronized
                //DateTime? lastUpdated = GetOutlookPropertyValueDateTime(match.OutlookAppointment, sync.OutlookPropertyNameUpdated);
                DateTime? lastSynced = AppointmentPropertiesUtils.GetOutlookLastSync(sync, match.OutlookAppointment);
                if (lastSynced.HasValue)
                {
                    //appointment pair was syncronysed before.

                    //determine if target appointment was updated since last sync

                    //lastSynced is stored without seconds. take that into account.
                    DateTime lastUpdatedOutlook = match.OutlookAppointment.LastModificationTime.AddSeconds(-match.OutlookAppointment.LastModificationTime.Second);
                    DateTime lastUpdatedGoogle = match.GoogleAppointment.LastModificationTime.AddSeconds(-match.GoogleAppointment.LastModificationTime.Second);

                    //check if both outlok and target appointments where updated sync last sync
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
                                //overwrite target appointment
                                Logger.Log("Source and Target appointment have been updated, Source appointment is overwriting Target because of SyncOption " + sync.SyncOption + ": " + match.OutlookAppointment.Subject + ".", EventType.Information);
                                sync.SaveAppointment(match.OutlookAppointment, match.GoogleAppointment, Syncronizer.Source);
                                break;
                            case SyncOption.MergeGoogleWins:
                            case SyncOption.GoogleToOutlookOnly:
                                //overwrite outlook appointment
                                Logger.Log("Source and Target appointment have been updated, Target appointment is overwriting Source because of SyncOption " + sync.SyncOption + ": " + match.GoogleAppointment.Subject + ".", EventType.Information);
                                sync.SaveAppointment(match.GoogleAppointment, match.OutlookAppointment, Syncronizer.Target);
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
                                        sync.SaveAppointment(match.OutlookAppointment, match.GoogleAppointment, Syncronizer.Source);
                                        break;
                                    case ConflictResolution.GoogleWins:
                                    case ConflictResolution.GoogleWinsAlways:
                                        sync.SaveAppointment(match.GoogleAppointment, match.OutlookAppointment, Syncronizer.Target);
                                        break;
                                    default:
                                        throw new ApplicationException("Cancelled");
                                }
                                break;
                        }
                        return;
                    }


                    //check if source appointment was updated (with X second tolerance)
                    if (sync.SyncOption != SyncOption.GoogleToOutlookOnly &&
                        ((int)lastUpdatedOutlook.Subtract(lastSynced.Value).TotalSeconds > TimeTolerance ||
                         (int)lastUpdatedGoogle.Subtract(lastSynced.Value).TotalSeconds > TimeTolerance &&
                         sync.SyncOption == SyncOption.OutlookToGoogleOnly
                        )
                       )
                    {
                        //source appointment was changed or changed target appointment will be overwritten

                        if ((int)lastUpdatedGoogle.Subtract(lastSynced.Value).TotalSeconds > TimeTolerance &&
                            sync.SyncOption == SyncOption.OutlookToGoogleOnly)
                            Logger.Log("Target appointment has been updated since last sync, but Source appointment is overwriting Target because of SyncOption " + sync.SyncOption + ": " + match.OutlookAppointment.Subject + ".", EventType.Information);

                        sync.SaveAppointment(match.OutlookAppointment, match.GoogleAppointment, Syncronizer.Source);

                        //at the moment use source as "master" source of appointments - in the event of a conflict target appointment will be overwritten.
                        //TODO: control conflict resolution by SyncOption
                        return;
                    }

                    //check if target appointment was updated (with X second tolerance)
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

                        sync.SaveAppointment(match.GoogleAppointment, match.OutlookAppointment, Syncronizer.Target);
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
                            //overwrite target appointment
                            sync.SaveAppointment(match.OutlookAppointment, match.GoogleAppointment, Syncronizer.Source);
                            break;
                        case SyncOption.MergeGoogleWins:
                        case SyncOption.GoogleToOutlookOnly:
                            //overwrite outlook appointment
                            sync.SaveAppointment(match.GoogleAppointment, match.OutlookAppointment, Syncronizer.Target);
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
                                    sync.SaveAppointment(match.OutlookAppointment, match.GoogleAppointment, Syncronizer.Source);
                                    break;
                                case ConflictResolution.GoogleWins:
                                case ConflictResolution.GoogleWinsAlways:
                                    sync.SaveAppointment(match.GoogleAppointment, match.OutlookAppointment, Syncronizer.Target);
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
        public EventEntry GoogleAppointment;
        public readonly List<Outlook.AppointmentItem> AllGoogleAppointmentMatches = new List<Outlook.AppointmentItem>(1);
        public Outlook.AppointmentItem LastGoogleAppointment;

        public AppointmentMatch(Outlook.AppointmentItem outlookAppointment, EventEntry googleAppointment)
        {
            OutlookAppointment = outlookAppointment;
            GoogleAppointment = googleAppointment;
        }

        public void AddGoogleAppointment(EventEntry googleAppointment)
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
