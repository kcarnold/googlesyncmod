using System;
using System.Collections.Generic;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;
using Google.GData.Calendar;
using Google.GData.Extensions;

namespace GoContactSyncMod
{

    internal static class AppointmentSync
    {
        private const string dateFormat = "yyyyMMdd";
        private const string timeFormat = "HHmmss";
        internal static DateTime outlookDateMin = new DateTime(4501, 1, 1);
        internal static DateTime outlookDateMax = new DateTime(4500, 12, 31);

        const string DTSTART = "DTSTART";
        const string DTEND = "DTEND";
        const string RRULE = "RRULE";
        const string FREQ = "FREQ";
        const string DAILY = "DAILY";
        const string WEEKLY = "WEEKLY";
        const string MONTHLY = "MONTHLY";
        const string YEARLY = "YEARLY";

        const string BYMONTH = "BYMONTH";
        const string BYMONTHDAY = "BYMONTHDAY";
        const string BYDAY = "BYDAY";
        const string BYSETPOS= "BYSETPOS";

        const string VALUE = "VALUE";
        const string DATE = "DATE";
        const string DATETIME = "DATE-TIME";
        const string INTERVAL = "INTERVAL";
        const string COUNT = "COUNT";
        const string UNTIL = "UNTIL";
        const string TZID = "TZID";

        const string MO = "MO";
        const string TU = "TU";
        const string WE = "WE";
        const string TH = "TH";
        const string FR = "FR";
        const string SA = "SA";
        const string SU = "SU";

        /// <summary>
        /// Updates Outlook appointments (calendar) to Google Calendar
        /// </summary>
        public static void UpdateAppointment(Outlook.AppointmentItem master, EventEntry slave)
        {            
            slave.Title.Text = master.Subject;
            

            ////foreach (Outlook.Attachment attachment in master.Attachments)
            ////    slave.Attachments.Add(master.Attachments);
            
            slave.Content.Content = master.Body;
            slave.Status = Google.GData.Calendar.EventEntry.EventStatus.CONFIRMED;
            if (master.BusyStatus.Equals(Outlook.OlBusyStatus.olTentative))
                slave.Status = Google.GData.Calendar.EventEntry.EventStatus.TENTATIVE;
            
            //ToDo:slave.Categories = master.Categories;
            //slave.Duration = master.Duration;

            var location = new Google.GData.Extensions.Where();
            location.ValueString = master.Location;
            slave.Locations.Clear();
            slave.Locations.Add(location);

            slave.Times.Clear();
            if (!master.IsRecurring)
                slave.Times.Add(new Google.GData.Extensions.When(master.Start, master.End, master.AllDayEvent));
            ////slave.StartInStartTimeZone = master.StartInStartTimeZone;
            ////slave.StartTimeZone = master.StartTimeZone;
            ////slave.StartUTC = master.StartUTC;

            slave.Participants.Clear();
            int i = 0;
            foreach (Outlook.Recipient recipient in master.Recipients)
            {
             
                var participant = new Who();
                participant.Email = recipient.Address!=null? recipient.Address:recipient.Name;

                participant.Rel = (i == 0 ? Who.RelType.EVENT_ORGANIZER : Who.RelType.EVENT_ATTENDEE);
                slave.Participants.Add(participant);
                i++;
            }
            //slave.RequiredAttendees = master.RequiredAttendees;
            //slave.OptionalAttendees = master.OptionalAttendees;

            if (slave.Reminders != null)
            {
                slave.Reminders.Clear();
                if (master.ReminderSet)
                {
                    var reminder = new Google.GData.Extensions.Reminder();
                    reminder.Minutes = master.ReminderMinutesBeforeStart;
                    reminder.Method = Google.GData.Extensions.Reminder.ReminderMethod.alert;
                    slave.Reminders.Add(reminder);
                }
            }

            //slave.Resources = master.Resources;
            //slave.RTFBody = master.RTFBody;

            UpdateRecurrence(master, slave);

            
        }

        /// <summary>
        /// Updates Outlook appointments (calendar) to Google Calendar
        /// </summary>
        public static void UpdateAppointment(EventEntry master, Outlook.AppointmentItem slave)
        {
            slave.Subject = master.Title.Text;            

            //foreach (Outlook.Attachment attachment in master.Attachments)
            //    slave.Attachments.Add(master.Attachments);

            slave.Body = master.Content.Content;

            slave.BusyStatus = Outlook.OlBusyStatus.olBusy;
            if (master.Status.Equals(Google.GData.Calendar.EventEntry.EventStatus.TENTATIVE))
                slave.BusyStatus = Outlook.OlBusyStatus.olTentative;
             else if (master.Status.Equals(Google.GData.Calendar.EventEntry.EventStatus.CANCELED))
                slave.BusyStatus = Outlook.OlBusyStatus.olFree;
                  
            //slave.Categories = master.Categories;
            //slave.Duration = master.Duration;

            slave.Location = string.Empty;
            if (master.Locations.Count > 0)
                slave.Location = master.Locations[0].ValueString;

            if (master.Times.Count != 1 && master.Recurrence == null)
                Logger.Log("Google Appointment with multiple or no times found: " + master.Title.Text + " - " + (master.Times.Count == 0 ? null : master.Times[0].StartTime.ToString()), EventType.Warning);

            if (master.RecurrenceException != null)
                Logger.Log("Google Appointment with RecurrenceException found: " + master.Title.Text + " - " + (master.Times.Count == 0 ? null : master.Times[0].StartTime.ToString()), EventType.Warning);


            if (master.Times.Count == 1 || master.Recurrence == null)
            {//only sync times for not recurrent events
                //ToDo: How to sync recurrence exceptions?
                slave.AllDayEvent = master.Times[0].AllDay;
                slave.Start = master.Times[0].StartTime;
                slave.End = master.Times[0].EndTime;
            }
            
            //slave.StartInStartTimeZone = master.StartInStartTimeZone;
            //slave.StartTimeZone = master.StartTimeZone;
            //slave.StartUTC = master.StartUTC;
           
            for (int i = slave.Recipients.Count; i > 0; i--)
                slave.Recipients.Remove(i);

            foreach (Who participant in master.Participants)
            {

                if (participant.Rel == Who.RelType.EVENT_ORGANIZER && participant.Email != Syncronizer.UserName)
                    //ToDo: Doesn't Work, because Organizer cannot be set on Outlook side. Maybe somehow at least on behalf?
                    //slave.GetOrganizer().Address = participant.Email;
                    //Workaround: Assign organizer at least as first participant                    
                    slave.Recipients.Add(participant.Email);                
                    
            }
            foreach (Who participant in master.Participants)
            {

                if (participant.Rel != Who.RelType.EVENT_ORGANIZER && participant.Email != Syncronizer.UserName)                    
                    slave.Recipients.Add(participant.Email);

            }
            //slave.RequiredAttendees = master.RequiredAttendees;		master.Title.Text	"Rodeln"	string

            //slave.OptionalAttendees = master.OptionalAttendees;
            //slave.Resources = master.Resources;

            slave.ReminderSet = false;
            if (master.Reminder != null && !master.Reminder.Method.Equals(Google.GData.Extensions.Reminder.ReminderMethod.none) && master.Reminder.AbsoluteTime >= DateTime.Now)
            { 
                slave.ReminderSet = true;
                slave.ReminderMinutesBeforeStart = master.Reminder.Minutes;
            }

            
            //slave.RTFBody = master.RTFBody;

  
            //ToDo: Check Exceptions, how to sync
            
            UpdateRecurrence(master, slave);
            
            if (!IsOrganizer(GetOrganizer(master)) || !IsOrganizer(GetOrganizer(slave), slave))
                slave.MeetingStatus = Outlook.OlMeetingStatus.olMeetingReceived;
            
        }

        public static void UpdateRecurrence(Outlook.AppointmentItem master, EventEntry slave)
        {
            try
            {                               

                if (master.RecurrenceState != Outlook.OlRecurrenceState.olApptMaster)
                    return;

                if (!master.IsRecurring)
                {
                    if (slave.Recurrence != null)
                        slave.Recurrence = null;
                    return;
                }

                Outlook.RecurrencePattern masterRecurrence = master.GetRecurrencePattern();

                var slaveRecurrence = new Recurrence();
                

                string format = dateFormat;
                string key = VALUE + "=" + DATE;
                if (!master.AllDayEvent)
                {
                    format += "'T'"+timeFormat;
                    key = VALUE + "=" + DATETIME;
                }

                //ToDo: Find a way how to handle timezones, per default GMT (UTC+0:00) is taken
                if (master.StartTimeZone.ID == "W. Europe Standard Time")
                    key = TZID + "=" + "Europe/Berlin";

                DateTime date = masterRecurrence.PatternStartDate.Date;
                DateTime time = new DateTime(date.Year, date.Month, date.Day, masterRecurrence.StartTime.Hour, masterRecurrence.StartTime.Minute, masterRecurrence.StartTime.Second);
                
                slaveRecurrence.Value += DTSTART;                    
                slaveRecurrence.Value += ";" + key + ":" + time.ToString(format) + "\r\n";
                                
                time = new DateTime(date.Year, date.Month, date.Day, masterRecurrence.EndTime.Hour, masterRecurrence.EndTime.Minute, masterRecurrence.EndTime.Second);               
                
                slaveRecurrence.Value += DTEND;
                slaveRecurrence.Value += ";"+key+":" + time.ToString(format) + "\r\n";
                
                slaveRecurrence.Value += RRULE + ":" + FREQ +"=";
                switch (masterRecurrence.RecurrenceType)
                {
                    case Outlook.OlRecurrenceType.olRecursDaily: slaveRecurrence.Value += DAILY; break;
                    case Outlook.OlRecurrenceType.olRecursWeekly: slaveRecurrence.Value += WEEKLY; break;
                    case Outlook.OlRecurrenceType.olRecursMonthly: 
                    case Outlook.OlRecurrenceType.olRecursMonthNth: slaveRecurrence.Value += MONTHLY; break;
                    case Outlook.OlRecurrenceType.olRecursYearly:
                    case Outlook.OlRecurrenceType.olRecursYearNth: slaveRecurrence.Value += YEARLY; break;
                    default: throw new NotSupportedException("RecurrenceType not supported by Google: " + masterRecurrence.RecurrenceType);                                     
                }

                string byDay = string.Empty;
                if ((masterRecurrence.DayOfWeekMask & Outlook.OlDaysOfWeek.olMonday) == Outlook.OlDaysOfWeek.olMonday)
                    byDay = MO;
                if ((masterRecurrence.DayOfWeekMask & Outlook.OlDaysOfWeek.olTuesday) == Outlook.OlDaysOfWeek.olTuesday)
                    byDay += (string.IsNullOrEmpty(byDay) ? "" : ",") + TU;
                if ((masterRecurrence.DayOfWeekMask & Outlook.OlDaysOfWeek.olWednesday) == Outlook.OlDaysOfWeek.olWednesday)
                    byDay += (string.IsNullOrEmpty(byDay) ? "" : ",") + WE;
                if ((masterRecurrence.DayOfWeekMask & Outlook.OlDaysOfWeek.olThursday) == Outlook.OlDaysOfWeek.olThursday)
                    byDay += (string.IsNullOrEmpty(byDay) ? "" : ",") + TH;
                if ((masterRecurrence.DayOfWeekMask & Outlook.OlDaysOfWeek.olFriday) == Outlook.OlDaysOfWeek.olFriday)
                    byDay += (string.IsNullOrEmpty(byDay) ? "" : ",") + FR;
                if ((masterRecurrence.DayOfWeekMask & Outlook.OlDaysOfWeek.olSaturday) == Outlook.OlDaysOfWeek.olSaturday)
                   byDay += (string.IsNullOrEmpty(byDay) ? "" : ",") + SA;
                if ((masterRecurrence.DayOfWeekMask & Outlook.OlDaysOfWeek.olSunday) == Outlook.OlDaysOfWeek.olSunday)
                   byDay += (string.IsNullOrEmpty(byDay) ? "" : ",") + SU;

                if (!string.IsNullOrEmpty(byDay))
                {
                    if (masterRecurrence.Instance >= 1 && masterRecurrence.Instance <= 4)
                        byDay = masterRecurrence.Instance + byDay;
                    else if (masterRecurrence.Instance == 5)
                        slaveRecurrence.Value += ";" + BYSETPOS + "=-1";
                    else
                        throw new NotSupportedException("Outlook Appointment Instances 1-4 and 5 (last) are allowed but was: " + masterRecurrence.Instance);
                    slaveRecurrence.Value += ";" + BYDAY + "=" + byDay;
                }

                if (masterRecurrence.DayOfMonth != 0)
                    slaveRecurrence.Value += ";" + BYMONTHDAY + "=" + masterRecurrence.DayOfMonth;

                if (masterRecurrence.MonthOfYear != 0)
                    slaveRecurrence.Value += ";" + BYMONTH + "=" + masterRecurrence.MonthOfYear;

                if (masterRecurrence.RecurrenceType != Outlook.OlRecurrenceType.olRecursYearly && 
                    masterRecurrence.RecurrenceType != Outlook.OlRecurrenceType.olRecursYearNth &&
                    masterRecurrence.Interval > 1 ||
                    masterRecurrence.Interval > 12)
                {
                    if (masterRecurrence.RecurrenceType != Outlook.OlRecurrenceType.olRecursYearly &&
                        masterRecurrence.RecurrenceType != Outlook.OlRecurrenceType.olRecursYearNth)
                        slaveRecurrence.Value += ";" + INTERVAL+ "=" + masterRecurrence.Interval;
                    else
                        slaveRecurrence.Value += ";" + INTERVAL + "=" + masterRecurrence.Interval/12;
                }
                
                format = dateFormat;
                if (masterRecurrence.PatternEndDate.Date != outlookDateMin &&
                    masterRecurrence.PatternEndDate.Date != outlookDateMax)
                {
                    slaveRecurrence.Value += ";" + UNTIL + "=" + masterRecurrence.PatternEndDate.Date.ToString(format);
                }
                else if (masterRecurrence.Occurrences > 0)
                {
                    slaveRecurrence.Value += ";" + COUNT + "=" + masterRecurrence.Occurrences;
                }

                slave.Recurrence = slaveRecurrence;
            }
            catch (Exception ex)
            {
                ErrorHandler.Handle(ex);
            }


        }

        /// <summary>
        /// Update Recurrence pattern from Google by parsing the string, see also specification http://tools.ietf.org/html/rfc2445
        /// </summary>
        /// <param name="master"></param>
        /// <param name="slave"></param>
        public static void UpdateRecurrence(EventEntry master, Outlook.AppointmentItem slave)
        {
            Recurrence masterRecurrence = master.Recurrence;   
            if (masterRecurrence == null)
            {
                if (slave.IsRecurring)
                    slave.ClearRecurrencePattern();

                return;
            }

            try
            {  
                 
                 Outlook.RecurrencePattern slaveRecurrence = slave.GetRecurrencePattern();      



                string[] patterns = masterRecurrence.Value.Split(new char[] {'\r','\n'}, StringSplitOptions.RemoveEmptyEntries);
                foreach (string pattern in patterns)
                {
                    if (pattern.StartsWith(DTSTART)) 
                    {
                        //DTSTART;VALUE=DATE:20070501
                        //DTSTART;TZID=US-Eastern:19970905T090000
                        string[] parts = pattern.Split(new char[] {';',':'});
                        
                        slaveRecurrence.StartTime = GetDateTime(parts[parts.Length-1]);
                        slaveRecurrence.PatternStartDate = GetDateTime(parts[parts.Length - 1]);
                        break;
                    }
                }

                foreach (string pattern in patterns)
                {
                    if (pattern.StartsWith(DTEND))
                    {
                        string[] parts = pattern.Split(new char[] { ';', ':' });
                        
                        slaveRecurrence.EndTime = GetDateTime(parts[parts.Length-1]);
                        
                        break;
                    }
                }

                foreach (string pattern in patterns)
                {
                    if (pattern.StartsWith(RRULE))
                    {
                        string[] parts = pattern.Split(new char[] { ';', ':' });

                        int instance = 0;
                        foreach (string part in parts)
                        {
                            if (part.StartsWith(BYDAY))
                            {
                                string[] days = part.Split(',');
                                foreach (string day in days)
                                {
                                    string dayValue = day.Substring(day.IndexOf("=") + 1);
                                    if (dayValue.StartsWith("1"))
                                        instance = 1;
                                    else if (dayValue.StartsWith("2"))
                                        instance = 2;
                                    else if (dayValue.StartsWith("3"))
                                        instance = 3;
                                    else if (dayValue.StartsWith("4"))
                                        instance = 4;


                                    break;
                                }

                                break;
                            }

                        }

                        foreach (string part in parts)
                        {
                            if (part.StartsWith(FREQ))
                            {
                                switch (part.Substring(part.IndexOf('=') + 1))
                                {
                                    case DAILY: slaveRecurrence.RecurrenceType = Outlook.OlRecurrenceType.olRecursDaily; break;
                                    case WEEKLY: slaveRecurrence.RecurrenceType = Outlook.OlRecurrenceType.olRecursWeekly; break;
                                    case MONTHLY:
                                        if (instance == 0)
                                            slaveRecurrence.RecurrenceType = Outlook.OlRecurrenceType.olRecursMonthly;
                                        else
                                            slaveRecurrence.RecurrenceType = Outlook.OlRecurrenceType.olRecursMonthNth;
                                        break;
                                    case YEARLY:
                                        if (instance == 0)
                                            slaveRecurrence.RecurrenceType = Outlook.OlRecurrenceType.olRecursYearly;
                                        else
                                            slaveRecurrence.RecurrenceType = Outlook.OlRecurrenceType.olRecursYearNth;
                                        break;
                                    default: throw new NotSupportedException("RecurrenceType not supported by Outlook: " + part);
                                                                        
                                }
                                break;
                            }

                        }

                        foreach (string part in parts)
                        {
                            if (part.StartsWith(BYDAY))
                            {
                                string[] days = part.Split(',');
                                foreach (string day in days)
                                {
                                    string dayValue = day.Substring(day.IndexOf("=") + 1);
                                    
                                    switch (dayValue.Trim(new char[] { '1', '2', '3', '4', ' ' }))
                                    {
                                        case MO: slaveRecurrence.DayOfWeekMask = slaveRecurrence.DayOfWeekMask | Outlook.OlDaysOfWeek.olMonday; break;
                                        case TU: slaveRecurrence.DayOfWeekMask = slaveRecurrence.DayOfWeekMask | Outlook.OlDaysOfWeek.olTuesday; break;
                                        case WE: slaveRecurrence.DayOfWeekMask = slaveRecurrence.DayOfWeekMask | Outlook.OlDaysOfWeek.olWednesday; break;
                                        case TH: slaveRecurrence.DayOfWeekMask = slaveRecurrence.DayOfWeekMask | Outlook.OlDaysOfWeek.olThursday; break;
                                        case FR: slaveRecurrence.DayOfWeekMask = slaveRecurrence.DayOfWeekMask | Outlook.OlDaysOfWeek.olFriday; break;
                                        case SA: slaveRecurrence.DayOfWeekMask = slaveRecurrence.DayOfWeekMask | Outlook.OlDaysOfWeek.olSaturday; break;
                                        case SU: slaveRecurrence.DayOfWeekMask = slaveRecurrence.DayOfWeekMask | Outlook.OlDaysOfWeek.olSunday; break;

                                    }
                                    //Don't break because multiple days possible;
                                }

                                break;
                            }

                            if (part.StartsWith(BYSETPOS))
                            {
                                string pos = part.Substring(part.IndexOf("=") + 1);

                                if (pos.Trim() == "-1")
                                    slaveRecurrence.Instance = 5;
                                else
                                    throw new Exception("Only 'BYSETPOS=-1' is allowed by Outlook, but it was: " +part);
                            }
                        }

                        foreach (string part in parts)
                        {
                            if (part.StartsWith(COUNT))
                            {
                                if (master.Times.Count > 0)
                                    slaveRecurrence.PatternStartDate = master.Times[0].StartTime;
                                slaveRecurrence.Occurrences = int.Parse(part.Substring(part.IndexOf('=') + 1));
                                break;
                            }
                            else if (part.StartsWith(UNTIL))
                            {
                                //either UNTIL or COUNT may appear in a 'recur',
                                //but UNTIL and COUNT MUST NOT occur in the same 'recur'
                                slaveRecurrence.PatternEndDate = GetDateTime(part.Substring(part.IndexOf('=') + 1));
                                break;
                            }
                        }

                        foreach (string part in parts)
                        {
                            if (part.StartsWith(INTERVAL))
                            {
                                slaveRecurrence.Interval = int.Parse(part.Substring(part.IndexOf('=') + 1));
                                break;
                            }

                        }





                        foreach (string part in parts)
                        {
                            if (part.StartsWith(BYMONTHDAY))
                            {
                                slaveRecurrence.DayOfMonth = int.Parse(part.Substring(part.IndexOf('=') + 1));
                                break;
                            }
                        }

                        foreach (string part in parts)
                        {
                            if (part.StartsWith(BYMONTH + "="))
                            {
                                slaveRecurrence.MonthOfYear = int.Parse(part.Substring(part.IndexOf('=') + 1));
                                break;
                            }
                        }


                        break;
                    }

                }
            }
            catch (Exception ex)
            {
                ErrorHandler.Handle(ex);
            }


        }

        public static bool UpdateRecurrenceExceptions(Outlook.AppointmentItem master, EventEntry slave)
        {
            
            bool ret = false;

            Outlook.Exceptions exceptions = master.GetRecurrencePattern().Exceptions;

            if (exceptions == null || exceptions.Count == 0)
                ret= false;
            else
            {
                foreach (Outlook.Exception exception in exceptions)
                {
                    if (!exception.Deleted)
                    {//Add exception time (but only if in 
                        if ((Syncronizer.MonthsInPast == 0 ||
                             exception.AppointmentItem.End >= DateTime.Now.AddMonths(-Syncronizer.MonthsInPast)) &&
                             (Syncronizer.MonthsInFuture == 0 ||
                             exception.AppointmentItem.Start <= DateTime.Now.AddMonths(Syncronizer.MonthsInFuture)))
                        {
                            slave.Times.Add(new Google.GData.Extensions.When(exception.AppointmentItem.Start, exception.AppointmentItem.Start, exception.AppointmentItem.AllDayEvent));
                            ret = true;
                        }
                    }
                    else
                    {//ToDo: Delete exception time
                        //for (int i=slave.Times.Count;i>0;i--)
                        //{
                        //    When time = slave.Times[i-1];
                        //    if (time.StartTime.Equals(exception.AppointmentItem.Start))
                        //    {
                        //        slave.Times.Remove(time);
                        //        ret = true;
                        //        break;
                        //    }
                        //}
                    }
                }
            }

            return ret;
        }

        private static DateTime GetDateTime(string dateTime)
        {
            string format = dateFormat;
            if (dateTime.Contains("T"))
                format += "'T'"+timeFormat;
            if (dateTime.EndsWith("Z"))
                format += "'Z'";
            return DateTime.ParseExact(dateTime, format, new System.Globalization.CultureInfo("en-US"));
        }

        internal static Who GetOrganizer(EventEntry googleAppointment)
        {
            foreach (var person in googleAppointment.Participants)
            {

                if (person.Rel == Who.RelType.EVENT_ORGANIZER)
                {
                    return person;
                }
            }
            return null;
        }

        internal static bool IsOrganizer(Who person)
        {
            if (person != null && person.Email != null && person.Email.Trim().Equals(Syncronizer.UserName.Trim(), StringComparison.InvariantCultureIgnoreCase))
                return true;
            else
                return false;
        }

        internal static string GetOrganizer(Outlook.AppointmentItem outlookAppointment)
        {
            Outlook.AddressEntry organizer = outlookAppointment.GetOrganizer();
            if (organizer != null)
            {
                if (string.IsNullOrEmpty(organizer.Address))
                    return organizer.Address;
                else
                    return organizer.Name;
            }

            return outlookAppointment.Organizer;            


        }

        internal static bool IsOrganizer(string person, Outlook.AppointmentItem outlookAppointment)
        {
            if (!string.IsNullOrEmpty(person) && 
                (person.Trim().Equals(outlookAppointment.Session.CurrentUser.Address, StringComparison.InvariantCultureIgnoreCase) || 
                person.Trim().Equals(outlookAppointment.Session.CurrentUser.Name, StringComparison.InvariantCultureIgnoreCase)
                ))
                return true;
            else
                return false;
        }

    }
}
