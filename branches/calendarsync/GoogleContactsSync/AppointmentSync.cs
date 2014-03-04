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

        //internal static DateTime outlookDateMin = new DateTime(4501, 1, 1);
        //internal static DateTime outlookDateMax = new DateTime(4500, 12, 31);
        const string DTSTART = "DTSTART";
        const string DTEND = "DTEND";
        const string RRULE = "RRULE";
        const string FREQ = "FREQ";
        const string DAILY = "DAILY";
        const string WEEKLY = "WEEKLY";
        const string MONTHLY = "MONTHLY";
        const string YEARLY = "YEARLY";



        /// <summary>
        /// Updates Outlook appointments (calendar) to Google Calendar
        /// </summary>
        public static void Update(Outlook.AppointmentItem master, EventEntry slave)
        {            
            slave.Title.Text = master.Subject;
            

            ////foreach (Outlook.Attachment attachment in master.Attachments)
            ////    slave.Attachments.Add(master.Attachments);
            
            slave.Content.Content = master.Body;
            slave.Status = Google.GData.Calendar.EventEntry.EventStatus.CONFIRMED;
            if (master.BusyStatus.Equals(Outlook.OlBusyStatus.olTentative))
                slave.Status = Google.GData.Calendar.EventEntry.EventStatus.TENTATIVE;
            
            //slave.Categories = master.Categories;
            //slave.Duration = master.Duration;

            var location = new Google.GData.Extensions.Where();
            location.ValueString = master.Location;
            slave.Locations.Clear();
            slave.Locations.Add(location);

            slave.Times.Clear();
            slave.Times.Add(new Google.GData.Extensions.When(master.Start, master.End, master.AllDayEvent));
            ////slave.StartInStartTimeZone = master.StartInStartTimeZone;
            ////slave.StartTimeZone = master.StartTimeZone;
            ////slave.StartUTC = master.StartUTC;
            
            //slave.RequiredAttendees = master.RequiredAttendees;
            //slave.OptionalAttendees = master.OptionalAttendees;
            slave.Reminder = null;
            if (master.ReminderSet)
            {
                var reminder = new Google.GData.Extensions.Reminder();
                reminder.Minutes = master.ReminderMinutesBeforeStart;
                reminder.Method = Google.GData.Extensions.Reminder.ReminderMethod.alert;
                slave.Reminder = reminder;
            }

            //slave.Resources = master.Resources;
            //slave.RTFBody = master.RTFBody;

            UpdateRecurrence(master, slave);
            
        }

        /// <summary>
        /// Updates Outlook appointments (calendar) to Google Calendar
        /// </summary>
        public static void Update(EventEntry master, Outlook.AppointmentItem slave)
        {
            slave.Subject = master.Title.Text;            

            //foreach (Outlook.Attachment attachment in master.Attachments)
            //    slave.Attachments.Add(master.Attachments);

            slave.Body = master.Content.Content;

            slave.BusyStatus = Outlook.OlBusyStatus.olBusy;
            if (master.Status.Equals(Google.GData.Calendar.EventEntry.EventStatus.TENTATIVE))
                slave.BusyStatus = Outlook.OlBusyStatus.olTentative;
             if (master.Status.Equals(Google.GData.Calendar.EventEntry.EventStatus.CANCELED))
                slave.BusyStatus = Outlook.OlBusyStatus.olFree;
                  
            //slave.Categories = master.Categories;
            //slave.Duration = master.Duration;

            slave.Location = string.Empty;
            if (master.Locations.Count > 0)
                slave.Location = master.Locations[0].ValueString;

            if (master.Times.Count > 0)
            {
                slave.AllDayEvent = master.Times[0].AllDay;
                slave.Start = master.Times[0].StartTime;
                slave.End = master.Times[0].EndTime;
                
            }
            //slave.StartInStartTimeZone = master.StartInStartTimeZone;
            //slave.StartTimeZone = master.StartTimeZone;
            //slave.StartUTC = master.StartUTC;

            //ToDo: slave.RequiredAttendees = master.Participants;
            //ToDo: slave.OptionalAttendees = master.OptionalAttendees;
          
            slave.ReminderSet = false;
            if (master.Reminder != null && !master.Reminder.Method.Equals(Google.GData.Extensions.Reminder.ReminderMethod.none) && master.Reminder.AbsoluteTime >= DateTime.Now)
            { 
                slave.ReminderSet = true;
                slave.ReminderMinutesBeforeStart = master.Reminder.Minutes;
            }

            //ToDo: slave.Resources = master.Resources;
            //slave.RTFBody = master.RTFBody;

            //ToDo: Check, why some appointments are created twice (e.g. SCRUM Master Certification, ...) and don't keep olRecurDaily
            //ToDo: Check Exceptions, how to sync
            
            UpdateRecurrence(master, slave);
            
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
                slaveRecurrence.Value += DTSTART;
                slaveRecurrence.Value += ";VALUE=DATE:" + masterRecurrence.PatternStartDate + "\r\n";
                
                slaveRecurrence.Value += DTEND;
                slaveRecurrence.Value += ";VALUE=DATE:" + masterRecurrence.PatternEndDate + "\r\n";

                slaveRecurrence.Value += RRULE + ":" + FREQ +"=";
                switch (masterRecurrence.RecurrenceType)
                {
                    case Outlook.OlRecurrenceType.olRecursDaily: slaveRecurrence.Value += DAILY; break;
                    case Outlook.OlRecurrenceType.olRecursWeekly: slaveRecurrence.Value += WEEKLY; break;
                    case Outlook.OlRecurrenceType.olRecursMonthly: slaveRecurrence.Value += MONTHLY; break;
                    case Outlook.OlRecurrenceType.olRecursYearly: slaveRecurrence.Value += YEARLY; break;
                    default: throw new NotSupportedException("RecurrenceType not supported by Google: " + masterRecurrence.RecurrenceType);                                     
                }
                


                //ToDo:";BYDAY=Tu;UNTIL=20070904\r\n";


                //ToDo: Implement Recurrence Update
                //if (masterRecurrence.DayOfMonth > 0)
                //    slaveRecurrence.dDayOfMonth = masterRecurrence.DayOfMonth;
                //if (masterRecurrence.MonthOfYear > 0)
                //    slaveRecurrence.MonthOfYear = masterRecurrence.MonthOfYear;
                //if (masterRecurrence.DayOfWeekMask > 0)
                //    slaveRecurrence.DayOfWeekMask = masterRecurrence.DayOfWeekMask;
                //slaveRecurrence.Duration = masterRecurrence.Duration;
                ////if (master.StartTime.Date < new DateTime(1900, 1, 1)
                ////    || master.StartTime.Date > new DateTime(2100, 1, 1))
                //slaveRecurrence.StartTime = masterRecurrence.StartTime;
                ////if (master.EndTime.Date != new DateTime(1601, 1, 2))
                ////if (master.EndTime.Date < new DateTime(1900, 1, 1)
                ////    || master.EndTime.Date > new DateTime(2100, 1, 1))
                //slaveRecurrence.EndTime = masterRecurrence.EndTime;
                //slaveRecurrence.NoEndDate = masterRecurrence.NoEndDate;

                ////slave.Instance = master.Instance;
                //slaveRecurrence.Interval = masterRecurrence.Interval;
                ////slave.Occurrences = master.Occurrences;
                ////if (master.PatternEndDate.Date != new DateTime(4500, 12, 31))
                //if (masterRecurrence.PatternEndDate.Date > new DateTime(1900, 1, 1)
                //     && masterRecurrence.PatternEndDate.Date < new DateTime(2100, 1, 1))
                //    slaveRecurrence.PatternEndDate = masterRecurrence.PatternEndDate;
                ////if (master.PatternStartDate != new DateTime(4501, 1, 1))

                //slaveRecurrence.PatternStartDate = masterRecurrence.PatternStartDate;

                slave.Recurrence = slaveRecurrence;
            }
            catch (Exception ex)
            {
                throw ex;
            }


        }

        /// <summary>
        /// http://tools.ietf.org/html/rfc2445
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
                        //Todo: consider also DTSTART;VALUE=DATE:20070501
                        //Currently: DTSTART;TZID=US-Eastern:19970905T090000
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

                        foreach (string part in parts)
                        {
                            if (part.StartsWith(FREQ))
                            {
                                switch (part.Substring(part.IndexOf('=') + 1))
                                {
                                    case DAILY: slaveRecurrence.RecurrenceType = Outlook.OlRecurrenceType.olRecursDaily; break;
                                    case WEEKLY: slaveRecurrence.RecurrenceType = Outlook.OlRecurrenceType.olRecursWeekly; break;
                                    case MONTHLY: slaveRecurrence.RecurrenceType = Outlook.OlRecurrenceType.olRecursMonthly; break;
                                    case YEARLY: slaveRecurrence.RecurrenceType = Outlook.OlRecurrenceType.olRecursYearly; break;
                                    default: throw new NotSupportedException("RecurrenceType not supported by Outlook: " + part);
                                    //ToDo: Outlook.OlRecurrenceType.olRecursMonthNth
                                    //ToDo: Outlook.OlRecurrenceType.olRecursYearNth                                        
                                }
                                break;
                            }
                            
                        }

                        foreach (string part in parts)
                        {
                            if (part.StartsWith("COUNT"))
                            {
                                if (master.Times.Count > 0)
                                    slaveRecurrence.PatternStartDate = master.Times[0].StartTime;
                                slaveRecurrence.Occurrences = int.Parse(part.Substring(part.IndexOf('=') + 1));
                                break;
                            }
                            else if (part.StartsWith("UNTIL"))
                            {
                                //either UNTIL or COUNT may appear in a 'recur',
                                //but UNTIL and COUNT MUST NOT occur in the same 'recur'
                                slaveRecurrence.PatternEndDate = GetDateTime(part.Substring(part.IndexOf('=') + 1));
                                break;
                            }
                        }

                        foreach (string part in parts)
                        {
                            if (part.StartsWith("INTERVAL"))
                            {                                
                                slaveRecurrence.Interval = int.Parse(part.Substring(part.IndexOf('=') + 1));
                                break;
                            }
                            
                        }                       

                        //if (slaveRecurrence.RecurrenceType == Outlook.OlRecurrenceType.olRecursWeekly)
                        //{
                            foreach (string part in parts)
                            {
                                if (part.StartsWith("BYDAY"))
                                {
                                    string[] days = part.Split(',');
                                    foreach (string day in days)
                                    {
                                        switch (day.Trim())
                                        {
                                            case "MO": slaveRecurrence.DayOfWeekMask = slaveRecurrence.DayOfWeekMask | Outlook.OlDaysOfWeek.olMonday; break;
                                            case "TU": slaveRecurrence.DayOfWeekMask = slaveRecurrence.DayOfWeekMask | Outlook.OlDaysOfWeek.olTuesday; break;
                                            case "WE": slaveRecurrence.DayOfWeekMask = slaveRecurrence.DayOfWeekMask | Outlook.OlDaysOfWeek.olWednesday; break;
                                            case "TH": slaveRecurrence.DayOfWeekMask = slaveRecurrence.DayOfWeekMask | Outlook.OlDaysOfWeek.olThursday; break;
                                            case "FR": slaveRecurrence.DayOfWeekMask = slaveRecurrence.DayOfWeekMask | Outlook.OlDaysOfWeek.olFriday; break;
                                            case "SA": slaveRecurrence.DayOfWeekMask = slaveRecurrence.DayOfWeekMask | Outlook.OlDaysOfWeek.olSaturday; break;
                                            case "SO": slaveRecurrence.DayOfWeekMask = slaveRecurrence.DayOfWeekMask | Outlook.OlDaysOfWeek.olSunday; break;
                                            //ToDo: Check how to interprete 1MO for 1st Monday of a month or year

                                        }
                                        //Don't break because multiple days;
                                    }

                                    break;
                                }

                            }

                            //break;
                        //}
                        //else if (slaveRecurrence.RecurrenceType == Outlook.OlRecurrenceType.olRecursYearly)
                        //{
                            foreach (string part in parts)
                            {
                                if (part.StartsWith("BYMONTHDAY"))
                                {
                                    slaveRecurrence.DayOfMonth = int.Parse(part.Substring(part.IndexOf('=') + 1));
                                    break;
                                }                               
                            }

                            foreach (string part in parts)
                            {
                                if (part.StartsWith("BYMONTH="))
                                {
                                    slaveRecurrence.MonthOfYear = int.Parse(part.Substring(part.IndexOf('=') + 1));
                                    break;
                                }
                            }

                            //break;
                            
                        //}
                        //else if (slaveRecurrence.RecurrenceType == Outlook.OlRecurrenceType.olRecursMonthly)
                        //{
                            foreach (string part in parts)
                            {
                                if (part.StartsWith("BYMONTHDAY"))
                                {
                                    slaveRecurrence.DayOfMonth = int.Parse(part.Substring(part.IndexOf('=') + 1));
                                    break;
                                }
                            }

                            //foreach (string part in parts)
                            //{
                                //ToDo: Every Second Wednesday comes back as BYDAY=2WE
                                //if (part.StartsWith("BYDAY"))
                                //{
                                //    slave.DayOfMonth = int.Parse(part.Substring(part.IndexOf('=')+1,1));
                                //    switch (part.Substring(part.IndexOf('=') + 2))
                                //    {
                                //        case "MO": slave.DayOfWeekMask = Outlook.OlDaysOfWeek.olMonday; break;
                                //        case "TU": slave.DayOfWeekMask = Outlook.OlDaysOfWeek.olTuesday; break;
                                //        case "WE": slave.DayOfWeekMask = Outlook.OlDaysOfWeek.olWednesday; break;
                                //        case "TH": slave.DayOfWeekMask = Outlook.OlDaysOfWeek.olThursday; break;
                                //        case "FR": slave.DayOfWeekMask = Outlook.OlDaysOfWeek.olFriday; break;
                                //        case "SA": slave.DayOfWeekMask = Outlook.OlDaysOfWeek.olSaturday; break;
                                //        case "SO": slave.DayOfWeekMask = Outlook.OlDaysOfWeek.olSunday; break;


                                //    }
                                //    break ;
                                //}

                            //}

                        //    break;
                        //}

                        break;
                    }

                    

                //if (master.DayOfMonth > 0)
                //    slave.dDayOfMonth = master.DayOfMonth;
                //if (master.MonthOfYear > 0)
                //    slave.MonthOfYear = master.MonthOfYear;
                //if (master.DayOfWeekMask > 0)
                //    slave.DayOfWeekMask = master.DayOfWeekMask;
                //slave.Duration = master.Duration;
                ////if (master.StartTime.Date < new DateTime(1900, 1, 1)
                ////    || master.StartTime.Date > new DateTime(2100, 1, 1))
                //slave.StartTime = master.StartTime;
                ////if (master.EndTime.Date != new DateTime(1601, 1, 2))
                ////if (master.EndTime.Date < new DateTime(1900, 1, 1)
                ////    || master.EndTime.Date > new DateTime(2100, 1, 1))
                //slave.EndTime = master.EndTime;
                //slave.NoEndDate = master.NoEndDate;

                ////slave.Instance = master.Instance;
                //slave.Interval = master.Interval;
                ////slave.Occurrences = master.Occurrences;
                ////if (master.PatternEndDate.Date != new DateTime(4500, 12, 31))
                //if (master.PatternEndDate.Date > new DateTime(1900, 1, 1)
                //     && master.PatternEndDate.Date < new DateTime(2100, 1, 1))
                //    slave.PatternEndDate = master.PatternEndDate;
                ////if (master.PatternStartDate != new DateTime(4501, 1, 1))

                //slave.PatternStartDate = master.PatternStartDate;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }


        }

        private static DateTime GetDateTime(string dateTime)
        {
            string format = "yyyyMMdd";
            if (dateTime.Contains("T"))
                format += "'T'HHmmss";
            if (dateTime.EndsWith("Z"))
                format += "'Z'";
            return DateTime.ParseExact(dateTime, format, new System.Globalization.CultureInfo("en-US"));
        }
    }
}
