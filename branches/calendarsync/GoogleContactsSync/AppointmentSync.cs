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

            if (master.IsRecurring)
                ;//ToDo: Update(master.GetRecurrencePattern(), slave.Recurrence);
            else if (slave.Recurrence != null)
                slave.Recurrence = null;
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
            if (master.Recurrence!=null)
                Update(master, master.Recurrence, slave.GetRecurrencePattern());
            else if (slave.IsRecurring)
                slave.ClearRecurrencePattern();
        }

        //public static void Update(Outlook.RecurrencePattern master, Recurrence slave)
        //{
        //    try
        //    {
        //        //slave.RecurrenceType = master.RecurrenceType;
        //        if (master.DayOfMonth > 0)
        //            slave.dDayOfMonth = master.DayOfMonth;
        //        if (master.MonthOfYear > 0)
        //            slave.MonthOfYear = master.MonthOfYear;
        //        if (master.DayOfWeekMask > 0)
        //            slave.DayOfWeekMask = master.DayOfWeekMask;
        //        slave.Duration = master.Duration;
        //        //if (master.StartTime.Date < new DateTime(1900, 1, 1)
        //        //    || master.StartTime.Date > new DateTime(2100, 1, 1))
        //            slave.StartTime = master.StartTime;
        //        //if (master.EndTime.Date != new DateTime(1601, 1, 2))
        //        //if (master.EndTime.Date < new DateTime(1900, 1, 1)
        //        //    || master.EndTime.Date > new DateTime(2100, 1, 1))
        //            slave.EndTime = master.EndTime;
        //        slave.NoEndDate = master.NoEndDate;

        //        //slave.Instance = master.Instance;
        //        slave.Interval = master.Interval;
        //        //slave.Occurrences = master.Occurrences;
        //        //if (master.PatternEndDate.Date != new DateTime(4500, 12, 31))
        //        if (master.PatternEndDate.Date > new DateTime(1900, 1, 1)
        //             && master.PatternEndDate.Date < new DateTime(2100, 1, 1))
        //            slave.PatternEndDate = master.PatternEndDate;
        //        //if (master.PatternStartDate != new DateTime(4501, 1, 1))

        //            slave.PatternStartDate = master.PatternStartDate;
        //    }
        //    catch (Exception ex)
        //    {
        //        throw ex;
        //    }


        //}

        public static void Update(EventEntry googleAppointment, Recurrence master, Outlook.RecurrencePattern slave)
        {
            try
            {           
                string[] patterns = master.Value.Split(new char[] {'\r','\n'}, StringSplitOptions.RemoveEmptyEntries);
                foreach (string pattern in patterns)
                {
                    if (pattern.StartsWith("DTSTART")) 
                    {
                        string[] parts = pattern.Split(new char[] {';',':'});
                        if (parts.Length >= 2)
                        {
                            string format = "yyyyMMdd";
                            if (parts[2].Contains("T"))
                                format += "'T'HHmmss";
                            slave.StartTime = DateTime.ParseExact(parts[2], format, new System.Globalization.CultureInfo("en-US"));
                        }
                        break;
                    }
                }

                foreach (string pattern in patterns)
                {
                    if (pattern.StartsWith("DTEND"))
                    {
                        string[] parts = pattern.Split(new char[] { ';', ':' });
                        if (parts.Length >= 2)
                        {
                            string format = "yyyyMMdd";
                            if (parts[2].Contains("T"))
                                format += "'T'HHmmss";
                            slave.EndTime = DateTime.ParseExact(parts[2], format, new System.Globalization.CultureInfo("en-US"));                            
                        }
                        break;
                    }
                }

                foreach (string pattern in patterns)
                {
                    if (pattern.StartsWith("RRULE"))
                    {
                        string[] parts = pattern.Split(new char[] { ';', ':' });

                        foreach (string part in parts)
                        {
                            if (part.StartsWith("FREQ"))
                            {
                                switch (part.Substring(part.IndexOf('=') + 1))
                                {
                                    case "DAILY": slave.RecurrenceType = Outlook.OlRecurrenceType.olRecursDaily; break;
                                    case "WEEKLY": slave.RecurrenceType = Outlook.OlRecurrenceType.olRecursWeekly; break;
                                    case "MONTHLY": slave.RecurrenceType = Outlook.OlRecurrenceType.olRecursMonthly; break;
                                    case "YEARLY": slave.RecurrenceType = Outlook.OlRecurrenceType.olRecursYearly; break;

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
                                if (googleAppointment.Times.Count > 0)
                                    slave.PatternStartDate = googleAppointment.Times[0].StartTime;
                                slave.Occurrences = int.Parse(part.Substring(part.IndexOf('=') + 1));
                                break;
                            }
                        }                        

                        if (slave.RecurrenceType == Outlook.OlRecurrenceType.olRecursWeekly)
                        {
                            foreach (string part in parts)
                            {
                                if (part.StartsWith("BYDAY"))
                                {
                                    switch (part.Substring(part.IndexOf('=') + 1))
                                    {
                                        case "MO": slave.DayOfWeekMask = slave.DayOfWeekMask | Outlook.OlDaysOfWeek.olMonday; break;
                                        case "TU": slave.DayOfWeekMask = slave.DayOfWeekMask | Outlook.OlDaysOfWeek.olTuesday; break;
                                        case "WE": slave.DayOfWeekMask = slave.DayOfWeekMask | Outlook.OlDaysOfWeek.olWednesday; break;
                                        case "TH": slave.DayOfWeekMask = slave.DayOfWeekMask | Outlook.OlDaysOfWeek.olThursday; break;
                                        case "FR": slave.DayOfWeekMask = slave.DayOfWeekMask | Outlook.OlDaysOfWeek.olFriday; break;
                                        case "SA": slave.DayOfWeekMask = slave.DayOfWeekMask | Outlook.OlDaysOfWeek.olSaturday; break;
                                        case "SO": slave.DayOfWeekMask = slave.DayOfWeekMask | Outlook.OlDaysOfWeek.olSunday; break;


                                    }
                                    //Don't break because multiple days;
                                }

                            }

                            break;
                        }
                        else if (slave.RecurrenceType == Outlook.OlRecurrenceType.olRecursYearly)
                        {
                            foreach (string part in parts)
                            {
                                if (part.StartsWith("BYMONTHDAY"))
                                {
                                    slave.DayOfMonth = int.Parse(part.Substring(part.IndexOf('=') + 1));
                                    break;
                                }                               
                            }

                            foreach (string part in parts)
                            {
                                if (part.StartsWith("BYMONTH="))
                                {
                                    slave.MonthOfYear = int.Parse(part.Substring(part.IndexOf('=') + 1));
                                    break;
                                }
                            }
                            
                        }
                        else if (slave.RecurrenceType == Outlook.OlRecurrenceType.olRecursMonthly)
                        {
                            foreach (string part in parts)
                            {
                                if (part.StartsWith("BYMONTHDAY"))
                                {
                                    slave.DayOfMonth = int.Parse(part.Substring(part.IndexOf('=') + 1));
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

                            break;
                        }

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
    }
}
