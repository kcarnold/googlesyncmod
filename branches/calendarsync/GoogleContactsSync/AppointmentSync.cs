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

            //if (master.IsRecurring)
            //    Update(master.GetRecurrencePattern(), slave.Recurrence);
            //else if (slave.Recurrence != null)
            //    slave.Recurrence = null;
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

            //slave.RequiredAttendees = master.Participants;
            //slave.OptionalAttendees = master.OptionalAttendees;

            slave.ReminderSet = false;
            if (master.Reminder != null && !master.Reminder.Method.Equals(Google.GData.Extensions.Reminder.ReminderMethod.none) && master.Reminder.AbsoluteTime >= DateTime.Now)
            { 
                slave.ReminderSet = true;
                slave.ReminderMinutesBeforeStart = master.Reminder.Minutes;
            }

            //slave.Resources = master.Resources;
            //slave.RTFBody = master.RTFBody;

            //if (master.IsRecurring)
            //    Update(master.GetRecurrencePattern(), slave.GetRecurrencePattern());
            //else if (slave.IsRecurring)
            //    slave.ClearRecurrencePattern();
        }       

        //ToDo: Implement recurrent Patterns
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
    }
}
