using System;
using System.Collections.Generic;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace GoContactSyncMod
{

    internal static class AppointmentSync
    {

        //internal static DateTime outlookDateMin = new DateTime(4501, 1, 1);
        //internal static DateTime outlookDateMax = new DateTime(4500, 12, 31);

        /// <summary>
        /// Updates Outlook appointments (calendar)
        /// </summary>
        public static void Update(Outlook.AppointmentItem master, Outlook.AppointmentItem slave)
        {            
            slave.Subject = master.Subject;
            slave.AllDayEvent = master.AllDayEvent;

            //foreach (Outlook.Attachment attachment in master.Attachments)
            //    slave.Attachments.Add(master.Attachments);
            
            slave.Body = master.Body;
            slave.BusyStatus = master.BusyStatus;
            slave.Categories = master.Categories;
            slave.Duration = master.Duration;
    
            slave.Start = master.Start;
            //slave.StartInStartTimeZone = master.StartInStartTimeZone;
            //slave.StartTimeZone = master.StartTimeZone;
            //slave.StartUTC = master.StartUTC;
            slave.End = master.End;
            slave.RequiredAttendees = master.RequiredAttendees;
            slave.OptionalAttendees = master.OptionalAttendees;
            slave.ReminderMinutesBeforeStart = master.ReminderMinutesBeforeStart;
            
            slave.Resources = master.Resources;
            slave.RTFBody = master.RTFBody;

            if (master.IsRecurring)
                Update(master.GetRecurrencePattern(), slave.GetRecurrencePattern());
            else if (slave.IsRecurring)
                slave.ClearRecurrencePattern();
        }       
        public static void Update(Outlook.RecurrencePattern master, Outlook.RecurrencePattern slave)
        {
            try
            {
                slave.RecurrenceType = master.RecurrenceType;
                if (master.DayOfMonth > 0)
                    slave.DayOfMonth = master.DayOfMonth;
                if (master.MonthOfYear > 0)
                    slave.MonthOfYear = master.MonthOfYear;
                if (master.DayOfWeekMask > 0)
                    slave.DayOfWeekMask = master.DayOfWeekMask;
                slave.Duration = master.Duration;
                //if (master.StartTime.Date < new DateTime(1900, 1, 1)
                //    || master.StartTime.Date > new DateTime(2100, 1, 1))
                    slave.StartTime = master.StartTime;
                //if (master.EndTime.Date != new DateTime(1601, 1, 2))
                //if (master.EndTime.Date < new DateTime(1900, 1, 1)
                //    || master.EndTime.Date > new DateTime(2100, 1, 1))
                    slave.EndTime = master.EndTime;
                slave.NoEndDate = master.NoEndDate;

                //slave.Instance = master.Instance;
                slave.Interval = master.Interval;
                //slave.Occurrences = master.Occurrences;
                //if (master.PatternEndDate.Date != new DateTime(4500, 12, 31))
                if (master.PatternEndDate.Date > new DateTime(1900, 1, 1)
                     && master.PatternEndDate.Date < new DateTime(2100, 1, 1))
                    slave.PatternEndDate = master.PatternEndDate;
                //if (master.PatternStartDate != new DateTime(4501, 1, 1))

                    slave.PatternStartDate = master.PatternStartDate;
            }
            catch (Exception ex)
            {
                throw ex;
            }
   
            
        }
    }
}
