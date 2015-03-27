using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Google.Apis.Calendar.v3.Data;

namespace GoContactSyncMod
{
    class Factory
    {

        internal static Google.Apis.Calendar.v3.Data.Event NewEvent()
        {
            Event ev = new Event();
            ev.Reminders = new Event.RemindersData { Overrides = new List<EventReminder>(), UseDefault = false };
            ev.Recurrence = new List<String>();
            ev.ExtendedProperties = new Event.ExtendedPropertiesData { Shared = new Dictionary<String, String>() };
            ev.Start = new EventDateTime();
            ev.End = new EventDateTime();
            ev.Locked = false;
            
            return ev;
        }
    }
}
