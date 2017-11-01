using System;
using System.IO;
//using Microsoft.Office.Interop.Outlook;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace GetCalendar
{
    class Program
    {
        static void Main(string[] args)
        {
            GetOutlookCalendar("Alvarez Daniel Calendar.ics");
        }

        // https://msdn.microsoft.com/en-us/library/office/bb647583.aspx
        private static void GetOutlookCalendar(string calendarFileName)
        {
            if (string.IsNullOrEmpty(calendarFileName))
                throw new ArgumentException("calendarFileName",
                "Parameter must contain a value.");

            var app = new Outlook.Application();

            Outlook.Folder calendar = app.Session.GetDefaultFolder(
                Outlook.OlDefaultFolders.olFolderCalendar) as Outlook.Folder;
            Outlook.CalendarSharing exporter = calendar.GetCalendarExporter();

            // Set the properties for the export
            exporter.CalendarDetail = Outlook.OlCalendarDetail.olFullDetails;
            exporter.IncludeAttachments = true;
            exporter.IncludePrivateDetails = true;
            exporter.RestrictToWorkingHours = false;
            exporter.IncludeWholeCalendar = true;

            // Save the calendar to disk
            exporter.SaveAsICal(Path.Combine(Environment.CurrentDirectory, calendarFileName));
        }
    }
}
