using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace QApps
{
    internal static class ViewScheduleUtils
    {
        /// <summary>
        /// Lấy về List ScheduleField có thể export ra file txt
        /// </summary>
        /// <param name="schedule"></param>
        /// <returns></returns>
        internal static IList<ScheduleField> GetCanExportScheduleFields(this ViewSchedule schedule)
        {
            if (schedule == null) return null;

            List<ScheduleField> scheduleFields = new List<ScheduleField>();

            foreach (ScheduleFieldId fieldOrder in schedule.Definition.GetFieldOrder())
            {
                ScheduleField field = schedule.Definition.GetField(fieldOrder);
                if (field.IsHidden) continue;
                scheduleFields.Add(field);
            }
            return scheduleFields;
        }

        /// <summary>
        /// Lấy về tên của các Field có trong bảng thống kê viewSchedule
        /// </summary>
        /// <param name="viewSchedule"></param>
        /// <returns></returns>
        internal static List<string> GetNameOfFields(this ViewSchedule viewSchedule)
        {
            if (viewSchedule == null) return new List<string>();

            IList<ScheduleField> canExportScheduleFieldList = GetCanExportScheduleFields(viewSchedule);

            List<string> strs = new List<string>();
            foreach (ScheduleField scheduleField in canExportScheduleFieldList)
            {
                strs.Add(scheduleField.GetName());
            }
            return strs;
        }
    }
}
