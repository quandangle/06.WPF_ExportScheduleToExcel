using Autodesk.Revit.DB;

namespace QApps
{
    public class ViewScheduleExtension : ViewModelBase
    {
        public string ViewScheduleName { get; set; }
        public ViewSchedule ViewSchedule { get; set; }

        private bool isExport;
        public bool IsExport
        {
            get => isExport;

            set
            {
                isExport = value;
                OnPropertyChanged();
            }
        }

        public ViewScheduleExtension(ViewSchedule viewSchedule)
        {
            ViewSchedule = viewSchedule;
            ViewScheduleName = viewSchedule.Name;
        }
    }
}