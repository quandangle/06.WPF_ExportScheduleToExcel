#region Namespaces

using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
#endregion

namespace QApps
{
    public class ExportScheduleToExcelViewModel : ViewModelBase
    {
        public Document Doc;
        public UIDocument UiDoc;

        #region Khai báo Binding Properties 

        public string ExportExcelFolderPath
        {
            get { return _exportExcelFolderPath; }
            set
            {
                _exportExcelFolderPath = value;
                OnPropertyChanged();
            }
        }

        public List<ViewScheduleExtension> AllViewSchedules { get; set; } = new List<ViewScheduleExtension>();
        public List<ViewScheduleExtension> SelectedViewSchedules { get; set; } = new List<ViewScheduleExtension>();

        public bool IsMultiWorkbooks { get; set; }
        public bool IsOneWorkbook { get; set; } = true;

        #endregion Khai báo Binding Properties 


        public ExportScheduleToExcelViewModel(UIDocument uidoc)
        {
            // Lưu trữ Data từ Revit
            Doc = uidoc.Document;
            UiDoc = uidoc;

            // Khởi tạo data cho WPF window

            List<ViewSchedule> allViewSchedule =
                new FilteredElementCollector(Doc)
                    .OfCategory(BuiltInCategory.OST_Schedules)
                    .Cast<ViewSchedule>()
                    .Where(vs=>vs.CropBox!=null)
                    .Where(vs => vs.Definition.CategoryId.IntegerValue
                                 != (int)BuiltInCategory.OST_Revisions)
                    .ToList();

            foreach (ViewSchedule v in allViewSchedule)
            {
                ViewScheduleExtension viewExtension
                    = new ViewScheduleExtension(v);

                AllViewSchedules.Add(viewExtension);
            }

            AllViewSchedules.Sort((v1, v2)
                => string.CompareOrdinal(v1.ViewScheduleName, v2.ViewScheduleName));

            ExportExcelFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        }

        private string _exportExcelFolderPath;
    }
}
