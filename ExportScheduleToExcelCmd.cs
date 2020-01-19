#region Namespaces

using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Application = Autodesk.Revit.ApplicationServices.Application;
#endregion

namespace QApps
{
    [Transaction(TransactionMode.Manual)]
    public class ExportScheduleToExcelCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData,
            ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            // code

            ExportScheduleToExcelViewModel viewModel =
                      new ExportScheduleToExcelViewModel(uidoc);

            ExportScheduleToExcelWindow window =
                new ExportScheduleToExcelWindow(viewModel);

            if (window.ShowDialog() == false) return Result.Cancelled;

            return Result.Succeeded;

        }
    }
}


