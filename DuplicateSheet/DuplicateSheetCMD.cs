using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Tools.DuplicateSheet.ViewModel;

namespace Tools.DuplicateSheet
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class DuplicateSheetCMD : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            Document doc = uiApp.ActiveUIDocument.Document;
            DupSheetViewModel dupSheetViewModel = new DupSheetViewModel(uiApp);
            dupSheetViewModel.DupSheetView.ShowDialog();
            return Result.Succeeded;
        }
    }
}
