using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Tools.AutoTag
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class AutoTagRoomCMD : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            using (TransactionGroup tranG = new TransactionGroup(uidoc.Document, ""))
            {
                tranG.Start();
                AutoTagRoomVM autoTagRoomVM = new AutoTagRoomVM(uidoc);
                tranG.Assimilate();
            }
            return Result.Succeeded;
        }
    }
}
