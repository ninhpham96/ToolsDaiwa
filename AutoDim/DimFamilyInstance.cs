using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace Tools
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class DimFamilyInstance : IExternalCommand
    {
        UIApplication uiapp;
        UIDocument uidoc;
        Document doc;
        List<FamilyInstance> instances;
        ReferenceArray refs;
        XYZ dir = null;
        Line line = null;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            uiapp = commandData.Application;
            uidoc = uiapp.ActiveUIDocument;
            doc = uidoc.Document;
            var viewScale = doc.ActiveView.Scale;
            instances = new List<FamilyInstance>();
            refs = new ReferenceArray();

            var pickobs = uidoc.Selection.PickElementsByRectangle(new DetailItemsFilter(),"Picks elements");
            var point1 = uidoc.Selection.PickPoint();
            var point2 = uidoc.Selection.PickPoint();
            foreach (Element ele in pickobs)
            {
                var e = ele as FamilyInstance;
                if ( e != null)
                {
                    refs.Append(e.GetReferenceByName("Center (Front/Back)"));
                    continue;
                }
                var lv = ele as Level;
                if (lv != null)
                {
                    refs.Append(lv.GetPlaneReference());
                    dir = new XYZ(1,0,0).CrossProduct(point1-point2);
                }
            }
            if (doc.ActiveView.ViewDirection.X == 1)
            {
                dir = new XYZ(1, 0, 0).CrossProduct(point1 - point2);
                line = Line.CreateUnbound(point1, point1 + dir * 100);
            }
            else if (doc.ActiveView.ViewDirection.Y == 1)
            {

            }

            using(Transaction tran = new Transaction(doc,"create dim"))
            {
                tran.Start();
                doc.Create.NewDimension(doc.ActiveView, line, refs);
                tran.Commit();
            }
            return Result.Succeeded;
        }
    }
}
