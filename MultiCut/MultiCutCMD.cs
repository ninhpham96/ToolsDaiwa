using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Windows;

namespace Tools.MultiCut
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class MultiCutCMD:IExternalCommand
    {
        UIApplication? uiapp;
        UIDocument? uidoc;
        Document? doc;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            uiapp = commandData.Application;
            uidoc = uiapp.ActiveUIDocument;
            doc = uidoc.Document;
            List<Element> fromlistElement = new List<Element>();
            List<Element> tolistElement = new List<Element>();
            MuliCutFilter genericModelFileter = new MuliCutFilter();

            try
            {
                IList<Reference> fromlistReference = uidoc.Selection.PickObjects(ObjectType.Element, genericModelFileter, "Chọn đối tượng bị cắt!");
                IList<Reference> tolistReference = uidoc.Selection.PickObjects(ObjectType.Element, genericModelFileter, "Chọn đối tượng cắt!");
                foreach (var item in fromlistReference)
                {
                    fromlistElement.Add(doc.GetElement(item));
                }
                foreach (var item in tolistReference)
                {
                    tolistElement.Add(doc.GetElement(item));
                }
                foreach (var item in fromlistElement)
                {
                    foreach (var ite in tolistElement)
                    {
                        using (Transaction transaction = new Transaction(doc))
                        {
                            transaction.Start("Try cut via API");
                            SolidSolidCutUtils.AddCutBetweenSolids(doc, item, ite);
                            transaction.Commit();
                        }
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            return Result.Succeeded;
        }
    }
    class MuliCutFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            if (elem.Category.Name == "Generic Models" || elem.Category.Name == "Structural Framing")
                return true;
            else return false;
        }
        public bool AllowReference(Reference reference, XYZ position)
        {
            throw new NotImplementedException();
        }
    }
}
