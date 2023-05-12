using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Visual;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace Tools
{
    public class Data
    {
        private static Data? _instance;
        private Data()
        {
        }
        public static Data Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new Data();

                return _instance;
            }
        }
        public List<Family> GetAllFamilyInstances(Document doc)
        {
            var res = new FilteredElementCollector(doc).OfClass(typeof(Family)).Cast<Family>().ToList();
            return res;
        }
        public List<string> GetAllFamilyNameInView(
        Document doc)
        {
            return new FilteredElementCollector(doc, doc.ActiveView.Id)
                    .ToElements()
                    .Select<Element, string>(e => e.Category != null ? e.Category.Name : string.Empty)
                    .Distinct<string>().Where(p => p != string.Empty).ToList();
        }
        public ICollection<Element> GetAllElementsInView(Document doc)
        {
             return new FilteredElementCollector(doc,doc.ActiveView.Id).ToElements();             
        }
        public ICollection<Element> GetAllElementsInProject(Document doc)
        {
            return new FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements();
        }
        public List<RoomTagType>? GetAllTagRoom(Document doc)
        {
            try
            {

                List<RoomTagType> result = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol))
                        .Where(p => (p as FamilySymbol).Category.Name == "Room Tags").Cast<RoomTagType>().ToList();
                return result;
            }
            catch (System.Exception e)
            {
                MessageBox.Show(e.Message);
            }
            return null;
        }
        public List<Room> GetRooms(Document doc)
        {
            List<Room> rooms = new FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(BuiltInCategory.OST_Rooms).Cast<Room>().ToList();
            return rooms;
        }
        public List<ViewSheet> GetAllViewSheet(Document doc)
        {
            return new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Sheets).OfType<ViewSheet>().ToList();
        }
        public List<ScheduleSheetInstance> GetAllViewSchedule(Document doc)
        {
            return new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ScheduleGraphics).OfType<ScheduleSheetInstance>().ToList();
        }
    }
    public class DetailItemsFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            if (elem.Category.Name == "Detail Items"|| elem.Category.Name == "Levels")
                return true;
            else return false;
        }
        public bool AllowReference(Reference reference, XYZ position)
        {
            throw new NotImplementedException();
        }
    }
}
