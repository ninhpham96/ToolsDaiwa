using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System.Collections.Generic;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows;
using Tools.DuplicateSheet.View;
using vView = Autodesk.Revit.DB.View;
using System;

namespace Tools.DuplicateSheet.ViewModel
{
    public partial class DupSheetViewModel : ObservableObject
    {
        #region properties and field
        private DupSheetView dupSheetView;
        private UIDocument uidoc { get; }
        private Document doc { get; }
        List<ViewSheet> selectedSheets = new List<ViewSheet>();
        public DupSheetView DupSheetView
        {
            get
            {
                if (dupSheetView == null)
                {
                    dupSheetView = new DupSheetView() { DataContext = this };
                }
                return dupSheetView;
            }
            set
            {
                dupSheetView = value;
                OnPropertyChanged();
            }
        }
        [ObservableProperty]
        private int countNumber = 1;
        #endregion
        #region command
        [RelayCommand]
        private void Run()
        {
            if (selectedSheets.Count == 0)
            {
                TaskDialog.Show("Thông báo", "Bạn chưa chon sheet nào để copy.", TaskDialogCommonButtons.Ok, TaskDialogResult.Ok);
                return;
            }
            foreach (var sheet in selectedSheets)
            {
                for (int i = 1; i <= CountNumber; i++)
                {
                    DuplicateSelectedSheet(sheet, i);
                }
            }
            DupSheetView.Close();
        }
        [RelayCommand]
        private void Clickme(ViewSheet vs)
        {
            if (selectedSheets.Contains(vs))
                selectedSheets.Remove(vs);
            else selectedSheets.Add(vs);
        }
        #endregion
        #region constructor
        public DupSheetViewModel(UIApplication uiApp)
        {
            this.uidoc = uiApp.ActiveUIDocument;
            this.doc = uiApp.ActiveUIDocument.Document;
            DupSheetView.lsvDuplicateSheet.ItemsSource = Data.Instance.GetAllViewSheet(doc);
        }
        #endregion
    }
    public partial class DupSheetViewModel
    {
        #region methods
        private ElementId? GetSheetTitleBlock(ElementId id)
        {
            var all_title_block = new FilteredElementCollector(doc).
                WhereElementIsNotElementType().
                OfCategory(BuiltInCategory.OST_TitleBlocks).
                ToElements();
            foreach (Element item in all_title_block)
            {
                if (item.OwnerViewId == id)
                {
                    return item.GetTypeId();
                }
            }
            return null;
        }
        void UpdateSheetName(ViewSheet sourceview, ViewSheet targetview, int i)
        {
            string current_name = sourceview.Name;
            string name = string.Empty;
            int number = 0;
            string new_name = "";
            for (int j = 0; j < current_name.Length; j++)
            {
                if ((int)current_name[j] >= 48 && (int)current_name[j] <= 57)
                {
                    name = current_name.Substring(0, j);
                    number = int.Parse(current_name.Substring(j));
                    break;
                }
            }
            if (name == string.Empty)
            {
                new_name = current_name + "02";
                sourceview.Name = current_name + "01";
            }
            else
            {
                int new_number = number + i;
                if (new_number < 10)
                    new_name = name + 0 + new_number;
                else
                    new_name = name + new_number;
            }
            targetview.Name = new_name;
        }
        void UpdateSheetNumer(ViewSheet sourceview, ViewSheet targetview, int i)
        {
            string current_sheetnumber = sourceview.SheetNumber;
            int index = current_sheetnumber.LastIndexOf("-");
            string fisrtsubname = current_sheetnumber.Substring(0,index);
            string lastsubname = current_sheetnumber.Substring(index + 1);
            int sheetnumber = Int32.Parse(lastsubname)+i;
            string newsheetnumber = fisrtsubname+"-"+sheetnumber;

            //for (int j = 0; j < current_sheetnumber.Length; j++)
            //{
            //    if ((int)current_sheetnumber[j] >= 48 && (int)current_sheetnumber[j] <= 57)
            //    {
            //        sheetnumber = current_sheetnumber.Substring(0, j);
            //        number = int.Parse(current_sheetnumber.Substring(j));
            //        break;
            //    }
            //}
            //if (sheetnumber == string.Empty)
            //{
            //    new_sheetnumber = sheetnumber + "02";
            //    sourceview.Name = sheetnumber + "01";
            //}
            //else
            //{
            //    int new_number = number + i;
            //    if (new_number < 10)
            //        new_sheetnumber = sheetnumber + 0 + new_number;
            //    else
            //        new_sheetnumber = sheetnumber + new_number;
            //}
            targetview.SheetNumber = newsheetnumber;
        }
        void DuplicateSelectedSheet(ViewSheet vs, int i)
        {
            var title_block = GetSheetTitleBlock(vs.Id);
            if (title_block != null)
            {
                using (Transaction tran = new Transaction(doc, "create sheet"))
                {
                    tran.Start();
                    var new_sheet = ViewSheet.Create(doc, title_block);
                    var para = vs.LookupParameter("シート 発行目的").AsValueString();
                    new_sheet.LookupParameter("シート 発行目的").Set(para);
                    UpdateSheetName(vs, new_sheet, i);
                    UpdateSheetNumer(vs, new_sheet, i);
                    if (DupSheetView.ckbSchedules.IsChecked == true)
                        DuplicateSchedules(vs.Id, new_sheet.Id);
                    if (DupSheetView.ckbView.IsChecked == true)
                        DuplicateViews(vs, new_sheet);
                    if (DupSheetView.ckbLines.IsChecked == true)
                        DuplicateLines(vs, new_sheet);
                    if (DupSheetView.ckbClouds.IsChecked == true)
                        DuplicateClouds(vs, new_sheet);
                    if (DupSheetView.ckbLegend.IsChecked == true)
                        Duplicatelegends(vs, new_sheet);
                    if (DupSheetView.ckbImages.IsChecked == true)
                        DuplicateImages(vs, new_sheet);
                    if (DupSheetView.ckbView.IsChecked == true)
                        DuplicateTexts(vs, new_sheet);
                    if (DupSheetView.ckbDimensions.IsChecked == true)
                        DuplicateDimensions(vs, new_sheet);
                    if (DupSheetView.ckbSymbols.IsChecked == true)
                        DuplicateSymbols(vs, new_sheet);
                    if (DupSheetView.ckbDWGs.IsChecked == true)
                        DuplicateDwgs(vs, new_sheet);
                    tran.Commit();
                }
            }
        }
        void Duplicatelegends(ViewSheet sourceview, ViewSheet destinationview)
        {
            ICollection<ElementId> viewports_ids = sourceview.GetAllViewports();
            foreach (ElementId viewport_id in viewports_ids)
            {
                Viewport viewport = doc.GetElement(viewport_id) as Viewport;
                var viewport_type_id = viewport.GetTypeId();
                var viewport_origin = viewport.GetBoxCenter();
                var view_id = viewport.ViewId;
                vView view = doc.GetElement(view_id) as vView;

                if (view.ViewType == ViewType.Legend)
                {
                    vView legend_view = view;
                    Viewport new_viewport = Viewport.Create(doc, destinationview.Id, legend_view.Id, viewport_origin);
                    if (new_viewport != null)
                    {
                        var new_viewport_type_id = new_viewport.GetTypeId();
                        if (viewport_type_id != new_viewport_type_id)
                        {
                            new_viewport.ChangeTypeId(viewport_type_id);
                        }
                    }
                }
            }
        }
        void DuplicateSchedules(ElementId oldsheetID, ElementId newsheetID)
        {
            List<ScheduleSheetInstance> viewSchedule = Data.Instance.GetAllViewSchedule(doc);
            foreach (ScheduleSheetInstance item in viewSchedule)
            {
                if (item.OwnerViewId == oldsheetID)
                {
                    if (!item.IsTitleblockRevisionSchedule)
                    {
                        var origin = item.Point;
                        var scheduleID = item.ScheduleId;
                        if (scheduleID == ElementId.InvalidElementId) continue;
                        ViewSchedule viewschedule = doc.GetElement(scheduleID) as ViewSchedule;
                        var schedule_view_id = viewschedule.Duplicate(ViewDuplicateOption.Duplicate);
                        ScheduleSheetInstance.Create(doc, newsheetID, schedule_view_id, origin);
                    }
                }
            }
        }
        void DuplicateViews(ViewSheet oldsheet, ViewSheet newsheet)
        {
            ICollection<ElementId> viewports_ids = oldsheet.GetAllViewports();
            foreach (ElementId viewport_id in viewports_ids)
            {
                Viewport viewport = doc.GetElement(viewport_id) as Viewport;
                var viewport_type_id = viewport.GetTypeId();
                var viewport_origin = viewport.GetBoxCenter();
                var view_id = viewport.ViewId;
                vView view = doc.GetElement(view_id) as vView;
                vView new_view = null;

                if (view.ViewType == ViewType.Legend)
                    continue;
                else if (view.ViewType == ViewType.ThreeD)
                {
                    var NewViewId = view.Duplicate(ViewDuplicateOption.Duplicate);
                    new_view = doc.GetElement(NewViewId) as vView;
                }
                else
                {
                    if (DupSheetView.ckbOp1.IsChecked == true)
                    {
                        var NewViewId = view.Duplicate(ViewDuplicateOption.Duplicate);
                        new_view = doc.GetElement(NewViewId) as vView;
                    }
                    else if (DupSheetView.ckbOp2.IsChecked == true)
                    {
                        var NewViewId = view.Duplicate(ViewDuplicateOption.WithDetailing);
                        new_view = doc.GetElement(NewViewId) as vView;
                    }
                    else if (DupSheetView.ckbOp3.IsChecked == true)
                    {
                        var NewViewId = view.Duplicate(ViewDuplicateOption.AsDependent);
                        new_view = doc.GetElement(NewViewId) as vView;
                    }

                }
                if (new_view != null)
                {
                    Viewport new_vp = Viewport.Create(doc, newsheet.Id, new_view.Id, viewport_origin);
                    var new_viewport_type_id = new_vp.GetTypeId();
                    if (viewport_type_id != new_viewport_type_id)
                    {
                        new_vp.ChangeTypeId(viewport_type_id);
                    }
                }
            }
        }
        void DuplicateElements(ViewSheet sourceview, ViewSheet destinationview, ICollection<ElementId> elementIds)
        {
            if (elementIds.Count > 0)
            {
                CopyPasteOptions copyPasteOptions = new CopyPasteOptions();
                ElementTransformUtils.CopyElements(sourceview, elementIds, destinationview, null, copyPasteOptions);
            }
        }
        void DuplicateLines(ViewSheet sourceview, ViewSheet destinationview)
        {
            var lines_on_sheet = new FilteredElementCollector(doc, sourceview.Id).OfCategory(BuiltInCategory.OST_Lines).ToElementIds();
            DuplicateElements(sourceview, destinationview, lines_on_sheet);
        }
        void DuplicateClouds(ViewSheet sourceview, ViewSheet destinationview)
        {
            List<BuiltInCategory> categories = new List<BuiltInCategory>() { BuiltInCategory.OST_RevisionClouds, BuiltInCategory.OST_RevisionCloudTags };
            var filter = new ElementMulticategoryFilter(categories);
            var clouds_and_tags_ids = new FilteredElementCollector(doc, sourceview.Id).WherePasses(filter).ToElementIds();
            DuplicateElements(sourceview, destinationview, clouds_and_tags_ids);
        }
        void DuplicateImages(ViewSheet sourceview, ViewSheet destinationview)
        {
            var images_on_sheet = new FilteredElementCollector(doc, sourceview.Id).OfCategory(BuiltInCategory.OST_RasterImages).ToElementIds();
            DuplicateElements(sourceview, destinationview, images_on_sheet);
        }
        void DuplicateTexts(ViewSheet sourceview, ViewSheet destinationview)
        {
            var texts_on_sheet = new FilteredElementCollector(doc, sourceview.Id).OfCategory(BuiltInCategory.OST_TextNotes).ToElementIds();
            DuplicateElements(sourceview, destinationview, texts_on_sheet);
        }
        void DuplicateDimensions(ViewSheet sourceview, ViewSheet destinationview)
        {
            var dims_on_sheet = new FilteredElementCollector(doc, sourceview.Id).OfCategory(BuiltInCategory.OST_Dimensions).ToElementIds();
            DuplicateElements(sourceview, destinationview, dims_on_sheet);
        }
        void DuplicateSymbols(ViewSheet sourceview, ViewSheet destinationview)
        {
            var symbols_on_sheet = new FilteredElementCollector(doc, sourceview.Id).OfCategory(BuiltInCategory.OST_GenericAnnotation).ToElementIds();
            DuplicateElements(sourceview, destinationview, symbols_on_sheet);
        }
        void DuplicateDwgs(ViewSheet sourceview, ViewSheet destinationview)
        {
            var dims_on_sheet = new FilteredElementCollector(doc, sourceview.Id).OfClass(typeof(ImportInstance)).ToElementIds();
            DuplicateElements(sourceview, destinationview, dims_on_sheet);
        }
        #endregion
        #region
        void check(bool b)
        {
            foreach (var item in DupSheetView.lsvDuplicateSheet.Items)
            {
                var container = DupSheetView.lsvDuplicateSheet.ItemContainerGenerator.ContainerFromItem(item) as ListViewItem;
                if (container != null)
                {
                    var checkbox = FindVisualChild<CheckBox>(container);
                    if (checkbox != null)
                    {
                        checkbox.IsChecked = b;
                    }
                }
            }
        }
        private T? FindVisualChild<T>(DependencyObject parent) where T : DependencyObject
        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child != null && child is T)
                    return (T)child;
                else
                {
                    var result = FindVisualChild<T>(child);
                    if (result != null)
                        return result;
                }
            }
            return null;
        }
        #endregion
    }
}
