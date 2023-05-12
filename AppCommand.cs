#region Namespaces
using System;
using System.Collections.Generic;
using Autodesk.Revit.UI;
using System.Reflection;
using TaskDialog = Autodesk.Revit.UI.TaskDialog;
using Tools.Utilities;
using System.Configuration.Assemblies;
using Tools.QuickSelect.ViewModel;
using Tools.QuickSelect;
using Tools.AutoTag;
using Tools.DuplicateSheet;
using Tools.MultiCut;
#endregion

namespace Tools
{
    public class AppCommand : IExternalApplication
    {
        #region properties and fields
        public static QuickSelectHandler? Handler { get; set; } = null;
        public static ExternalEvent? ExEvent { get; set; } = null;
        private static UIControlledApplication? _uiApp;
        internal static string assemblyPath = typeof(AppCommand).Assembly.Location;
        internal static AppCommand? GetInstance { get; private set; } = null;
        #endregion
        #region methods
        public Result OnStartup(UIControlledApplication a)
        {
            GetInstance = this;
            _uiApp = a;
            BuildUI(a);
            Handler = new QuickSelectHandler();
            ExEvent = ExternalEvent.Create(Handler);
            return Result.Succeeded;
        }
        public Result OnShutdown(UIControlledApplication a)
        {
            return Result.Succeeded;
        }
        private void BuildUI(UIControlledApplication uiApp)
        {
            RibbonPanel? panel = RibbonUtils.CreatePanel(uiApp, "Select");
            var data1 = new PushButtonData("btnSelect", "Quick\nSelect", assemblyPath, typeof(CmdQuickSelect).FullName);
            data1.LargeImage = RibbonUtils.ConvertFromBitmap(Properties.Resources.QuickSelect32);
            data1.Image = RibbonUtils.ConvertFromBitmap(Properties.Resources.QuickSelect16);
            data1.ToolTip = "Quick Select Elements In Activeview";

            var data2 = new PushButtonData("btnTagRoom","Auto Tag", assemblyPath, typeof(AutoTagRoomCMD).FullName);
            data2.LargeImage = RibbonUtils.ConvertFromBitmap(Properties.Resources.tagroom32);
            data2.Image = RibbonUtils.ConvertFromBitmap(Properties.Resources.tagroom16);
            data2.ToolTip = "Auto Tag Room";

            var btn3 =RibbonUtils.AddPushButton<DuplicateSheetCMD>(panel, "Duplicate\nSheet");
            btn3.LargeImage = RibbonUtils.ConvertFromBitmap(Properties.Resources.Duplicate32);
            btn3.Image = RibbonUtils.ConvertFromBitmap(Properties.Resources.Duplicate16);
            btn3.ToolTip = "Duplicate multi sheet, rename and renumber sheet.";

            var btn4 = RibbonUtils.AddPushButton<MultiCutCMD>(panel, "Multi Cut");            
            btn4.LargeImage = RibbonUtils.ConvertFromBitmap(Properties.Resources.Multicut32);
            btn4.Image = RibbonUtils.ConvertFromBitmap(Properties.Resources.Multicut16);
            btn4.ToolTip = "Multi cut element beetween genericmodel with structuralframing.";

            if (panel != null)
            {
                var btnColor = panel.AddItem(data1) as PushButton;
                var btnColor1 = panel.AddItem(data2) as PushButton;
            }
        }
        #endregion
    }
}
