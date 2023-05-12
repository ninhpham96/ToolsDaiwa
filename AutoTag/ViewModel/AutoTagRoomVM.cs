using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using Tools.AutoTag.View;

namespace Tools
{
    public partial class AutoTagRoomVM : ObservableObject
    {
        #region field
        private AutoTagRoomView _view;
        public AutoTagRoomView view
        {
            get
            {
                if (_view == null)
                    _view = new AutoTagRoomView() { DataContext = this };
                return _view;
            }
            set
            {
                _view = value;
                OnPropertyChanged();
            }
        }
        private UIDocument uidoc;
        private Document doc;
        List<Room> rooms;
        List<RoomTagType> source;
        [ObservableProperty]
        public ObservableCollection<RoomTagType> srcroomName = new ObservableCollection<RoomTagType>();
        [ObservableProperty]
        private ObservableCollection<RoomTagType> srcroomWall = new ObservableCollection<RoomTagType>();
        [ObservableProperty]
        private ObservableCollection<RoomTagType> srcroomFloor = new ObservableCollection<RoomTagType>();
        [ObservableProperty]
        private ObservableCollection<RoomTagType> srcroomCeiling = new ObservableCollection<RoomTagType>();
        [ObservableProperty]
        private ObservableCollection<RoomTagType> srcroomHabaki = new ObservableCollection<RoomTagType>();
        [ObservableProperty]
        private ObservableCollection<RoomTagType> srcroomCeilingConnect = new ObservableCollection<RoomTagType>();
        #endregion
        #region command
        [RelayCommand]
        void Run()
        {
            try
            {
                foreach (Room room in rooms)
                {
                    if (room == null)
                        return;
                    else
                    {
                        TagRoom(room);
                    }
                }
                view.Close();
                return;
            }
            catch (System.Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }
        [RelayCommand]
        void Check()
        {
            ClearSource();
            if (view.rbtNotTruss.IsChecked == true)
            {
                SetItemSource(source, true);
            }
            if (view.rbtTruss.IsChecked == true)
            {
                SetItemSource(source, false);
            }
            SetSelectedIndex();
        }
        #endregion
        #region constructor
        public AutoTagRoomVM(UIDocument uidoc)
        {
            this.uidoc = uidoc;
            doc = uidoc.Document;
            if (doc.ActiveView.ViewType != ViewType.Section)
            {
                TaskDialog.Show("Cảnh báo", "Chuyển qua ViewSection");
                return;
            }
            else if (doc.ActiveView.Scale != 50 && doc.ActiveView.Scale != 60)
            {
                TaskDialog.Show("Cảnh báo", "Cần chuyển tỉ lệ về 1:50 hoặc 1:60");
                return;
            }
            rooms = Data.Instance.GetRooms(doc);
            source = Data.Instance.GetAllTagRoom(uidoc.Document);
            if (true)
            {
                SetItemSource(source, true);
                SetSelectedIndex();
            }
            view.ShowDialog();
        }
        #endregion
        #region methods
        void TagRoom(Room room)
        {
            if (doc.ActiveView.ViewDirection.Y == 1 || doc.ActiveView.ViewDirection.Y == -1)
            {
                if (doc.ActiveView.Scale == 50)
                {
                    using (Transaction tran = new Transaction(doc, "create tag"))
                    {
                        tran.Start();
                        var loca = room.get_BoundingBox(doc.ActiveView);
                        //create tag room name
                        RoomTag roomTag = doc.Create.NewRoomTag(new LinkElementId(room.Id), new UV(0, 0), doc.ActiveView.Id);
                        roomTag.ChangeTypeId((view.cbtagRoom.SelectedItem as RoomTagType).Id);
                        roomTag.HasLeader = true;
                        roomTag.TagHeadPosition = roomTag.LeaderEnd;
                        roomTag.HasLeader = false;

                        //CreateDuct tag room wall
                        RoomTag roomTag1 = doc.Create.NewRoomTag(new LinkElementId(room.Id), new UV(0, 0), doc.ActiveView.Id);
                        roomTag1.ChangeTypeId((view.cbtagWall.SelectedItem as RoomTagType).Id);
                        roomTag1.HasLeader = true;
                        roomTag1.LeaderEnd = new XYZ(roomTag1.LeaderEnd.X - 2, roomTag1.LeaderEnd.Y, roomTag1.LeaderEnd.Z + 1);
                        roomTag1.HasLeader = false;

                        //create tag ceiling
                        RoomTag roomTag2 = doc.Create.NewRoomTag(new LinkElementId(room.Id), new UV(0, 0), doc.ActiveView.Id);
                        roomTag2.ChangeTypeId((view.cbtagCeil.SelectedItem as RoomTagType).Id);
                        roomTag2.HasLeader = true;
                        roomTag2.LeaderEnd = new XYZ(roomTag2.LeaderEnd.X, roomTag2.LeaderEnd.Y, loca.Max.Z);
                        roomTag2.TagHeadPosition = new XYZ(roomTag2.LeaderEnd.X - 2, roomTag2.LeaderEnd.Y, loca.Max.Z);
                        roomTag2.LeaderElbow = new XYZ(roomTag2.LeaderEnd.X - 1, roomTag2.LeaderEnd.Y, loca.Max.Z - 0.7139017157558);

                        //create tag floor
                        RoomTag roomTag3 = doc.Create.NewRoomTag(new LinkElementId(room.Id), new UV(0, 0), doc.ActiveView.Id);
                        roomTag3.ChangeTypeId((view.cbtagFloor.SelectedItem as RoomTagType).Id);
                        roomTag3.HasLeader = true;
                        roomTag3.LeaderEnd = new XYZ(roomTag3.LeaderEnd.X - 2, roomTag3.LeaderEnd.Y, loca.Min.Z);
                        roomTag3.TagHeadPosition = new XYZ(roomTag3.LeaderEnd.X + 1, roomTag3.LeaderEnd.Y, loca.Min.Z + 3);
                        roomTag3.LeaderElbow = new XYZ(roomTag3.LeaderEnd.X + 1, roomTag3.LeaderEnd.Y, loca.Min.Z + 0.764648401984138);

                        //create tag habaki
                        RoomTag roomTag4 = doc.Create.NewRoomTag(new LinkElementId(room.Id), new UV(0, 0), doc.ActiveView.Id);
                        roomTag4.ChangeTypeId((view.cbtagHabaki.SelectedItem as RoomTagType).Id);
                        roomTag4.HasLeader = true;
                        roomTag4.LeaderEnd = new XYZ(loca.Min.X, roomTag4.LeaderEnd.Y, loca.Min.Z);
                        roomTag4.TagHeadPosition = new XYZ(loca.Min.X + 0.5, roomTag4.LeaderEnd.Y, loca.Min.Z + 3);
                        roomTag4.LeaderElbow = new XYZ(loca.Min.X + 0.5, roomTag4.LeaderEnd.Y, loca.Min.Z + 1.069843382038858);

                        RoomTag roomTag5 = doc.Create.NewRoomTag(new LinkElementId(room.Id), new UV(0, 0), doc.ActiveView.Id);
                        roomTag5.ChangeTypeId((view.cbtagHabaki.SelectedItem as RoomTagType).Id);
                        roomTag5.HasLeader = true;
                        roomTag5.LeaderEnd = new XYZ(loca.Max.X, roomTag5.LeaderEnd.Y, loca.Min.Z);
                        roomTag5.TagHeadPosition = new XYZ(loca.Max.X - 6, roomTag5.LeaderEnd.Y, loca.Min.Z + 3);
                        roomTag5.LeaderElbow = new XYZ(loca.Max.X - 1, roomTag5.LeaderEnd.Y, loca.Min.Z + 1.069843382038858);


                        //create tag ...
                        RoomTag roomTag6 = doc.Create.NewRoomTag(new LinkElementId(room.Id), new UV(0, 0), doc.ActiveView.Id);
                        roomTag6.ChangeTypeId((view.cbtagMawari.SelectedItem as RoomTagType).Id);
                        roomTag6.HasLeader = true;
                        roomTag6.LeaderEnd = new XYZ(loca.Min.X, roomTag6.LeaderEnd.Y, loca.Max.Z);
                        roomTag6.TagHeadPosition = new XYZ(loca.Min.X, roomTag6.LeaderEnd.Y, loca.Max.Z + 1.3);
                        roomTag6.LeaderElbow = new XYZ(loca.Min.X + 0.5, roomTag6.LeaderEnd.Y, loca.Max.Z - 0.7259865516096);

                        RoomTag roomTag7 = doc.Create.NewRoomTag(new LinkElementId(room.Id), new UV(0, 0), doc.ActiveView.Id);
                        roomTag7.ChangeTypeId((view.cbtagMawari.SelectedItem as RoomTagType).Id);
                        roomTag7.HasLeader = true;
                        roomTag7.LeaderEnd = new XYZ(loca.Max.X, roomTag7.LeaderEnd.Y, loca.Max.Z);
                        roomTag7.TagHeadPosition = new XYZ(loca.Max.X - 5, roomTag7.LeaderEnd.Y, loca.Max.Z + 1.3);
                        roomTag7.LeaderElbow = new XYZ(loca.Max.X - 0.5, roomTag7.LeaderEnd.Y, loca.Max.Z - 0.7259865516096);

                        view.Close();
                        tran.Commit();
                    }
                }
                else if (doc.ActiveView.Scale == 60)
                {
                    using (Transaction tran = new Transaction(doc, "create tag"))
                    {
                        tran.Start();
                        var loca = room.get_BoundingBox(doc.ActiveView);
                        //create tag room name
                        RoomTag roomTag = doc.Create.NewRoomTag(new LinkElementId(room.Id), new UV(0, 0), doc.ActiveView.Id);
                        roomTag.ChangeTypeId((view.cbtagRoom.SelectedItem as RoomTagType).Id);
                        roomTag.HasLeader = true;
                        roomTag.TagHeadPosition = roomTag.LeaderEnd;
                        roomTag.HasLeader = false;

                        //CreateDuct tag room wall
                        RoomTag roomTag1 = doc.Create.NewRoomTag(new LinkElementId(room.Id), new UV(0, 0), doc.ActiveView.Id);
                        roomTag1.ChangeTypeId((view.cbtagWall.SelectedItem as RoomTagType).Id);
                        roomTag1.HasLeader = true;
                        roomTag1.LeaderEnd = new XYZ(roomTag1.LeaderEnd.X - 2, roomTag1.LeaderEnd.Y, roomTag1.LeaderEnd.Z + 1);
                        roomTag1.HasLeader = false;

                        //create tag ceiling
                        RoomTag roomTag2 = doc.Create.NewRoomTag(new LinkElementId(room.Id), new UV(0, 0), doc.ActiveView.Id);
                        roomTag2.ChangeTypeId((view.cbtagCeil.SelectedItem as RoomTagType).Id);
                        roomTag2.HasLeader = true;
                        roomTag2.LeaderEnd = new XYZ(roomTag2.LeaderEnd.X, roomTag2.LeaderEnd.Y, loca.Max.Z);
                        roomTag2.TagHeadPosition = new XYZ(roomTag2.LeaderEnd.X - 2, roomTag2.LeaderEnd.Y, loca.Max.Z);
                        roomTag2.LeaderElbow = new XYZ(roomTag2.LeaderEnd.X - 0.5, roomTag2.LeaderEnd.Y, loca.Max.Z - 0.85668205890687);

                        //create tag floor
                        RoomTag roomTag3 = doc.Create.NewRoomTag(new LinkElementId(room.Id), new UV(0, 0), doc.ActiveView.Id);
                        roomTag3.ChangeTypeId((view.cbtagFloor.SelectedItem as RoomTagType).Id);
                        roomTag3.HasLeader = true;
                        roomTag3.LeaderEnd = new XYZ(roomTag3.LeaderEnd.X - 2, roomTag3.LeaderEnd.Y, loca.Min.Z);
                        roomTag3.TagHeadPosition = new XYZ(roomTag3.LeaderEnd.X + 1, roomTag3.LeaderEnd.Y, loca.Min.Z + 3.5);
                        roomTag3.LeaderElbow = new XYZ(roomTag3.LeaderEnd.X + 1, roomTag3.LeaderEnd.Y, loca.Min.Z + 0.81757808238097);

                        //create tag habaki
                        RoomTag roomTag4 = doc.Create.NewRoomTag(new LinkElementId(room.Id), new UV(0, 0), doc.ActiveView.Id);
                        roomTag4.ChangeTypeId((view.cbtagHabaki.SelectedItem as RoomTagType).Id);
                        roomTag4.HasLeader = true;
                        roomTag4.LeaderEnd = new XYZ(loca.Min.X, roomTag4.LeaderEnd.Y, loca.Min.Z);
                        roomTag4.TagHeadPosition = new XYZ(loca.Min.X + 0.5, roomTag4.LeaderEnd.Y, loca.Min.Z + 3);
                        roomTag4.LeaderElbow = new XYZ(loca.Min.X + 0.5, roomTag4.LeaderEnd.Y, loca.Min.Z + 0.683812058446638);

                        RoomTag roomTag5 = doc.Create.NewRoomTag(new LinkElementId(room.Id), new UV(0, 0), doc.ActiveView.Id);
                        roomTag5.ChangeTypeId((view.cbtagHabaki.SelectedItem as RoomTagType).Id);
                        roomTag5.HasLeader = true;
                        roomTag5.LeaderEnd = new XYZ(loca.Max.X, roomTag5.LeaderEnd.Y, loca.Min.Z);
                        roomTag5.TagHeadPosition = new XYZ(loca.Max.X - 6, roomTag5.LeaderEnd.Y, loca.Min.Z + 3);
                        roomTag5.LeaderElbow = new XYZ(loca.Max.X - 1, roomTag5.LeaderEnd.Y, loca.Min.Z + 0.683812058446638);


                        //create tag ...
                        RoomTag roomTag6 = doc.Create.NewRoomTag(new LinkElementId(room.Id), new UV(0, 0), doc.ActiveView.Id);
                        roomTag6.ChangeTypeId((view.cbtagMawari.SelectedItem as RoomTagType).Id);
                        roomTag6.HasLeader = true;
                        roomTag6.LeaderEnd = new XYZ(loca.Min.X, roomTag6.LeaderEnd.Y, loca.Max.Z);
                        roomTag6.TagHeadPosition = new XYZ(loca.Min.X, roomTag6.LeaderEnd.Y, loca.Max.Z + 1.3);
                        roomTag6.LeaderElbow = new XYZ(loca.Min.X + 0.5, roomTag6.LeaderEnd.Y, loca.Max.Z - 1.1311838619316);

                        RoomTag roomTag7 = doc.Create.NewRoomTag(new LinkElementId(room.Id), new UV(0, 0), doc.ActiveView.Id);
                        roomTag7.ChangeTypeId((view.cbtagMawari.SelectedItem as RoomTagType).Id);
                        roomTag7.HasLeader = true;
                        roomTag7.LeaderEnd = new XYZ(loca.Max.X, roomTag7.LeaderEnd.Y, loca.Max.Z);
                        roomTag7.TagHeadPosition = new XYZ(loca.Max.X - 5, roomTag7.LeaderEnd.Y, loca.Max.Z + 1.3);
                        roomTag7.LeaderElbow = new XYZ(loca.Max.X - 0.5, roomTag7.LeaderEnd.Y, loca.Max.Z - 1.1311838619316);

                        view.Close();
                        tran.Commit();
                    }
                }
            }
            else if (doc.ActiveView.ViewDirection.X == 1 || doc.ActiveView.ViewDirection.X == -1 && doc.ActiveView.Scale == 50)
            {
                if (doc.ActiveView.Scale == 50)
                {
                    using (Transaction tran = new Transaction(doc, "create tag"))
                    {
                        tran.Start();
                        var loca = room.get_BoundingBox(doc.ActiveView);
                        #region
                        //create tag room name
                        RoomTag roomTag = doc.Create.NewRoomTag(new LinkElementId(room.Id), new UV(0, 0), doc.ActiveView.Id);
                        roomTag.ChangeTypeId((view.cbtagRoom.SelectedItem as RoomTagType).Id);
                        roomTag.HasLeader = true;
                        roomTag.TagHeadPosition = roomTag.LeaderEnd;
                        roomTag.HasLeader = false;

                        //Create tag room wall
                        RoomTag roomTag1 = doc.Create.NewRoomTag(new LinkElementId(room.Id), new UV(0, 0), doc.ActiveView.Id);
                        roomTag1.ChangeTypeId((view.cbtagWall.SelectedItem as RoomTagType).Id);
                        roomTag1.HasLeader = true;
                        roomTag1.LeaderEnd = new XYZ(roomTag1.LeaderEnd.X - 2, roomTag1.LeaderEnd.Y, roomTag1.LeaderEnd.Z + 1);
                        roomTag1.HasLeader = false;

                        //create tag ceiling
                        RoomTag roomTag2 = doc.Create.NewRoomTag(new LinkElementId(room.Id), new UV(0, 0), doc.ActiveView.Id);
                        roomTag2.ChangeTypeId((view.cbtagCeil.SelectedItem as RoomTagType).Id);
                        roomTag2.HasLeader = true;
                        roomTag2.LeaderEnd = new XYZ(roomTag2.LeaderEnd.X, roomTag2.LeaderEnd.Y, loca.Max.Z);
                        roomTag2.TagHeadPosition = new XYZ(roomTag2.LeaderEnd.X, roomTag2.LeaderEnd.Y - 2, loca.Max.Z);
                        roomTag2.LeaderElbow = new XYZ(roomTag2.LeaderEnd.X, roomTag2.LeaderEnd.Y - 0.5, loca.Max.Z - 0.7139017157558);

                        //create tag floor
                        RoomTag roomTag3 = doc.Create.NewRoomTag(new LinkElementId(room.Id), new UV(0, 0), doc.ActiveView.Id);
                        roomTag3.ChangeTypeId((view.cbtagFloor.SelectedItem as RoomTagType).Id);
                        roomTag3.HasLeader = true;
                        roomTag3.LeaderEnd = new XYZ(roomTag3.LeaderEnd.X, roomTag3.LeaderEnd.Y, loca.Min.Z);
                        roomTag3.TagHeadPosition = new XYZ(roomTag3.LeaderEnd.X, roomTag3.LeaderEnd.Y, loca.Min.Z + 3);
                        roomTag3.LeaderElbow = new XYZ(roomTag3.LeaderEnd.X, roomTag3.LeaderEnd.Y + 0.5, loca.Min.Z + 0.764648401984138);
                        #endregion
                        //create tag habaki
                        RoomTag roomTag4 = doc.Create.NewRoomTag(new LinkElementId(room.Id), new UV(0, 0), doc.ActiveView.Id);
                        roomTag4.ChangeTypeId((view.cbtagHabaki.SelectedItem as RoomTagType).Id);
                        roomTag4.HasLeader = true;
                        roomTag4.LeaderEnd = new XYZ(loca.Min.X, loca.Min.Y, loca.Min.Z);
                        roomTag4.TagHeadPosition = new XYZ(loca.Min.X, loca.Min.Y, loca.Min.Z + 3);
                        roomTag4.LeaderElbow = new XYZ(loca.Min.X, loca.Min.Y + 0.5, loca.Min.Z + 1.069843382038858);

                        RoomTag roomTag5 = doc.Create.NewRoomTag(new LinkElementId(room.Id), new UV(0, 0), doc.ActiveView.Id);
                        roomTag5.ChangeTypeId((view.cbtagHabaki.SelectedItem as RoomTagType).Id);
                        roomTag5.HasLeader = true;
                        roomTag5.LeaderEnd = new XYZ(loca.Max.X, loca.Max.Y, loca.Min.Z);
                        roomTag5.TagHeadPosition = new XYZ(loca.Max.X, loca.Max.Y - 4, loca.Min.Z + 3);
                        roomTag5.LeaderElbow = new XYZ(loca.Max.X, loca.Max.Y - 0.5, loca.Min.Z + 1.069843382038858);


                        //create tag ...
                        RoomTag roomTag6 = doc.Create.NewRoomTag(new LinkElementId(room.Id), new UV(0, 0), doc.ActiveView.Id);
                        roomTag6.ChangeTypeId((view.cbtagMawari.SelectedItem as RoomTagType).Id);
                        roomTag6.HasLeader = true;
                        roomTag6.LeaderEnd = new XYZ(loca.Min.X, loca.Max.Y, loca.Max.Z);
                        roomTag6.TagHeadPosition = new XYZ(loca.Min.X, loca.Max.Y - 4, loca.Max.Z + 1.3);
                        roomTag6.LeaderElbow = new XYZ(loca.Min.X, loca.Max.Y - 0.5, loca.Max.Z - 0.7259865516096);

                        RoomTag roomTag7 = doc.Create.NewRoomTag(new LinkElementId(room.Id), new UV(0, 0), doc.ActiveView.Id);
                        roomTag7.ChangeTypeId((view.cbtagMawari.SelectedItem as RoomTagType).Id);
                        roomTag7.HasLeader = true;
                        roomTag7.LeaderEnd = new XYZ(loca.Max.X, loca.Min.Y, loca.Max.Z);
                        roomTag7.TagHeadPosition = new XYZ(loca.Max.X - 5, loca.Min.Y, loca.Max.Z + 1.3);
                        roomTag7.LeaderElbow = new XYZ(loca.Max.X - 0.5, loca.Min.Y + 0.5, loca.Max.Z - 0.7259865516096);

                        view.Close();
                        tran.Commit();
                    }
                }
                else if (doc.ActiveView.Scale == 60)
                {
                    using (Transaction tran = new Transaction(doc, "create tag"))
                    {
                        tran.Start();
                        var loca = room.get_BoundingBox(doc.ActiveView);
                        #region
                        //create tag room name
                        RoomTag roomTag = doc.Create.NewRoomTag(new LinkElementId(room.Id), new UV(0, 0), doc.ActiveView.Id);
                        roomTag.ChangeTypeId((view.cbtagRoom.SelectedItem as RoomTagType).Id);
                        roomTag.HasLeader = true;
                        roomTag.TagHeadPosition = roomTag.LeaderEnd;
                        roomTag.HasLeader = false;

                        //Create tag room wall
                        RoomTag roomTag1 = doc.Create.NewRoomTag(new LinkElementId(room.Id), new UV(0, 0), doc.ActiveView.Id);
                        roomTag1.ChangeTypeId((view.cbtagWall.SelectedItem as RoomTagType).Id);
                        roomTag1.HasLeader = true;
                        roomTag1.LeaderEnd = new XYZ(roomTag1.LeaderEnd.X, roomTag1.LeaderEnd.Y - 2, roomTag1.LeaderEnd.Z + 1);
                        roomTag1.HasLeader = false;

                        //create tag ceiling
                        RoomTag roomTag2 = doc.Create.NewRoomTag(new LinkElementId(room.Id), new UV(0, 0), doc.ActiveView.Id);
                        roomTag2.ChangeTypeId((view.cbtagCeil.SelectedItem as RoomTagType).Id);
                        roomTag2.HasLeader = true;
                        roomTag2.LeaderEnd = new XYZ(roomTag2.LeaderEnd.X, roomTag2.LeaderEnd.Y + 1, loca.Max.Z);
                        roomTag2.TagHeadPosition = new XYZ(roomTag2.LeaderEnd.X, roomTag2.LeaderEnd.Y - 2, loca.Max.Z);
                        roomTag2.LeaderElbow = new XYZ(roomTag2.LeaderEnd.X, roomTag2.LeaderEnd.Y - 0.5, loca.Max.Z - 0.8566820589069);

                        //create tag floor
                        RoomTag roomTag3 = doc.Create.NewRoomTag(new LinkElementId(room.Id), new UV(0, 0), doc.ActiveView.Id);
                        roomTag3.ChangeTypeId((view.cbtagFloor.SelectedItem as RoomTagType).Id);
                        roomTag3.HasLeader = true;
                        roomTag3.LeaderEnd = new XYZ(roomTag3.LeaderEnd.X, roomTag3.LeaderEnd.Y - 1, loca.Min.Z);
                        roomTag3.TagHeadPosition = new XYZ(roomTag3.LeaderEnd.X, roomTag3.LeaderEnd.Y, loca.Min.Z + 3.5);
                        roomTag3.LeaderElbow = new XYZ(roomTag3.LeaderEnd.X, roomTag3.LeaderEnd.Y + 0.5, loca.Min.Z + 0.81757808238097);
                        #endregion
                        //create tag habaki
                        RoomTag roomTag4 = doc.Create.NewRoomTag(new LinkElementId(room.Id), new UV(0, 0), doc.ActiveView.Id);
                        roomTag4.ChangeTypeId((view.cbtagHabaki.SelectedItem as RoomTagType).Id);
                        roomTag4.HasLeader = true;
                        roomTag4.LeaderEnd = new XYZ(loca.Min.X, loca.Min.Y, loca.Min.Z);
                        roomTag4.TagHeadPosition = new XYZ(loca.Min.X, loca.Min.Y, loca.Min.Z + 3);
                        roomTag4.LeaderElbow = new XYZ(loca.Min.X, loca.Min.Y + 0.5, loca.Min.Z + 0.683812058446638);

                        RoomTag roomTag5 = doc.Create.NewRoomTag(new LinkElementId(room.Id), new UV(0, 0), doc.ActiveView.Id);
                        roomTag5.ChangeTypeId((view.cbtagHabaki.SelectedItem as RoomTagType).Id);
                        roomTag5.HasLeader = true;
                        roomTag5.LeaderEnd = new XYZ(loca.Max.X, loca.Max.Y, loca.Min.Z);
                        roomTag5.TagHeadPosition = new XYZ(loca.Max.X, loca.Max.Y - 5, loca.Min.Z + 3);
                        roomTag5.LeaderElbow = new XYZ(loca.Max.X, loca.Max.Y - 0.5, loca.Min.Z + 0.683812058446638);


                        //create tag ...
                        RoomTag roomTag6 = doc.Create.NewRoomTag(new LinkElementId(room.Id), new UV(0, 0), doc.ActiveView.Id);
                        roomTag6.ChangeTypeId((view.cbtagMawari.SelectedItem as RoomTagType).Id);
                        roomTag6.HasLeader = true;
                        roomTag6.LeaderEnd = new XYZ(loca.Min.X, loca.Max.Y, loca.Max.Z);
                        roomTag6.TagHeadPosition = new XYZ(loca.Min.X, loca.Max.Y - 5, loca.Max.Z + 1.3);
                        roomTag6.LeaderElbow = new XYZ(loca.Min.X, loca.Max.Y - 0.5, loca.Max.Z - 1.1311838619316);

                        RoomTag roomTag7 = doc.Create.NewRoomTag(new LinkElementId(room.Id), new UV(0, 0), doc.ActiveView.Id);
                        roomTag7.ChangeTypeId((view.cbtagMawari.SelectedItem as RoomTagType).Id);
                        roomTag7.HasLeader = true;
                        roomTag7.LeaderEnd = new XYZ(loca.Max.X, loca.Min.Y, loca.Max.Z);
                        roomTag7.TagHeadPosition = new XYZ(loca.Max.X - 5, loca.Min.Y, loca.Max.Z + 1.3);
                        roomTag7.LeaderElbow = new XYZ(loca.Max.X - 0.5, loca.Min.Y + 0.5, loca.Max.Z - 1.1311838619316);

                        view.Close();
                        tran.Commit();
                    }
                }
            }
        }
        public XYZ GetRoomCenter(Room room)
        {
            XYZ boundCenter = GetElementCenter(room);
            LocationPoint locPt = (LocationPoint)room.Location;
            XYZ roomCenter = new XYZ(boundCenter.X, boundCenter.Y, locPt.Point.Z);
            return roomCenter;
        }
        public XYZ GetElementCenter(Element elem)
        {
            BoundingBoxXYZ bounding = elem.get_BoundingBox(doc.ActiveView);
            XYZ center = (bounding.Max + bounding.Min) * 0.5;
            return center;
        }
        void SetItemSource(List<RoomTagType> source, bool b)
        {
            if (b)
            {
                SrcroomName = new ObservableCollection<RoomTagType>(source.Where(p => p.FamilyName == "dタグ_部屋_一般").Cast<RoomTagType>().ToList());

                SrcroomWall = new ObservableCollection<RoomTagType>(source.Where(p => p.FamilyName == "dタグ_部屋_壁").Cast<RoomTagType>().ToList());

                SrcroomCeiling = new ObservableCollection<RoomTagType>(source.Where(p => p.FamilyName == "dタグ_部屋_天井").Cast<RoomTagType>().ToList());

                SrcroomFloor = new ObservableCollection<RoomTagType>(source.Where(p => p.FamilyName == "dタグ_部屋_床").Cast<RoomTagType>().ToList());

                SrcroomHabaki = new ObservableCollection<RoomTagType>(source.Where(p => p.FamilyName == "dタグ_部屋_幅木").Cast<RoomTagType>().ToList());

                SrcroomCeilingConnect = new ObservableCollection<RoomTagType>(source.Where(p => p.FamilyName == "dタグ_部屋_見切縁").Cast<RoomTagType>().ToList());
            }
            else
            {
                SrcroomName = new ObservableCollection<RoomTagType>(source.Where(p => p.FamilyName == "dタグ_部屋_一般").Cast<RoomTagType>().ToList());

                SrcroomWall = new ObservableCollection<RoomTagType>(source.Where(p => p.FamilyName == "dタグ_部屋_壁_Truss対応").Cast<RoomTagType>().ToList());

                SrcroomCeiling = new ObservableCollection<RoomTagType>(source.Where(p => p.FamilyName == "dタグ_部屋_天井_Truss対応").Cast<RoomTagType>().ToList());

                SrcroomFloor = new ObservableCollection<RoomTagType>(source.Where(p => p.FamilyName == "dタグ_部屋_床_Truss対応").Cast<RoomTagType>().ToList());

                SrcroomHabaki = new ObservableCollection<RoomTagType>(source.Where(p => p.FamilyName == "dタグ_部屋_幅木_Truss対応").Cast<RoomTagType>().ToList());

                SrcroomCeilingConnect = new ObservableCollection<RoomTagType>(source.Where(p => p.FamilyName == "dタグ_部屋_廻り縁_Truss対応").Cast<RoomTagType>().ToList());
            }
        }
        void SetSelectedIndex()
        {
            view.cbtagCeil.SelectedIndex = 0;
            view.cbtagFloor.SelectedIndex = 0;
            view.cbtagRoom.SelectedIndex = 0;
            view.cbtagMawari.SelectedIndex = 0;
            view.cbtagWall.SelectedIndex = 0;
            view.cbtagHabaki.SelectedIndex = 0;

        }
        void ClearSource()
        {
            SrcroomName.Clear();
            SrcroomWall.Clear();
            SrcroomCeiling.Clear();
            SrcroomFloor.Clear();
            SrcroomCeilingConnect.Clear();
        }
        #endregion
    }
}
