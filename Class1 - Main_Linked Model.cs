using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;

[Transaction(TransactionMode.Manual)]
public class TagAllRooms : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uiDoc = commandData.Application.ActiveUIDocument;
        Document doc = uiDoc.Document;

        try
        {
            using (Transaction trans = new Transaction(doc, "Tag All Rooms Including Linked Models"))
            {
                trans.Start();

                int tagCount = 0;

                // Get all floor plan views (non-template)
                var viewPlans = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewPlan))
                    .Cast<ViewPlan>()
                    .Where(v => v.ViewType == ViewType.FloorPlan && !v.IsTemplate)
                    .ToList();

                // Map LevelId to a corresponding ViewPlan
                var levelViewMap = viewPlans
                    .GroupBy(v => v.GenLevel.Id)
                    .ToDictionary(g => g.Key, g => g.First());

                // Get room tag symbol
                var roomTagType = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilySymbol))
                    .OfCategory(BuiltInCategory.OST_RoomTags)
                    .Cast<FamilySymbol>()
                    .FirstOrDefault();

                if (roomTagType == null)
                {
                    TaskDialog.Show("Error", "No room tag family loaded.");
                    return Result.Failed;
                }

                if (!roomTagType.IsActive)
                {
                    roomTagType.Activate();
                    doc.Regenerate();
                }

                //// Use a local function instead of a lambda for tagging a room
                //void TagRoom(Room room, XYZ point, ElementId levelId)
                //{
                //    if (levelViewMap.TryGetValue(levelId, out ViewPlan view))
                //    {
                //        UV uv = new UV(point.X, point.Y);
                //        doc.Create.NewRoomTag(new LinkElementId(room.Id), uv, view.Id);
                //        tagCount++;
                //    }
                //}

                // Helper method: tag a room
                Action<Room, XYZ, ElementId> TagRoom = (room, point, levelId) =>
                {
                    if (levelViewMap.TryGetValue(levelId, out ViewPlan view))
                    {
                        UV uv = new UV(point.X, point.Y);
                        doc.Create.NewRoomTag(new LinkElementId(room.Id), uv, view.Id);
                        tagCount++;
                    }
                };

                // ➤ Step 1: Tag rooms in the host model
                var hostRooms = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Rooms)
                    .WhereElementIsNotElementType()
                    .Cast<Room>()
                    .Where(r => r.Area > 0 && r.Location is LocationPoint)
                    .ToList();

                foreach (Room room in hostRooms)
                {
                    XYZ point = (room.Location as LocationPoint).Point;
                    TagRoom(room, point, room.LevelId);
                }

                // ➤ Step 2: Tag rooms in linked models
                var linkInstances = new FilteredElementCollector(doc)
                    .OfClass(typeof(RevitLinkInstance))
                    .Cast<RevitLinkInstance>()
                    .ToList();

                foreach (RevitLinkInstance linkInstance in linkInstances)
                {
                    Document linkedDoc = linkInstance.GetLinkDocument();
                    if (linkedDoc == null || !linkedDoc.IsLinked) continue;

                    Transform linkTransform = linkInstance.GetTransform();

                    var linkedRooms = new FilteredElementCollector(linkedDoc)
                        .OfCategory(BuiltInCategory.OST_Rooms)
                        .WhereElementIsNotElementType()
                        .Cast<Room>()
                        .Where(r => r.Area > 0 && r.Location is LocationPoint)
                        .ToList();

                    foreach (Room room in linkedRooms)
                    {
                        XYZ linkedPoint = (room.Location as LocationPoint).Point;
                        XYZ hostPoint = linkTransform.OfPoint(linkedPoint);
                        TagRoom(room, hostPoint, room.LevelId); // Tag in corresponding level in host
                    }
                }

                trans.Commit();
                TaskDialog.Show("Room Tagger", $"Room tags placed: {tagCount}");
                return Result.Succeeded;
            }
        }
        catch (Exception ex)
        {
            message = ex.Message;
            return Result.Failed;
        }
    }
}